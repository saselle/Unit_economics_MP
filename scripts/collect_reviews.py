"""Сбор отзывов Wildberries по самым сильным карточкам из выдачи.

Как это работает:
1. берём data/products.csv и выбираем топ-N карточек по числу отзывов;
2. для каждой узнаём «корневой» id товара (imt_id) — отзывы на WB хранятся
   не по артикулу, а по карточке-родителю, общей для всех цветов и размеров;
3. скачиваем отзывы и складываем в data/reviews.csv и data/raw_wb_reviews.json.

Запуск:
    python scripts/collect_reviews.py
    python scripts/collect_reviews.py --top 30        # больше карточек
    python scripts/collect_reviews.py --nm 123456789  # только один артикул
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import curl_requests, ensure_parent, load_config, make_session, resolve  # noqa: E402

CARD_DETAIL_URL = "https://card.wb.ru/cards/v2/detail"
FEEDBACK_URLS = [
    "https://feedbacks1.wb.ru/feedbacks/v2/{imt}",
    "https://feedbacks2.wb.ru/feedbacks/v2/{imt}",
    "https://feedbacks1.wb.ru/feedbacks/v1/{imt}",
    "https://feedbacks2.wb.ru/feedbacks/v1/{imt}",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "*/*",
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Referer": "https://www.wildberries.ru/",
}

REVIEW_COLUMNS = [
    "marketplace", "product_id", "imt_id", "brand", "product_title",
    "review_id", "rating", "created_at", "text", "pros", "cons",
    "color", "size", "useful_votes", "collected_at",
]


def get_imt_id(session, nm_id: str, dest: int, timeout: float) -> str | None:
    """Артикул -> id карточки-родителя, по которому лежат отзывы."""
    params = {"appType": 1, "curr": "rub", "dest": dest, "spp": 30,
              "hide_dtype": 13, "lang": "ru", "nm": nm_id}
    try:
        response = session.get(CARD_DETAIL_URL, params=params, headers=HEADERS, timeout=timeout)
        if response.status_code != 200:
            print(f"    карточка {nm_id}: ответ {response.status_code}")
            return None
        products = ((response.json() or {}).get("data") or {}).get("products") or []
        if not products:
            return None
        return str(products[0].get("root") or products[0].get("imtId") or "") or None
    except Exception as exc:
        print(f"    карточка {nm_id}: {type(exc).__name__}: {str(exc)[:80]}")
        return None


def fetch_feedbacks(session, imt_id: str, timeout: float) -> list[dict]:
    """Скачивает отзывы по imt_id. Пустой список, если ни один адрес не ответил."""
    for template in FEEDBACK_URLS:
        url = template.format(imt=imt_id)
        try:
            response = session.get(url, headers=HEADERS, timeout=timeout)
            if response.status_code != 200:
                continue
            payload = response.json() or {}
            feedbacks = payload.get("feedbacks")
            if feedbacks:
                return feedbacks
        except Exception:
            continue
    return []


def normalize(feedback: dict, product: dict, imt_id: str, collected_at: str) -> dict:
    """Приводит отзыв к плоской строке. Чего нет — оставляем пустым.

    Имя автора отзыва намеренно не сохраняем: для анализа оно не нужно.
    """
    votes = feedback.get("votes")
    return {
        "marketplace": "wildberries",
        "product_id": product.get("product_id", ""),
        "imt_id": imt_id,
        "brand": product.get("brand", ""),
        "product_title": product.get("title", ""),
        "review_id": feedback.get("id") or feedback.get("globalUserId") or "",
        "rating": feedback.get("productValuation") or feedback.get("valuation") or "",
        "created_at": (feedback.get("createdDate") or "")[:19].replace("T", " "),
        "text": (feedback.get("text") or "").replace("\n", " ").strip(),
        "pros": (feedback.get("pros") or "").replace("\n", " ").strip(),
        "cons": (feedback.get("cons") or "").replace("\n", " ").strip(),
        "color": (feedback.get("color") or "").strip(),
        "size": (feedback.get("size") or "").strip(),
        "useful_votes": votes.get("pluses", "") if isinstance(votes, dict) else "",
        "collected_at": collected_at,
    }


def pick_products(df: pd.DataFrame, top: int, only_nm: str | None) -> list[dict]:
    """Выбирает карточки для сбора: самые «отзывные», без дублей по артикулу."""
    df = df.copy()
    df["review_count"] = pd.to_numeric(df.get("review_count"), errors="coerce").fillna(0)
    df = df.drop_duplicates("product_id")

    if only_nm:
        df = df[df["product_id"].astype(str) == str(only_nm)]
        if df.empty:
            return [{"product_id": str(only_nm), "brand": "", "title": ""}]
    else:
        df = df.sort_values("review_count", ascending=False).head(top)

    return [{"product_id": str(r.get("product_id", "")),
             "brand": str(r.get("brand", "")),
             "title": str(r.get("title", "")),
             "review_count": int(r.get("review_count", 0))}
            for _, r in df.iterrows()]


def write_reviews(path: Path, rows: list[dict]) -> None:
    ensure_parent(path)
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=REVIEW_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({col: row.get(col, "") for col in REVIEW_COLUMNS})


def main() -> None:
    parser = argparse.ArgumentParser(description="Сбор отзывов Wildberries")
    parser.add_argument("--config", default="config.yaml")
    parser.add_argument("--top", type=int, default=None, help="сколько карточек обойти")
    parser.add_argument("--nm", default=None, help="собрать отзывы только по этому артикулу")
    args = parser.parse_args()

    cfg = load_config(args.config)
    out_cfg = cfg.get("output") or {}
    rev_cfg = cfg.get("reviews") or {}
    wb_cfg = (cfg.get("marketplaces") or {}).get("wildberries") or {}

    products_path = resolve(out_cfg.get("products_file", "data/products.csv"))
    reviews_path = resolve(out_cfg.get("reviews_file", "data/reviews.csv"))
    raw_path = resolve(out_cfg.get("raw_reviews_file", "data/raw_wb_reviews.json"))

    top = args.top or int(rev_cfg.get("top_products", 15))
    max_per_product = int(rev_cfg.get("max_per_product", 300))
    delay = float(rev_cfg.get("request_delay_sec", 1.5))
    timeout = float(wb_cfg.get("timeout_sec", 20))
    dest = int(wb_cfg.get("dest", -1257786))

    if not products_path.exists():
        sys.exit(f"Нет файла {products_path}. Сначала запустите scripts/collect_wb.py")

    df = pd.read_csv(products_path, dtype=str, keep_default_na=False)
    products = pick_products(df, top, args.nm)
    if not products:
        sys.exit("Не нашёл карточек для сбора отзывов.")

    session, engine = make_session()
    print(f"Движок запросов: {engine}")
    print(f"Карточек к обходу: {len(products)}, максимум {max_per_product} отзывов с каждой\n")

    collected_at = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    rows: list[dict] = []
    raw: dict[str, list] = {}

    for i, product in enumerate(products, start=1):
        nm = product["product_id"]
        title = (product.get("title") or "")[:48]
        print(f"[{i}/{len(products)}] {nm} {title}")

        imt_id = get_imt_id(session, nm, dest, timeout)
        if not imt_id:
            print("    не удалось узнать id карточки, пропускаю")
            time.sleep(delay)
            continue

        feedbacks = fetch_feedbacks(session, imt_id, timeout)
        if not feedbacks:
            print(f"    отзывов не получено (imt {imt_id})")
            time.sleep(delay)
            continue

        feedbacks = feedbacks[:max_per_product]
        raw[nm] = feedbacks
        for feedback in feedbacks:
            try:
                rows.append(normalize(feedback, product, imt_id, collected_at))
            except Exception as exc:
                print(f"    пропускаю отзыв: {type(exc).__name__}: {exc}")
        print(f"    собрано отзывов: {len(feedbacks)}")
        time.sleep(delay)

    if not rows:
        print("\nНи одного отзыва собрать не удалось.")
        if curl_requests is None:
            print("Поставьте маскировку под Chrome:  py -m pip install curl_cffi")
        else:
            print("Проверьте доступ к WB командой:  py scripts\\check_wb.py")
        return

    write_reviews(reviews_path, rows)
    ensure_parent(raw_path)
    raw_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\nГотово: {len(rows)} отзывов по {len(raw)} карточкам -> {reviews_path}")
    print(f"Сырые данные -> {raw_path}")
    print("Дальше: py scripts\\analyze_reviews.py")


if __name__ == "__main__":
    main()
