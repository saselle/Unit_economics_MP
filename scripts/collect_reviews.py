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
from common import (  # noqa: E402
    basket_candidates,
    basket_parts,
    curl_requests,
    ensure_parent,
    load_config,
    make_session,
    resolve,
)

# WB регулярно меняет версию карточного API, поэтому пробуем несколько.
CARD_DETAIL_URLS = [
    "https://card.wb.ru/cards/v4/detail",
    "https://card.wb.ru/cards/v3/detail",
    "https://card.wb.ru/cards/v2/detail",
    "https://card.wb.ru/cards/v1/detail",
    "https://u-card.wb.ru/cards/v4/detail",
    "https://u-card.wb.ru/cards/v2/detail",
]
# ВАЖНО: v2 ждёт id карточки-родителя (imt_id), а v1 — артикул (nm_id).
# Перепутать их нельзя: WB не ругается, а отдаёт отзывы постороннего товара,
# у которого артикул совпал с нашим imt_id. Поэтому у каждого адреса указано,
# какой идентификатор подставлять, и ответ дополнительно проверяется.
FEEDBACK_URLS = [
    ("https://feedbacks1.wb.ru/feedbacks/v2/{id}", "imt"),
    ("https://feedbacks2.wb.ru/feedbacks/v2/{id}", "imt"),
    ("https://feedbacks1.wb.ru/feedbacks/v1/{id}", "nm"),
    ("https://feedbacks2.wb.ru/feedbacks/v1/{id}", "nm"),
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
    "color", "size", "useful_votes",
    # поля из самого отзыва — по ним видно, к какому товару он относится
    "review_nm_id", "review_brand", "review_product", "collected_at",
]


def imt_from_card_api(session, nm_id: str, dest: int, timeout: float) -> str | None:
    """Способ 1: карточный API. Отдаёт imt_id в поле root."""
    params = {"appType": 1, "curr": "rub", "dest": dest, "spp": 30,
              "hide_dtype": 13, "lang": "ru", "nm": nm_id}
    for url in CARD_DETAIL_URLS:
        try:
            response = session.get(url, params=params, headers=HEADERS, timeout=timeout)
            if response.status_code != 200:
                continue
            products = ((response.json() or {}).get("data") or {}).get("products") or []
            if products:
                imt = products[0].get("root") or products[0].get("imtId")
                if imt:
                    return str(imt)
        except Exception:
            continue
    return None


def imt_from_basket(session, nm_id: str, timeout: float, hint: int | None) -> tuple[str | None, int | None]:
    """Способ 2: card.json в корзине. Возвращает (imt_id, номер сработавшей корзины).

    Номер корзины считается по таблице диапазонов, но WB её иногда меняет,
    поэтому при промахе просто перебираем остальные. Удачный номер запоминаем
    и пробуем первым для следующих товаров — обычно соседние артикулы лежат рядом.
    """
    try:
        vol, part = basket_parts(nm_id)
    except (TypeError, ValueError):
        return None, hint

    for attempt, number in enumerate(basket_candidates(nm_id, hint), start=1):
        url = (f"https://basket-{number:02d}.wbbasket.ru"
               f"/vol{vol}/part{part}/{nm_id}/info/ru/card.json")
        try:
            response = session.get(url, headers=HEADERS, timeout=timeout)
            if response.status_code != 200:
                continue
            imt = (response.json() or {}).get("imt_id")
            if imt:
                if attempt > 1:
                    print(f"    товар нашёлся в basket-{number:02d} (проб: {attempt})")
                return str(imt), number
        except Exception:
            continue
    return None, hint


def get_imt_id(session, nm_id: str, dest: int, timeout: float,
               hint: int | None = None) -> tuple[str | None, int | None]:
    """Артикул -> id карточки-родителя, по которому лежат отзывы."""
    imt = imt_from_card_api(session, nm_id, dest, timeout)
    if imt:
        return imt, hint
    if hint is None:
        print("    карточный API не отвечает, ищу товар по корзинам WB "
              "(первый товар — до минуты, дальше быстро)")
    return imt_from_basket(session, nm_id, timeout, hint)


def feedback_belongs(feedback: dict, nm_id: str, imt_id: str) -> bool | None:
    """Относится ли отзыв к нашему товару.

    True/False — если в отзыве есть данные о товаре, None — если проверить нечем.
    """
    details = feedback.get("productDetails")
    if not isinstance(details, dict):
        return None
    for key, expected in (("nmId", nm_id), ("imtId", imt_id)):
        value = details.get(key)
        if value and str(value) == str(expected):
            return True
    return False


def fetch_feedbacks(session, imt_id: str, nm_id: str, timeout: float) -> list[dict]:
    """Скачивает отзывы нашего товара. Чужие ответы отбрасываются целиком."""
    for template, id_kind in FEEDBACK_URLS:
        url = template.format(id=imt_id if id_kind == "imt" else nm_id)
        try:
            response = session.get(url, headers=HEADERS, timeout=timeout)
            if response.status_code != 200:
                continue
            feedbacks = (response.json() or {}).get("feedbacks")
            if not feedbacks:
                continue

            checks = [feedback_belongs(f, nm_id, imt_id) for f in feedbacks]
            ours = [f for f, ok in zip(feedbacks, checks) if ok]
            if ours:
                if len(ours) < len(feedbacks):
                    print(f"    отброшено чужих отзывов: {len(feedbacks) - len(ours)}")
                return ours
            if any(ok is False for ok in checks):
                # ответ есть, но он про другой товар — адрес не подходит
                print(f"    {url.rsplit('/', 2)[-2]} вернул отзывы другого товара, пропускаю")
                continue
            # проверить нечем (старый формат ответа) — доверяем только v2 по imt_id
            if id_kind == "imt":
                return feedbacks
        except Exception:
            continue
    return []


def normalize(feedback: dict, product: dict, imt_id: str, collected_at: str) -> dict:
    """Приводит отзыв к плоской строке. Чего нет — оставляем пустым.

    Имя автора отзыва намеренно не сохраняем: для анализа оно не нужно.
    """
    votes = feedback.get("votes")
    details = feedback.get("productDetails") if isinstance(feedback.get("productDetails"), dict) else {}
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
        "review_nm_id": details.get("nmId", ""),
        "review_brand": (details.get("brandName") or "").strip(),
        "review_product": (details.get("productName") or "").strip(),
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
    basket_hint: int | None = None  # номер корзины, сработавший на прошлом товаре
    seen_imt: dict[str, str] = {}   # imt_id -> артикул, по которому уже собрали

    for i, product in enumerate(products, start=1):
        nm = product["product_id"]
        title = (product.get("title") or "")[:48]
        print(f"[{i}/{len(products)}] {nm} {title}")

        imt_id, basket_hint = get_imt_id(session, nm, dest, timeout, basket_hint)
        if not imt_id:
            print("    не удалось узнать id карточки, пропускаю")
            time.sleep(delay)
            continue

        if imt_id in seen_imt:
            # один товар на WB часто продаётся несколькими артикулами с общей
            # родительской карточкой — отзывы у них одни и те же
            print(f"    та же карточка, что у артикула {seen_imt[imt_id]}, пропускаю")
            time.sleep(delay)
            continue
        seen_imt[imt_id] = nm

        feedbacks = fetch_feedbacks(session, imt_id, nm, timeout)
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
            print("Если поиск работает, а отзывы нет — пришлите этот экран: "
                  "значит, WB снова сменил адреса карточек или отзывов.")
        sys.exit(1)

    write_reviews(reviews_path, rows)
    ensure_parent(raw_path)
    raw_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\nГотово: {len(rows)} отзывов по {len(raw)} карточкам -> {reviews_path}")
    print(f"Сырые данные -> {raw_path}")
    print("Дальше: py scripts\\analyze_reviews.py")


if __name__ == "__main__":
    main()
