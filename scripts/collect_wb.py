"""Сбор карточек Wildberries по списку поисковых запросов.

Что делает:
1. читает запросы из data/queries.csv;
2. по каждому запросу дёргает публичный поисковый endpoint WB;
3. складывает сырые ответы в data/raw_wb_products.json;
4. пишет нормализованную таблицу в data/products.csv.

Если WB недоступен (блокировка, смена endpoint, нет интернета) — скрипт не падает,
а переключается на fallback с мок-данными (data/sample_products.csv),
чтобы анализ и отчёты можно было прогнать и проверить.

Запуск:
    python scripts/collect_wb.py
    python scripts/collect_wb.py --sample     # принудительно мок-данные
    python scripts/collect_wb.py --top-n 30   # переопределить top_n из конфига
"""

from __future__ import annotations

import argparse
import csv
import json
import random
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import (  # noqa: E402
    PRODUCT_COLUMNS,
    curl_requests,
    ensure_parent,
    load_config,
    make_session,
    read_queries,
    resolve,
)

MARKETPLACE = "wildberries"

# Несколько версий endpoint: WB периодически переключает версии, пробуем по очереди.
# Если все отдают 403/404 — запустите scripts/check_wb.py, он подберёт рабочий вариант.
SEARCH_ENDPOINTS = [
    "https://search.wb.ru/exactmatch/ru/common/v18/search",
    "https://search.wb.ru/exactmatch/ru/common/v17/search",
    "https://search.wb.ru/exactmatch/ru/common/v14/search",
    "https://search.wb.ru/exactmatch/ru/common/v13/search",
    "https://search.wb.ru/exactmatch/ru/common/v5/search",
    "https://search.wb.ru/exactmatch/ru/common/v4/search",
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

# Диапазоны корзин для картинок WB. Мапинг иногда меняется — если картинки
# перестали открываться, обновите таблицу (или просто игнорируйте поле image_url).
BASKET_RANGES = [
    (143, "01"), (287, "02"), (431, "03"), (719, "04"), (1007, "05"),
    (1061, "06"), (1115, "07"), (1169, "08"), (1313, "09"), (1601, "10"),
    (1655, "11"), (1919, "12"), (2045, "13"), (2189, "14"), (2405, "15"),
    (2621, "16"), (2837, "17"), (3053, "18"), (3269, "19"), (3485, "20"),
    (3701, "21"), (3917, "22"), (4133, "23"), (4349, "24"), (4565, "25"),
]


# --------------------------------------------------------------------------- #
#  Работа с WB
# --------------------------------------------------------------------------- #
def image_url(nm_id: int) -> str:
    """Собирает ссылку на главное фото карточки по её артикулу."""
    try:
        vol = int(nm_id) // 100_000
        part = int(nm_id) // 1_000
    except (TypeError, ValueError):
        return ""
    host = "26"
    for upper, candidate in BASKET_RANGES:
        if vol <= upper:
            host = candidate
            break
    return f"https://basket-{host}.wbbasket.ru/vol{vol}/part{part}/{nm_id}/images/big/1.webp"


def kopecks(value) -> str:
    """WB отдаёт цены в копейках. Возвращаем рубли строкой, пустую — если нет данных."""
    try:
        rub = round(float(value) / 100, 2)
    except (TypeError, ValueError):
        return ""
    if rub <= 0:
        return ""
    return f"{rub:.2f}"


def extract_products(payload: dict) -> list[dict]:
    """Достаёт список товаров из ответа: структура отличается между версиями API."""
    if not isinstance(payload, dict):
        return []
    data = payload.get("data")
    if isinstance(data, dict) and isinstance(data.get("products"), list):
        return data["products"]
    if isinstance(payload.get("products"), list):
        return payload["products"]
    return []


def prices_from_product(item: dict) -> tuple[str, str]:
    """Возвращает (актуальная цена, цена до скидки) в рублях.

    Старое API: priceU / salePriceU в корне товара.
    Новое API: sizes[].price.{basic,product,total}.
    """
    price = kopecks(item.get("salePriceU"))
    old_price = kopecks(item.get("priceU"))

    if not price:
        sizes = item.get("sizes") or []
        for size in sizes:
            size_price = (size or {}).get("price") or {}
            price = kopecks(size_price.get("product") or size_price.get("total"))
            old_price = old_price or kopecks(size_price.get("basic"))
            if price:
                break
    return price, old_price


def normalize(item: dict, query: str, rank: int, collected_at: str) -> dict:
    """Приводит карточку WB к нашей плоской схеме. Недостающие поля -> пустая строка."""
    nm_id = item.get("id") or item.get("nmId") or ""
    price, old_price = prices_from_product(item)

    discount = ""
    try:
        if price and old_price and float(old_price) > 0:
            discount = str(round((1 - float(price) / float(old_price)) * 100))
    except (TypeError, ValueError):
        discount = ""
    if not discount and item.get("sale") not in (None, ""):
        discount = str(item.get("sale"))

    rating = item.get("reviewRating") or item.get("nmReviewRating") or item.get("rating") or ""
    reviews = item.get("feedbacks") or item.get("nmFeedbacks") or ""

    return {
        "marketplace": MARKETPLACE,
        "query": query,
        "rank": rank,
        "product_id": nm_id,
        "product_url": f"https://www.wildberries.ru/catalog/{nm_id}/detail.aspx" if nm_id else "",
        "title": (item.get("name") or item.get("imt_name") or "").strip(),
        "brand": (item.get("brand") or "").strip(),
        "price": price,
        "old_price": old_price,
        "discount": discount,
        "rating": rating if rating not in (0, "0") else "",
        "review_count": reviews if reviews not in (None, "") else "",
        "image_url": image_url(nm_id) if nm_id else "",
        "collected_at": collected_at,
    }


def fetch_query(session, query: str, wb_cfg: dict) -> dict | None:
    """Пробует получить выдачу по запросу. Возвращает сырой JSON или None."""
    params = {
        "ab_testing": "false",
        "appType": 1,
        "curr": "rub",
        "dest": wb_cfg.get("dest", -1257786),
        "hide_dtype": 13,
        "lang": "ru",
        "query": query,
        "resultset": "catalog",
        "sort": "popular",
        "spp": 30,
        "suppressSpellcheck": "false",
        "page": 1,
    }
    retries = int(wb_cfg.get("retries", 3))
    timeout = float(wb_cfg.get("timeout_sec", 20))

    for endpoint in SEARCH_ENDPOINTS:
        for attempt in range(1, retries + 1):
            try:
                response = session.get(endpoint, params=params, headers=HEADERS, timeout=timeout)
                response.raise_for_status()
                payload = response.json()
                if extract_products(payload):
                    return payload
                print(f"    пустая выдача: {endpoint} (попытка {attempt})")
            except Exception as exc:  # сеть, таймаут, битый JSON, 4xx/5xx
                reason = str(exc).split("(Caused by")[0].strip()[:120]
                print(f"    ошибка {endpoint.rsplit('/common/', 1)[-1]} "
                      f"(попытка {attempt}): {type(exc).__name__}: {reason}")
            if attempt < retries:
                time.sleep(min(2 ** attempt, 8))
    return None


def collect_live(queries: list[str], wb_cfg: dict) -> tuple[list[dict], dict]:
    """Собирает карточки по всем запросам. Возвращает (строки, сырые ответы)."""
    top_n = int(wb_cfg.get("top_n", 20))
    delay = float(wb_cfg.get("request_delay_sec", 1.0))
    collected_at = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")

    rows: list[dict] = []
    raw: dict[str, dict] = {}
    session, engine = make_session()
    print(f"Движок запросов: {engine}")

    # если WB недоступен совсем (нет сети / блокировка), нет смысла ждать все 20 запросов
    max_fails_in_row = int(wb_cfg.get("abort_after_failed_queries", 3))
    fails_in_row = 0

    for i, query in enumerate(queries, start=1):
        print(f"[{i}/{len(queries)}] {query}")
        payload = fetch_query(session, query, wb_cfg)
        if payload is None:
            fails_in_row += 1
            print("    -> не удалось получить данные, идём дальше")
            if not rows and fails_in_row >= max_fails_in_row:
                print(f"    -> {fails_in_row} запроса подряд без ответа, прекращаю живой сбор")
                break
            continue
        fails_in_row = 0

        products = extract_products(payload)[:top_n]
        raw[query] = {"collected_at": collected_at, "products": products}
        for rank, item in enumerate(products, start=1):
            try:
                rows.append(normalize(item, query, rank, collected_at))
            except Exception as exc:  # одна кривая карточка не должна ронять сбор
                print(f"    пропускаем карточку #{rank}: {type(exc).__name__}: {exc}")
        print(f"    -> собрано {len(products)} карточек")
        time.sleep(delay)

    return rows, raw


# --------------------------------------------------------------------------- #
#  Fallback: мок-данные
# --------------------------------------------------------------------------- #
SAMPLE_BRANDS = [
    "Cleanelly", "Verossa", "Belezza", "Sofi De Marko", "Karna", "Hobby Home",
    "Ecotex", "Bio-Textiles", "Волжский текстиль", "Донецкая мануфактура",
    "Aquarelle", "Togas", "Arya Home", "Гармония", "Unison",
]
SAMPLE_PARTS = {
    "material": ["махровое", "хлопок", "бамбук", "органический хлопок", "микрофибра"],
    "size": ["50х90", "70х140", "100х150", "30х50", "80х150"],
    "kind": ["банное", "для лица", "для ванной", "для рук", "пляжное"],
    "pack": ["набор 2 шт", "набор 4 шт", "комплект 3 шт", "подарочный набор", ""],
    "color": ["бежевое", "белое", "серое", "молочное", "графит"],
}


def build_sample_rows(queries: list[str], top_n: int) -> list[dict]:
    """Генерит правдоподобные мок-карточки, чтобы прогнать анализ без сети."""
    rng = random.Random(42)
    collected_at = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    rows: list[dict] = []

    for query in queries:
        for rank in range(1, top_n + 1):
            brand = rng.choice(SAMPLE_BRANDS)
            parts = [
                "Полотенце",
                rng.choice(SAMPLE_PARTS["material"]),
                rng.choice(SAMPLE_PARTS["kind"]),
                rng.choice(SAMPLE_PARTS["size"]),
                rng.choice(SAMPLE_PARTS["pack"]),
                rng.choice(SAMPLE_PARTS["color"]),
            ]
            if "набор" in query or "комплект" in query:
                parts[4] = rng.choice(["набор 4 шт", "комплект 3 шт", "подарочный набор"])
            title = " ".join(p for p in parts if p)

            base = rng.choice([390, 590, 790, 990, 1290, 1690, 2390])
            if "набор" in query or "комплект" in query or "подарочн" in query:
                base = int(base * rng.uniform(1.6, 2.4))
            price = round(base * rng.uniform(0.85, 1.15), 2)
            old_price = round(price * rng.uniform(1.3, 2.2), 2)
            product_id = 100_000_00 + rng.randint(0, 9_000_000)

            rows.append({
                "marketplace": MARKETPLACE,
                "query": query,
                "rank": rank,
                "product_id": product_id,
                "product_url": f"https://www.wildberries.ru/catalog/{product_id}/detail.aspx",
                "title": title,
                "brand": brand,
                "price": f"{price:.2f}",
                "old_price": f"{old_price:.2f}",
                "discount": str(round((1 - price / old_price) * 100)),
                "rating": f"{rng.uniform(4.2, 5.0):.1f}",
                "review_count": rng.choice([3, 17, 45, 120, 380, 940, 2500, 7100]),
                "image_url": image_url(product_id),
                "collected_at": collected_at,
            })
    return rows


def load_or_build_sample(sample_path: Path, queries: list[str], top_n: int) -> list[dict]:
    """Берёт мок-данные с диска, а если их нет — генерит и сохраняет."""
    if sample_path.exists():
        with sample_path.open(encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.DictReader(fh))
        if rows:
            print(f"Использую готовые мок-данные: {sample_path} ({len(rows)} строк)")
            return rows

    rows = build_sample_rows(queries, top_n)
    write_products(sample_path, rows)
    print(f"Сгенерированы мок-данные: {sample_path} ({len(rows)} строк)")
    return rows


# --------------------------------------------------------------------------- #
#  Запись
# --------------------------------------------------------------------------- #
def write_products(path: Path, rows: list[dict]) -> None:
    ensure_parent(path)
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=PRODUCT_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({col: row.get(col, "") for col in PRODUCT_COLUMNS})


def write_raw(path: Path, raw: dict) -> None:
    ensure_parent(path)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(raw, fh, ensure_ascii=False, indent=2)


def main() -> None:
    parser = argparse.ArgumentParser(description="Сбор карточек Wildberries по запросам")
    parser.add_argument("--config", default="config.yaml")
    parser.add_argument("--sample", action="store_true",
                        help="не ходить в сеть, сразу использовать мок-данные")
    parser.add_argument("--top-n", type=int, default=None,
                        help="сколько карточек брать по запросу (переопределяет config.yaml)")
    args = parser.parse_args()

    cfg = load_config(args.config)
    wb_cfg = dict((cfg.get("marketplaces") or {}).get("wildberries") or {})
    if args.top_n:
        wb_cfg["top_n"] = args.top_n
    top_n = int(wb_cfg.get("top_n", 20))

    out_cfg = cfg.get("output") or {}
    products_path = resolve(out_cfg.get("products_file", "data/products.csv"))
    raw_path = resolve(out_cfg.get("raw_wb_file", "data/raw_wb_products.json"))
    sample_path = resolve(out_cfg.get("sample_file", "data/sample_products.csv"))

    queries = read_queries((cfg.get("input") or {}).get("queries_file", "data/queries.csv"))
    if not queries:
        sys.exit("В файле запросов пусто — добавьте строки в data/queries.csv")
    print(f"Запросов к сбору: {len(queries)}, top_n={top_n}")

    rows: list[dict] = []
    raw: dict = {}

    if args.sample:
        print("Режим --sample: сеть не используется.")
    elif not wb_cfg.get("enabled", True):
        print("Wildberries выключен в config.yaml — использую мок-данные.")
    else:
        rows, raw = collect_live(queries, wb_cfg)

    if rows:
        write_raw(raw_path, raw)
        write_products(products_path, rows)
        print(f"\nГотово: {len(rows)} карточек -> {products_path}")
        print(f"Сырые данные -> {raw_path}")
        return

    # Fallback: живых данных нет
    if not args.sample and wb_cfg.get("enabled", True):
        print("\nWildberries не ответил ни по одному запросу — включаю fallback на мок-данные.")
        print("Запустите scripts/check_wb.py — он покажет, в чём причина.")
        if curl_requests is None:
            print("Совет: поставьте маскировку под Chrome командой  pip install curl_cffi  —")
            print("обычный requests WB часто отсекает с ошибкой 403.")

    rows = load_or_build_sample(sample_path, queries, top_n)
    write_products(products_path, rows)
    if not raw_path.exists():
        write_raw(raw_path, {"_note": "живой сбор не выполнялся, использованы мок-данные"})
    print(f"\nГотово (fallback): {len(rows)} карточек -> {products_path}")


if __name__ == "__main__":
    main()
