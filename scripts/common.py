"""Общие мелочи для всех скриптов: конфиг, пути, чтение запросов.

Держим здесь только то, что реально используется больше чем в одном скрипте.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

import requests
import yaml

try:
    # curl_cffi повторяет TLS-отпечаток настоящего Chrome. Wildberries отдаёт 403
    # обычному requests именно из-за отпечатка, поэтому используем её, если стоит.
    from curl_cffi import requests as curl_requests
except ImportError:
    curl_requests = None

PROJECT_ROOT = Path(__file__).resolve().parent.parent

PRODUCT_COLUMNS = [
    "marketplace",
    "query",
    "rank",
    "product_id",
    "product_url",
    "title",
    "brand",
    "price",
    "old_price",
    "discount",
    "rating",
    "review_count",
    "image_url",
    "collected_at",
]


def load_config(path: str | Path = "config.yaml") -> dict:
    """Читает config.yaml. Если файла нет — падаем с понятным сообщением."""
    cfg_path = Path(path)
    if not cfg_path.is_absolute():
        cfg_path = PROJECT_ROOT / cfg_path
    if not cfg_path.exists():
        sys.exit(f"Не найден конфиг: {cfg_path}. Скопируйте config.yaml в корень проекта.")
    with cfg_path.open(encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def resolve(path_str: str) -> Path:
    """Путь из конфига -> абсолютный путь относительно корня проекта."""
    path = Path(path_str)
    return path if path.is_absolute() else PROJECT_ROOT / path


def read_queries(path_str: str) -> list[str]:
    """Читает список запросов из CSV.

    Поддерживает и файл с колонкой `query`, и файл из одной колонки без заголовка.
    """
    path = resolve(path_str)
    if not path.exists():
        sys.exit(f"Не найден файл запросов: {path}")

    queries: list[str] = []
    with path.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))

    if not rows:
        return []

    header = [c.strip().lower() for c in rows[0]]
    if "query" in header:
        idx = header.index("query")
        data_rows = rows[1:]
    else:
        idx = 0
        data_rows = rows

    for row in data_rows:
        if not row or idx >= len(row):
            continue
        value = row[idx].strip()
        if value:
            queries.append(value)

    # убираем дубли, сохраняя порядок
    seen: set[str] = set()
    unique: list[str] = []
    for q in queries:
        if q.lower() not in seen:
            seen.add(q.lower())
            unique.append(q)
    return unique


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def make_session() -> tuple[object, str]:
    """Возвращает HTTP-сессию и название движка.

    Если установлена curl_cffi — запросы уходят с отпечатком Chrome, и WB
    пропускает их. Если нет — работаем обычным requests (может прилететь 403).
    """
    if curl_requests is not None:
        return curl_requests.Session(impersonate="chrome"), "curl_cffi (маскировка под Chrome)"
    return requests.Session(), "requests (без маскировки)"


# Диапазоны корзин WB для картинок и card.json. Таблица со временем устаревает,
# поэтому basket_candidates() отдаёт расчётный хост первым, а дальше — перебор.
BASKET_RANGES = [
    (143, 1), (287, 2), (431, 3), (719, 4), (1007, 5), (1061, 6), (1115, 7),
    (1169, 8), (1313, 9), (1601, 10), (1655, 11), (1919, 12), (2045, 13),
    (2189, 14), (2405, 15), (2621, 16), (2837, 17), (3053, 18), (3269, 19),
    (3485, 20), (3701, 21), (3917, 22), (4133, 23), (4349, 24), (4565, 25),
]
MAX_BASKET = 60


def basket_parts(nm_id) -> tuple[int, int]:
    """Артикул -> (vol, part) для адресов вида /vol{vol}/part{part}/."""
    nm = int(nm_id)
    return nm // 100_000, nm // 1_000


def basket_number(nm_id) -> int:
    """Расчётный номер корзины по таблице диапазонов."""
    vol, _ = basket_parts(nm_id)
    for upper, number in BASKET_RANGES:
        if vol <= upper:
            return number
    # выше таблицы WB нарезает корзины шагом примерно 312 томов
    return min(26 + (vol - 4566) // 312, MAX_BASKET)


def basket_candidates(nm_id, first: int | None = None) -> list[int]:
    """Номера корзин в порядке проверки: подсказка, расчёт, затем остальные."""
    order = []
    for number in (first, basket_number(nm_id)):
        if number and number not in order:
            order.append(number)
    order += [n for n in range(1, MAX_BASKET + 1) if n not in order]
    return order


def basket_image_url(nm_id) -> str:
    """Ссылка на главное фото карточки."""
    try:
        vol, part = basket_parts(nm_id)
    except (TypeError, ValueError):
        return ""
    return (f"https://basket-{basket_number(nm_id):02d}.wbbasket.ru"
            f"/vol{vol}/part{part}/{nm_id}/images/big/1.webp")
