"""Общие мелочи для всех скриптов: конфиг, пути, чтение запросов.

Держим здесь только то, что реально используется больше чем в одном скрипте.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

import yaml

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
