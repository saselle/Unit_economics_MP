"""Аналитика по собранным карточкам.

Вход:  data/products.csv
Выход: reports/query_analysis.csv   — срез по каждому поисковому запросу
       reports/format_analysis.csv  — срез по форматам товара (набор / банное / бамбук ...)
       reports/price_buckets.csv    — распределение по ценовым корзинам из config.yaml

Скрипт не падает на пустых и неполных данных: пустые файлы просто получают заголовки.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import load_config, resolve  # noqa: E402

QUERY_COLUMNS = [
    "query", "product_count", "avg_price", "median_price", "min_price", "max_price",
    "avg_rating", "avg_review_count", "median_review_count", "strong_cards",
    "top_brands", "common_title_words",
]
FORMAT_COLUMNS = [
    "format", "product_count", "share_of_cards", "avg_price", "median_price",
    "min_price", "max_price", "avg_rating", "avg_review_count", "top_brands",
]
BUCKET_COLUMNS = ["price_bucket", "product_count", "share_of_cards", "avg_review_count", "avg_rating"]

# Слова, которые не несут смысла в анализе названий.
STOPWORDS = {
    "и", "для", "в", "на", "с", "от", "по", "шт", "см", "х", "x", "the", "of",
    "полотенце", "полотенца", "полотенец", "полотенцa",
}
WORD_RE = re.compile(r"[а-яёa-z0-9]+", re.IGNORECASE)


def load_products(path: Path) -> pd.DataFrame:
    """Читает products.csv и приводит числовые поля к числам."""
    if not path.exists():
        print(f"Нет файла {path}. Сначала запустите scripts/collect_wb.py")
        return pd.DataFrame()

    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    if df.empty:
        return df

    for col in ("price", "old_price", "discount", "rating", "review_count", "rank"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].replace("", None), errors="coerce")
        else:
            df[col] = pd.NA

    for col in ("query", "title", "brand", "marketplace"):
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()

    return df


def top_brands(series: pd.Series, limit: int) -> str:
    counts = Counter(b for b in series if b)
    return ", ".join(f"{brand} ({n})" for brand, n in counts.most_common(limit))


def common_words(titles: pd.Series, limit: int) -> str:
    counter: Counter[str] = Counter()
    for title in titles:
        words = {w.lower() for w in WORD_RE.findall(str(title))}
        counter.update(w for w in words if w not in STOPWORDS and len(w) > 2)
    return ", ".join(f"{word} ({n})" for word, n in counter.most_common(limit))


def round_or_blank(value) -> float | str:
    """Округляем, но не выдаём NaN в CSV — вместо него пустая ячейка."""
    if pd.isna(value):
        return ""
    return round(float(value), 2)


def detect_formats(title: str, format_rules: dict) -> list[str]:
    """Определяет форматы товара по названию. Один товар может попасть в несколько."""
    title_low = str(title).lower()
    found = [fmt for fmt, keys in format_rules.items()
             if any(str(k).lower() in title_low for k in (keys or []))]
    # всё, что не опознано как набор, считаем одиночным полотенцем
    if not any(fmt in ("набор", "подарочный набор") for fmt in found):
        found.append("одиночное полотенце")
    return found


def analyze_queries(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    analysis_cfg = cfg.get("analysis") or {}
    min_reviews = int(analysis_cfg.get("min_reviews_for_strong_card", 100))
    brands_limit = int(analysis_cfg.get("top_brands_limit", 3))
    words_limit = int(analysis_cfg.get("top_words_limit", 8))

    rows = []
    for query, group in df.groupby("query", sort=False):
        rows.append({
            "query": query,
            "product_count": len(group),
            "avg_price": round_or_blank(group["price"].mean()),
            "median_price": round_or_blank(group["price"].median()),
            "min_price": round_or_blank(group["price"].min()),
            "max_price": round_or_blank(group["price"].max()),
            "avg_rating": round_or_blank(group["rating"].mean()),
            "avg_review_count": round_or_blank(group["review_count"].mean()),
            "median_review_count": round_or_blank(group["review_count"].median()),
            "strong_cards": int((group["review_count"] >= min_reviews).sum()),
            "top_brands": top_brands(group["brand"], brands_limit),
            "common_title_words": common_words(group["title"], words_limit),
        })
    return pd.DataFrame(rows, columns=QUERY_COLUMNS)


def analyze_formats(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    analysis_cfg = cfg.get("analysis") or {}
    format_rules = analysis_cfg.get("formats") or {}
    brands_limit = int(analysis_cfg.get("top_brands_limit", 3))

    exploded = df.copy()
    exploded["format"] = exploded["title"].apply(lambda t: detect_formats(t, format_rules))
    exploded = exploded.explode("format")

    total = len(df) or 1
    rows = []
    for fmt, group in exploded.groupby("format", sort=False):
        rows.append({
            "format": fmt,
            "product_count": len(group),
            "share_of_cards": round(len(group) / total * 100, 1),
            "avg_price": round_or_blank(group["price"].mean()),
            "median_price": round_or_blank(group["price"].median()),
            "min_price": round_or_blank(group["price"].min()),
            "max_price": round_or_blank(group["price"].max()),
            "avg_rating": round_or_blank(group["rating"].mean()),
            "avg_review_count": round_or_blank(group["review_count"].mean()),
            "top_brands": top_brands(group["brand"], brands_limit),
        })
    result = pd.DataFrame(rows, columns=FORMAT_COLUMNS)
    if not result.empty:
        result = result.sort_values("product_count", ascending=False)
    return result


def analyze_buckets(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    buckets = (cfg.get("analysis") or {}).get("price_buckets") or []
    total = len(df) or 1
    rows = []
    for bucket in buckets:
        try:
            low, high = float(bucket[0]), float(bucket[1])
        except (TypeError, ValueError, IndexError):
            continue
        group = df[(df["price"] >= low) & (df["price"] < high)]
        rows.append({
            "price_bucket": f"{int(low)}–{int(high)}",
            "product_count": len(group),
            "share_of_cards": round(len(group) / total * 100, 1),
            "avg_review_count": round_or_blank(group["review_count"].mean()),
            "avg_rating": round_or_blank(group["rating"].mean()),
        })
    return pd.DataFrame(rows, columns=BUCKET_COLUMNS)


def save(df: pd.DataFrame, path: Path, columns: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if df.empty:
        df = pd.DataFrame(columns=columns)
    df.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"  {path}  ({len(df)} строк)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Аналитика по собранным карточкам")
    parser.add_argument("--config", default="config.yaml")
    args = parser.parse_args()

    cfg = load_config(args.config)
    out_cfg = cfg.get("output") or {}
    products_path = resolve(out_cfg.get("products_file", "data/products.csv"))
    reports_dir = resolve(out_cfg.get("reports_dir", "reports"))

    df = load_products(products_path)
    if df.empty:
        print("Данных нет — создаю пустые отчёты со схемой.")
        save(pd.DataFrame(), reports_dir / "query_analysis.csv", QUERY_COLUMNS)
        save(pd.DataFrame(), reports_dir / "format_analysis.csv", FORMAT_COLUMNS)
        save(pd.DataFrame(), reports_dir / "price_buckets.csv", BUCKET_COLUMNS)
        return

    print(f"Карточек: {len(df)}, запросов: {df['query'].nunique()}")
    print("Сохраняю:")
    save(analyze_queries(df, cfg), reports_dir / "query_analysis.csv", QUERY_COLUMNS)
    save(analyze_formats(df, cfg), reports_dir / "format_analysis.csv", FORMAT_COLUMNS)
    save(analyze_buckets(df, cfg), reports_dir / "price_buckets.csv", BUCKET_COLUMNS)


if __name__ == "__main__":
    main()
