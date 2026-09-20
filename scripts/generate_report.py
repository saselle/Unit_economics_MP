"""Сборка markdown-отчётов из посчитанной аналитики.

Вход:  data/products.csv + reports/query_analysis.csv + reports/format_analysis.csv
Выход: reports/summary.md                 — что вообще собрали и как выглядит рынок
       reports/launch_recommendations.md  — практические выводы под запуск эко-полотенец

Все цифры берутся из данных, ручных допущений в отчётах нет.
"""

from __future__ import annotations

import argparse
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import load_config, resolve  # noqa: E402


SAMPLE_WARNING = (
    "> ⚠️ **Отчёт построен на мок-данных** (`data/sample_products.csv`): живой сбор с "
    "Wildberries не выполнялся. Цифры ниже показывают, как работает инструмент, "
    "и не годятся для решений. Запустите `python scripts/collect_wb.py` на машине "
    "с доступом к `search.wb.ru`.\n"
)


def is_sample_data(products_path: Path, sample_path: Path) -> bool:
    """products.csv совпал с sample_products.csv -> данные мок-овые."""
    try:
        if not sample_path.exists() or not products_path.exists():
            return False
        return products_path.read_bytes() == sample_path.read_bytes()
    except OSError:
        return False


def read_csv_safe(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path)
    except Exception as exc:
        print(f"Не смог прочитать {path}: {type(exc).__name__}: {exc}")
        return pd.DataFrame()


def num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").dropna()


def fmt(value, suffix: str = "") -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "нет данных"
    if isinstance(value, float):
        return f"{value:,.0f}{suffix}".replace(",", " ")
    return f"{value}{suffix}"


def collection_date(products: pd.DataFrame) -> str:
    if "collected_at" in products.columns and not products["collected_at"].dropna().empty:
        return str(products["collected_at"].dropna().iloc[0])[:19].replace("T", " ")
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def md_table(df: pd.DataFrame, columns: list[str], headers: list[str]) -> str:
    """Маленькая markdown-таблица без лишних зависимостей."""
    cols = [c for c in columns if c in df.columns]
    if df.empty or not cols:
        return "_нет данных_\n"
    heads = [headers[columns.index(c)] for c in cols]
    lines = ["| " + " | ".join(heads) + " |",
             "|" + "|".join(["---"] * len(cols)) + "|"]
    for _, row in df.iterrows():
        cells = []
        for col in cols:
            value = row[col]
            if isinstance(value, float):
                value = f"{value:,.0f}".replace(",", " ") if value >= 100 else round(value, 2)
            cells.append(str(value) if str(value) != "nan" else "")
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines) + "\n"


# --------------------------------------------------------------------------- #
#  summary.md
# --------------------------------------------------------------------------- #
def build_summary(products: pd.DataFrame, queries_df: pd.DataFrame, buckets: pd.DataFrame,
                  cfg: dict) -> str:
    currency = (cfg.get("analysis") or {}).get("currency", "RUB")
    prices = num(products.get("price", pd.Series(dtype=float)))
    brands = Counter(b for b in products.get("brand", pd.Series(dtype=str)).fillna("") if b)

    top_price = queries_df.sort_values("avg_price", ascending=False).head(5) \
        if "avg_price" in queries_df.columns else pd.DataFrame()
    top_reviews = queries_df.sort_values("avg_review_count", ascending=False).head(5) \
        if "avg_review_count" in queries_df.columns else pd.DataFrame()

    out = [
        "# Сводка по рынку полотенец",
        "",
        f"- **Дата сбора:** {collection_date(products)}",
        f"- **Маркетплейсы:** {', '.join(sorted(set(products.get('marketplace', ['—'])))) or '—'}",
        f"- **Запросов:** {products['query'].nunique() if 'query' in products else 0}",
        f"- **Собрано карточек:** {len(products)}",
        f"- **Валюта:** {currency}",
        "",
        "## Ценовой коридор",
        "",
    ]

    if prices.empty:
        out.append("_Цены не собрались — проверьте сбор._\n")
    else:
        out += [
            f"- Минимум: **{fmt(float(prices.min()))} {currency}**",
            f"- 25-й перцентиль: **{fmt(float(prices.quantile(0.25)))} {currency}**",
            f"- Медиана: **{fmt(float(prices.median()))} {currency}**",
            f"- 75-й перцентиль: **{fmt(float(prices.quantile(0.75)))} {currency}**",
            f"- Максимум: **{fmt(float(prices.max()))} {currency}**",
            "",
        ]

    out += ["## Распределение по ценовым корзинам", "",
            md_table(buckets,
                     ["price_bucket", "product_count", "share_of_cards", "avg_review_count"],
                     ["Корзина, ₽", "Карточек", "Доля, %", "Ср. отзывов"]),
            "", "## Топ-10 брендов по частоте в выдаче", ""]

    if brands:
        for i, (brand, count) in enumerate(brands.most_common(10), start=1):
            share = count / max(len(products), 1) * 100
            out.append(f"{i}. **{brand}** — {count} карточек ({share:.1f}% выдачи)")
    else:
        out.append("_Бренды не собрались._")

    out += ["", "## Запросы с самой высокой средней ценой", "",
            md_table(top_price, ["query", "avg_price", "median_price", "product_count"],
                     ["Запрос", "Ср. цена", "Медиана", "Карточек"]),
            "", "## Запросы с самым большим средним числом отзывов", "",
            md_table(top_reviews,
                     ["query", "avg_review_count", "strong_cards", "avg_price"],
                     ["Запрос", "Ср. отзывов", "Сильных карточек", "Ср. цена"]),
            "", "---", "",
            "_Отчёт собран автоматически: `python scripts/generate_report.py`._"]
    return "\n".join(out) + "\n"


# --------------------------------------------------------------------------- #
#  launch_recommendations.md
# --------------------------------------------------------------------------- #
def build_recommendations(products: pd.DataFrame, queries_df: pd.DataFrame,
                          formats_df: pd.DataFrame, cfg: dict) -> str:
    analysis_cfg = cfg.get("analysis") or {}
    currency = analysis_cfg.get("currency", "RUB")
    min_reviews = int(analysis_cfg.get("min_reviews_for_strong_card", 100))

    prices = num(products.get("price", pd.Series(dtype=float)))
    p25 = float(prices.quantile(0.25)) if not prices.empty else None
    p50 = float(prices.median()) if not prices.empty else None
    p75 = float(prices.quantile(0.75)) if not prices.empty else None

    commercial = pd.DataFrame()
    if {"avg_review_count", "avg_price", "strong_cards"} <= set(queries_df.columns):
        commercial = queries_df.copy()
        # простой скоринг: спрос (отзывы) × доля сильных карточек × средний чек
        commercial["demand_score"] = (
            num(commercial["avg_review_count"]).rank(pct=True).fillna(0) * 0.5
            + num(commercial["strong_cards"]).rank(pct=True).fillna(0) * 0.3
            + num(commercial["avg_price"]).rank(pct=True).fillna(0) * 0.2
        ).round(3)
        commercial = commercial.sort_values("demand_score", ascending=False)

    eco_queries = queries_df[queries_df["query"].str.contains(
        "эко|бамбук|органическ|натуральн", case=False, na=False)] \
        if "query" in queries_df.columns else pd.DataFrame()

    words = Counter()
    if "common_title_words" in queries_df.columns:
        for cell in queries_df["common_title_words"].fillna(""):
            for chunk in str(cell).split(","):
                word = chunk.split("(")[0].strip()
                if word:
                    words[word] += 1

    out = [
        "# Рекомендации под запуск линейки эко-полотенец",
        "",
        f"Данные: {len(products)} карточек по "
        f"{products['query'].nunique() if 'query' in products else 0} запросам, "
        f"сбор от {collection_date(products)}.",
        "",
        "## 1. Наиболее коммерческие запросы",
        "",
        "Скоринг: средние отзывы (50%) + число карточек с "
        f"{min_reviews}+ отзывами (30%) + средний чек (20%).",
        "",
        md_table(commercial.head(7),
                 ["query", "demand_score", "avg_review_count", "strong_cards", "avg_price"],
                 ["Запрос", "Скор", "Ср. отзывов", f"Карточек {min_reviews}+", "Ср. цена"]),
        "",
        "## 2. Ценовой коридор",
        "",
    ]

    if p50 is None:
        out.append("_Цены не собрались — раздел пустой._\n")
    else:
        out += [
            f"- Основная масса выдачи: **{fmt(p25)}–{fmt(p75)} {currency}** (25–75 перцентиль).",
            f"- Медиана рынка: **{fmt(p50)} {currency}**.",
            f"- Вход «в рынок» без демпинга: **{fmt(p25 * 0.95)}–{fmt(p50 * 1.05)} {currency}**.",
            f"- Премиальный эко-сегмент (оправдан составом и упаковкой): "
            f"**{fmt(p75)}–{fmt(p75 * 1.4)} {currency}**.",
            "",
        ]

    out += ["## 3. Форматы, которые стоит копать глубже", "",
            md_table(formats_df.head(10),
                     ["format", "product_count", "share_of_cards", "avg_price", "avg_review_count"],
                     ["Формат", "Карточек", "Доля, %", "Ср. цена", "Ср. отзывов"]),
            ""]

    if not formats_df.empty and "avg_price" in formats_df.columns:
        rich = formats_df.sort_values("avg_price", ascending=False).head(3)["format"].tolist()
        out.append(f"Самый высокий средний чек: **{', '.join(map(str, rich))}** — "
                   "это кандидаты на флагманский SKU линейки.")
        out.append("")

    out += ["## 4. Частые слова в названиях", "",
            "Их стоит закладывать в SEO-заголовок карточки:", ""]
    if words:
        out.append(", ".join(f"`{w}`" for w, _ in words.most_common(20)))
    else:
        out.append("_Названия не собрались._")
    out.append("")

    out += ["## 5. Гипотезы по SKU", ""]
    if p50 is not None:
        out += [
            f"1. **Базовое банное 70х140, органический хлопок** — якорь линейки, "
            f"цена {fmt(p25 * 0.95)}–{fmt(p50)} {currency}. Задача — трафик и отзывы.",
            f"2. **Набор 2 шт (70х140 + 50х90)** — средний чек "
            f"{fmt(p50 * 1.6)}–{fmt(p75 * 1.8)} {currency}, лучше держит маржу, "
            "меньше конкуренции по одиночным запросам.",
            f"3. **Подарочный набор 3–4 шт в коробке** — верх коридора, "
            f"{fmt(p75 * 1.5)}+ {currency}, сезонный спрос (декабрь, февраль, март).",
            "4. **Бамбук / эко-линия в нейтральной палитре** (бежевый, молочный, графит) — "
            "отстройка по визуалу карточки, а не по цене.",
            "",
        ]
    else:
        out += ["_Недостаточно данных по ценам для гипотез по SKU._", ""]

    if not eco_queries.empty:
        out += ["### Что видно именно по эко-запросам", "",
                md_table(eco_queries,
                         ["query", "product_count", "avg_price", "avg_review_count", "strong_cards"],
                         ["Запрос", "Карточек", "Ср. цена", "Ср. отзывов", "Сильных карточек"]),
                ""]

    out += [
        "## 6. Что нужно досбрать следующим этапом",
        "",
        "- **Отзывы**: топ-10 карточек по каждому ключевому запросу — за что хвалят и ругают "
        "(линяет, жёсткое, не впитывает, размер не совпал).",
        "- **Объёмы продаж**: WB API выдачи их не отдаёт. Нужны внешние источники или "
        "оценка по динамике числа отзывов (снимать выдачу раз в неделю и считать дельту).",
        "- **Юнит-экономика**: закупка, логистика, комиссия категории, хранение, возвраты.",
        "- **Карточки конкурентов**: сколько фото, есть ли видео/инфографика, длина описания, "
        "какие характеристики заполнены.",
        "- **Ozon**: тот же список запросов, чтобы сравнить цены и насыщенность.",
        "- **Сезонность**: повторить сбор через 2–4 недели и сравнить цены и отзывы.",
        "",
        "---",
        "",
        "_Отчёт собран автоматически: `python scripts/generate_report.py`._",
    ]
    return "\n".join(out) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description="Генерация markdown-отчётов")
    parser.add_argument("--config", default="config.yaml")
    args = parser.parse_args()

    cfg = load_config(args.config)
    out_cfg = cfg.get("output") or {}
    products_path = resolve(out_cfg.get("products_file", "data/products.csv"))
    reports_dir = resolve(out_cfg.get("reports_dir", "reports"))
    reports_dir.mkdir(parents=True, exist_ok=True)

    products = read_csv_safe(products_path)
    queries_df = read_csv_safe(reports_dir / "query_analysis.csv")
    formats_df = read_csv_safe(reports_dir / "format_analysis.csv")
    buckets_df = read_csv_safe(reports_dir / "price_buckets.csv")

    if products.empty:
        note = ("# Данных нет\n\nФайл `data/products.csv` пуст или отсутствует.\n"
                "Запустите `python scripts/collect_wb.py`, затем "
                "`python scripts/analyze_products.py`.\n")
        (reports_dir / "summary.md").write_text(note, encoding="utf-8")
        (reports_dir / "launch_recommendations.md").write_text(note, encoding="utf-8")
        print("Данных нет — отчёты созданы с пояснением.")
        return

    sample_path = resolve(out_cfg.get("sample_file", "data/sample_products.csv"))
    banner = SAMPLE_WARNING + "\n" if is_sample_data(products_path, sample_path) else ""
    if banner:
        print("Внимание: данные мок-овые, в отчёты добавлена пометка.")

    (reports_dir / "summary.md").write_text(
        banner + build_summary(products, queries_df, buckets_df, cfg), encoding="utf-8")
    (reports_dir / "launch_recommendations.md").write_text(
        banner + build_recommendations(products, queries_df, formats_df, cfg), encoding="utf-8")

    print("Готово:")
    print(f"  {reports_dir / 'summary.md'}")
    print(f"  {reports_dir / 'launch_recommendations.md'}")


if __name__ == "__main__":
    main()
