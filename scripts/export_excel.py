"""Сборка Excel-отчёта из data/products.csv.

Делает reports/marketplace_report.xlsx с листами:
  Сводка   — ключевые цифры рынка и топ брендов;
  Карточки — все собранные товары с фильтрами и кликабельными ссылками;
  Запросы  — срез по каждому поисковому запросу;
  Форматы  — срезы по формату, материалу, назначению и размеру;
  Цены     — распределение по ценовым корзинам.

Все цифры на листах-срезах — формулы Excel, считаются от листа «Карточки».
Если отфильтровать или дописать строки, отчёт пересчитается сам.

Запуск:
    python scripts/export_excel.py
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import load_config, resolve  # noqa: E402

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="2F4858")
ACCENT_FILL = PatternFill("solid", fgColor="EFEAE3")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
GREY_RGB = "666666"
TITLE_FONT = Font(name=FONT, bold=True, size=13)
BODY_FONT = Font(name=FONT, size=10)
LINK_FONT = Font(name=FONT, size=10, color="1155CC", underline="single")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = "#,##0 ₽"
PERCENT = "0%"
RATING = "0.0"
COUNT = "#,##0"

# Колонки листа «Карточки» в порядке вывода.
CARD_COLUMNS = [
    ("Площадка", 12), ("Запрос", 26), ("Позиция", 9), ("Артикул", 13),
    ("Название", 52), ("Бренд", 18), ("Цена", 11), ("Цена до скидки", 14),
    ("Скидка", 9), ("Рейтинг", 9), ("Отзывов", 10), ("Формат", 17),
    ("Материал", 14), ("Назначение", 14), ("Размер", 11), ("Ссылка", 14),
    ("Дата сбора", 19),
]
COL = {name: get_column_letter(i) for i, (name, _) in enumerate(CARD_COLUMNS, start=1)}

STOPWORDS = {
    "и", "для", "в", "на", "с", "от", "по", "шт", "см", "х", "x",
    "полотенце", "полотенца", "полотенец", "набор", "наборы",
}
WORD_RE = re.compile(r"[а-яёa-z0-9]+", re.IGNORECASE)
SIZE_RE = re.compile(r"(\d{2,3})\s*[хx*]\s*(\d{2,3})", re.IGNORECASE)


# --------------------------------------------------------------------------- #
#  Разбор названий
# --------------------------------------------------------------------------- #
def detect_format(title: str) -> str:
    t = title.lower()
    if any(k in t for k in ("подарочн", "в коробке", "подарок")):
        return "Подарочный набор"
    if any(k in t for k in ("набор", "комплект", "2 шт", "3 шт", "4 шт", "5 шт", "6 шт",
                            "2шт", "3шт", "4шт", "5шт", "6шт")):
        return "Набор"
    return "Одиночное"


def detect_material(title: str) -> str:
    t = title.lower()
    for key, name in (("бамбук", "Бамбук"), ("микрофибр", "Микрофибра"),
                      ("вафельн", "Вафельное"), ("махров", "Махровое"),
                      ("органическ", "Органический хлопок"), ("хлопок", "Хлопок"),
                      ("хлопков", "Хлопок"), ("лен", "Лён")):
        if key in t:
            return name
    return "Не указан"


def detect_purpose(title: str) -> str:
    t = title.lower()
    for key, name in (("банн", "Банное"), ("для лица", "Для лица"), ("лицев", "Для лица"),
                      ("для рук", "Для рук"), ("пляжн", "Пляжное"),
                      ("кухон", "Кухонное"), ("детск", "Детское")):
        if key in t:
            return name
    return "Не указано"


def detect_size(title: str) -> str:
    match = SIZE_RE.search(title)
    return f"{match.group(1)}х{match.group(2)}" if match else ""


def top_words(titles, limit: int = 4) -> str:
    counter: Counter[str] = Counter()
    for title in titles:
        counter.update({w.lower() for w in WORD_RE.findall(str(title))
                        if w.lower() not in STOPWORDS and len(w) > 3})
    return ", ".join(word for word, _ in counter.most_common(limit))


# --------------------------------------------------------------------------- #
#  Оформление
# --------------------------------------------------------------------------- #
def write_header(ws, headers: list[str], row: int = 1) -> None:
    for i, name in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=i, value=name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER


def style_row(ws, row: int, columns: int) -> None:
    for i in range(1, columns + 1):
        cell = ws.cell(row=row, column=i)
        cell.font = cell.font if cell.font.color and cell.font.color.rgb == "001155CC" else BODY_FONT
        cell.border = BORDER


def set_widths(ws, widths: list[int]) -> None:
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width


# --------------------------------------------------------------------------- #
#  Листы
# --------------------------------------------------------------------------- #
def build_cards(ws, df: pd.DataFrame) -> int:
    """Лист «Карточки». Возвращает номер последней строки с данными."""
    write_header(ws, [name for name, _ in CARD_COLUMNS])
    set_widths(ws, [width for _, width in CARD_COLUMNS])

    for i, (_, item) in enumerate(df.iterrows(), start=2):
        title = str(item.get("title", ""))
        values = [
            item.get("marketplace", ""), item.get("query", ""), item.get("rank", ""),
            item.get("product_id", ""), title, item.get("brand", ""),
            item.get("price"), item.get("old_price"),
            (item["discount"] / 100) if pd.notna(item.get("discount")) else None,
            item.get("rating"), item.get("review_count"),
            detect_format(title), detect_material(title), detect_purpose(title),
            detect_size(title), "открыть карточку", item.get("collected_at", ""),
        ]
        for j, value in enumerate(values, start=1):
            cell = ws.cell(row=i, column=j)
            cell.value = None if (isinstance(value, float) and pd.isna(value)) else value
            cell.font = BODY_FONT
            cell.border = BORDER

        link = str(item.get("product_url", "") or "")
        if link:
            cell = ws.cell(row=i, column=16)
            cell.hyperlink = link
            cell.font = LINK_FONT

        ws.cell(row=i, column=7).number_format = MONEY
        ws.cell(row=i, column=8).number_format = MONEY
        ws.cell(row=i, column=9).number_format = PERCENT
        ws.cell(row=i, column=10).number_format = RATING
        ws.cell(row=i, column=11).number_format = COUNT
        ws.cell(row=i, column=5).alignment = Alignment(wrap_text=False)

    last = len(df) + 1
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CARD_COLUMNS))}{last}"
    ws.row_dimensions[1].height = 28
    return last


def rng(column: str, last: int) -> str:
    """Абсолютная ссылка на колонку листа «Карточки»."""
    return f"Карточки!${column}$2:${column}${last}"


def build_queries(ws, df: pd.DataFrame, last: int, min_reviews: int) -> None:
    headers = ["Запрос", "Карточек", "Средняя цена", "Средний рейтинг", "Ср. отзывов",
               f"Карточек {min_reviews}+ отзывов", "Топ-бренды выдачи",
               "Частые слова в названиях"]
    write_header(ws, headers)
    set_widths(ws, [26, 10, 13, 14, 12, 20, 34, 40])

    q_col, price, rating, reviews, brand, title = (
        rng(COL["Запрос"], last), rng(COL["Цена"], last), rng(COL["Рейтинг"], last),
        rng(COL["Отзывов"], last), rng(COL["Бренд"], last), rng(COL["Название"], last),
    )

    for i, query in enumerate(df["query"].drop_duplicates(), start=2):
        group = df[df["query"] == query]
        brands = Counter(b for b in group["brand"] if b)
        ws.cell(row=i, column=1, value=query)
        ws.cell(row=i, column=2, value=f'=COUNTIF({q_col},$A{i})')
        ws.cell(row=i, column=3, value=f'=IFERROR(AVERAGEIFS({price},{q_col},$A{i}),"")')
        ws.cell(row=i, column=4, value=f'=IFERROR(AVERAGEIFS({rating},{q_col},$A{i}),"")')
        ws.cell(row=i, column=5, value=f'=IFERROR(AVERAGEIFS({reviews},{q_col},$A{i}),"")')
        ws.cell(row=i, column=6,
                value=f'=COUNTIFS({q_col},$A{i},{reviews},">={min_reviews}")')
        ws.cell(row=i, column=7,
                value=", ".join(f"{b} ({n})" for b, n in brands.most_common(3)))
        ws.cell(row=i, column=8, value=top_words(group["title"], 5))

        ws.cell(row=i, column=3).number_format = MONEY
        ws.cell(row=i, column=4).number_format = RATING
        ws.cell(row=i, column=5).number_format = COUNT
        style_row(ws, i, len(headers))

    ws.freeze_panes = "B2"
    ws.row_dimensions[1].height = 30


def build_breakdowns(ws, df: pd.DataFrame, last: int) -> None:
    """Лист «Форматы»: несколько таблиц-срезов друг под другом."""
    set_widths(ws, [24, 11, 10, 13, 13, 13])
    price, reviews, rating = (rng(COL["Цена"], last), rng(COL["Отзывов"], last),
                              rng(COL["Рейтинг"], last))
    total = len(df)
    row = 1

    blocks = [
        ("Формат товара", COL["Формат"], "Формат"),
        ("Материал", COL["Материал"], "Материал"),
        ("Назначение", COL["Назначение"], "Назначение"),
        ("Размер в названии", COL["Размер"], "Размер"),
    ]

    for title, column, source in blocks:
        cell = ws.cell(row=row, column=1, value=title)
        cell.font = TITLE_FONT
        row += 1
        write_header(ws, ["Значение", "Карточек", "Доля", "Средняя цена",
                          "Ср. отзывов", "Средний рейтинг"], row=row)
        header_row = row
        row += 1

        column_values = df["title"].map(
            {"Формат": detect_format, "Материал": detect_material,
             "Назначение": detect_purpose, "Размер": detect_size}[source]
        )
        counts = Counter(v for v in column_values if str(v).strip())
        criteria = rng(column, last)

        for value, _ in counts.most_common(10):
            ws.cell(row=row, column=1, value=value)
            ws.cell(row=row, column=2, value=f'=COUNTIF({criteria},$A{row})')
            ws.cell(row=row, column=3, value=f'=IFERROR($B{row}/{total},"")')
            ws.cell(row=row, column=4, value=f'=IFERROR(AVERAGEIFS({price},{criteria},$A{row}),"")')
            ws.cell(row=row, column=5, value=f'=IFERROR(AVERAGEIFS({reviews},{criteria},$A{row}),"")')
            ws.cell(row=row, column=6, value=f'=IFERROR(AVERAGEIFS({rating},{criteria},$A{row}),"")')
            ws.cell(row=row, column=3).number_format = PERCENT
            ws.cell(row=row, column=4).number_format = MONEY
            ws.cell(row=row, column=5).number_format = COUNT
            ws.cell(row=row, column=6).number_format = RATING
            style_row(ws, row, 6)
            row += 1

        ws.row_dimensions[header_row].height = 26
        row += 2


def build_prices(ws, df: pd.DataFrame, last: int, buckets: list) -> None:
    ws.cell(row=1, column=1, value="Распределение выдачи по цене").font = TITLE_FONT
    write_header(ws, ["Ценовая корзина", "Карточек", "Доля", "Ср. отзывов",
                      "Средний рейтинг"], row=2)
    set_widths(ws, [22, 11, 10, 13, 15])

    price, reviews, rating = (rng(COL["Цена"], last), rng(COL["Отзывов"], last),
                              rng(COL["Рейтинг"], last))
    total = len(df)
    row = 3
    for bucket in buckets:
        try:
            low, high = float(bucket[0]), float(bucket[1])
        except (TypeError, ValueError, IndexError):
            continue
        label = f"{int(low)}–{int(high)} ₽"
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2,
                value=f'=COUNTIFS({price},">={low}",{price},"<{high}")')
        ws.cell(row=row, column=3, value=f'=IFERROR($B{row}/{total},"")')
        ws.cell(row=row, column=4,
                value=f'=IFERROR(AVERAGEIFS({reviews},{price},">={low}",{price},"<{high}"),"")')
        ws.cell(row=row, column=5,
                value=f'=IFERROR(AVERAGEIFS({rating},{price},">={low}",{price},"<{high}"),"")')
        ws.cell(row=row, column=3).number_format = PERCENT
        ws.cell(row=row, column=4).number_format = COUNT
        ws.cell(row=row, column=5).number_format = RATING
        style_row(ws, row, 5)
        row += 1

    ws.cell(row=row + 1, column=1,
            value="Границы корзин заданы в config.yaml (analysis.price_buckets).").font = BODY_FONT


def build_summary(ws, df: pd.DataFrame, last: int, min_reviews: int, collected: str) -> None:
    set_widths(ws, [34, 18, 14, 30])
    ws.cell(row=1, column=1, value="Рынок полотенец: сводка по выдаче Wildberries").font = TITLE_FONT
    ws.cell(row=2, column=1, value=f"Дата сбора: {collected}").font = BODY_FONT
    ws.cell(row=3, column=1,
            value="Источник: публичная поисковая выдача Wildberries, топ-20 по каждому запросу."
            ).font = BODY_FONT
    ws.cell(row=4, column=1,
            value="Цифры на всех листах — формулы от листа «Карточки»: правите данные — "
                  "отчёт пересчитывается."
            ).font = Font(name=FONT, size=9, italic=True, color=GREY_RGB)

    price, reviews, rating = (rng(COL["Цена"], last), rng(COL["Отзывов"], last),
                              rng(COL["Рейтинг"], last))
    rows = [
        ("Собрано карточек", f'=COUNTA({rng(COL["Артикул"], last)})', COUNT),
        ("Поисковых запросов", '=COUNTA(Запросы!$A$2:$A$500)', COUNT),
        ("Минимальная цена", f'=MIN({price})', MONEY),
        ("25-й перцентиль цены", f'=QUARTILE({price},1)', MONEY),
        ("Медианная цена", f'=MEDIAN({price})', MONEY),
        ("Средняя цена", f'=AVERAGE({price})', MONEY),
        ("75-й перцентиль цены", f'=QUARTILE({price},3)', MONEY),
        ("Максимальная цена", f'=MAX({price})', MONEY),
        ("Средний рейтинг", f'=AVERAGE({rating})', RATING),
        ("Среднее число отзывов", f'=AVERAGE({reviews})', COUNT),
        (f"Карточек с {min_reviews}+ отзывов", f'=COUNTIF({reviews},">={min_reviews}")', COUNT),
        ("Средняя скидка", f'=AVERAGE({rng(COL["Скидка"], last)})', PERCENT),
    ]
    row = 6
    for label, formula, fmt in rows:
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        cell = ws.cell(row=row, column=2, value=formula)
        cell.font = Font(name=FONT, size=10, bold=True)
        cell.number_format = fmt
        cell.fill = ACCENT_FILL
        cell.border = BORDER
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Топ-10 брендов в выдаче").font = TITLE_FONT
    row += 1
    write_header(ws, ["Бренд", "Карточек", "Доля", "Средняя цена"], row=row)
    row += 1

    brand_col = rng(COL["Бренд"], last)
    total = len(df)
    for brand, _ in Counter(b for b in df["brand"] if b).most_common(10):
        ws.cell(row=row, column=1, value=brand).font = BODY_FONT
        ws.cell(row=row, column=2, value=f'=COUNTIF({brand_col},$A{row})')
        ws.cell(row=row, column=3, value=f'=IFERROR($B{row}/{total},"")')
        ws.cell(row=row, column=4, value=f'=IFERROR(AVERAGEIFS({price},{brand_col},$A{row}),"")')
        ws.cell(row=row, column=3).number_format = PERCENT
        ws.cell(row=row, column=4).number_format = MONEY
        style_row(ws, row, 4)
        row += 1


def main() -> None:
    parser = argparse.ArgumentParser(description="Excel-отчёт по собранным карточкам")
    parser.add_argument("--config", default="config.yaml")
    parser.add_argument("--out", default=None, help="путь к .xlsx (по умолчанию reports/)")
    args = parser.parse_args()

    cfg = load_config(args.config)
    out_cfg = cfg.get("output") or {}
    analysis_cfg = cfg.get("analysis") or {}
    products_path = resolve(out_cfg.get("products_file", "data/products.csv"))
    reports_dir = resolve(out_cfg.get("reports_dir", "reports"))
    out_path = Path(args.out) if args.out else reports_dir / "marketplace_report.xlsx"

    if not products_path.exists():
        sys.exit(f"Нет файла {products_path}. Сначала запустите scripts/collect_wb.py")

    df = pd.read_csv(products_path, dtype=str, keep_default_na=False)
    if df.empty:
        sys.exit("В products.csv нет строк — нечего выгружать.")

    for col in ("price", "old_price", "discount", "rating", "review_count", "rank"):
        df[col] = pd.to_numeric(df.get(col, "").replace("", None), errors="coerce")
    for col in ("query", "title", "brand", "marketplace", "product_url", "collected_at"):
        df[col] = df.get(col, "").fillna("").astype(str).str.strip()

    min_reviews = int(analysis_cfg.get("min_reviews_for_strong_card", 100))
    buckets = analysis_cfg.get("price_buckets") or []
    collected = (df["collected_at"].iloc[0] or "")[:19].replace("T", " ")

    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Сводка"
    ws_cards = wb.create_sheet("Карточки")
    ws_queries = wb.create_sheet("Запросы")
    ws_breakdown = wb.create_sheet("Форматы")
    ws_prices = wb.create_sheet("Цены")

    last = build_cards(ws_cards, df)
    build_queries(ws_queries, df, last, min_reviews)
    build_breakdowns(ws_breakdown, df, last)
    build_prices(ws_prices, df, last, buckets)
    build_summary(ws_summary, df, last, min_reviews, collected)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    print(f"Готово: {out_path}")
    print("Откройте файл в Excel — все цифры на листах-срезах пересчитываются формулами.")


if __name__ == "__main__":
    main()
