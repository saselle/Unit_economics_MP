"""Матрица SKU в Excel: характеристики, бенчмарк рынка, таргет по цене.

Вход:  data/sku_matrix.csv  — сама матрица (правится в Excel или Блокноте),
       data/products.csv    — выдача WB для бенчмарка,
       data/reviews.csv     — отзывы для среза по цветам (если собраны).

Выход: reports/sku_matrix.xlsx с листами:
       Матрица  — SKU, характеристики, бенчмарк и таргет по цене;
       Рынок    — карточки из выдачи, от них считается бенчмарк;
       Цвета    — рейтинг и негатив по цветам из отзывов конкурентов;
       Юнитка   — заготовка под расчёт маржи: жёлтые ячейки заполняете вы.

Бенчмарк и вся арифметика — формулы Excel. Поменяете таргет или данные —
пересчитается само.

Запуск:
    python scripts/build_matrix.py
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import load_config, resolve  # noqa: E402
from export_excel import detect_kind, detect_pieces, detect_size  # noqa: E402

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="2F4858")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")   # жёлтый: сюда вводите свои цифры
CALC_FILL = PatternFill("solid", fgColor="EFEAE3")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name=FONT, bold=True, size=13)
BODY = Font(name=FONT, size=10)
BOLD = Font(name=FONT, size=10, bold=True)
INPUT_FONT = Font(name=FONT, size=10, color="0000FF")  # синий: ручной ввод
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = "#,##0 ₽"
PERCENT = "0%"
PERCENT1 = "0.0%"
COUNT = "#,##0"
RATING = "0.00"

MATRIX_HEADERS = [
    ("Код SKU", 14), ("Категория", 20), ("Размер", 22), ("Комплектность", 15),
    ("Цвета", 30), ("Позиций", 9), ("Плотность, г/м²", 13), ("Состав", 24),
    ("Впитываемость", 13), ("Обработка", 15),
    ("Предметов в SKU", 10),
    ("Рынок: сопоставимых карточек", 13), ("Рынок: медиана", 13),
    ("Рынок: 25%", 11), ("Рынок: 75%", 11), ("Рынок: медиана за 1 шт", 14),
    ("Рынок: ср. отзывов", 12),
    ("Таргет, ₽", 11), ("Таргет за 1 шт", 12),
    ("К медиане", 11), ("К медиане за шт", 13),
    ("Роль SKU", 26), ("УТП для карточки", 46),
]

UNIT_HEADERS = [
    ("Код SKU", 14), ("Таргет, ₽", 11), ("Закупка за ед., ₽", 14),
    ("Доставка до склада, ₽", 14), ("Упаковка, ₽", 12), ("Себестоимость, ₽", 14),
    ("Комиссия WB, ₽", 13), ("Логистика WB, ₽", 13), ("Потери на возвратах, ₽", 14),
    ("Выручка нетто, ₽", 14), ("Маржа, ₽", 12), ("Маржа, %", 10),
    ("Мин. цена без убытка, ₽", 16),
]


def write_header(ws, headers: list[tuple[str, int]], row: int = 1) -> None:
    for i, (name, width) in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=i, value=name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[row].height = 32


def build_market(ws, df: pd.DataFrame) -> int:
    """Лист «Рынок»: выдача WB, от неё считается бенчмарк."""
    headers = [("Запрос", 24), ("Артикул", 13), ("Название", 50), ("Бренд", 16),
               ("Цена", 11), ("Рейтинг", 9), ("Отзывов", 10), ("Размер", 11),
               ("Штук", 8), ("Цена за шт", 12), ("Тип", 12)]
    write_header(ws, headers)

    # один товар попадает в выдачу по нескольким запросам: без этого
    # он считался бы несколько раз и перекашивал медиану
    df = df.drop_duplicates("product_id")

    for i, (_, item) in enumerate(df.iterrows(), start=2):
        title = str(item.get("title", ""))
        pieces = detect_pieces(title)
        values = [item.get("query", ""), item.get("product_id", ""), title,
                  item.get("brand", ""), item.get("price"), item.get("rating"),
                  item.get("review_count"), detect_size(title), pieces]
        for j, value in enumerate(values, start=1):
            cell = ws.cell(row=i, column=j)
            cell.value = None if (isinstance(value, float) and pd.isna(value)) else value
            cell.font = BODY
            cell.border = BORDER
        # цена за штуку — формулой, чтобы пересчитывалась при правке цены
        per_piece = ws.cell(row=i, column=10, value=f"=IFERROR($E{i}/$I{i},\"\")")
        per_piece.font = BODY
        per_piece.border = BORDER
        per_piece.number_format = MONEY
        kind = ws.cell(row=i, column=11, value=detect_kind(title))
        kind.font = BODY
        kind.border = BORDER
        ws.cell(row=i, column=5).number_format = MONEY
        ws.cell(row=i, column=7).number_format = COUNT

    last = len(df) + 1
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{last}"
    return last


def build_matrix(ws, matrix: pd.DataFrame, market_last: int) -> int:
    ws.cell(row=1, column=1, value="Матрица SKU под запуск линейки").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="Бенчмарк считается формулами от листа «Рынок» (выдача Wildberries). "
                  "Таргет и характеристики правятся в data/sku_matrix.csv или прямо здесь."
            ).font = NOTE
    write_header(ws, MATRIX_HEADERS, row=3)

    size_rng = f"Рынок!$H$2:$H${market_last}"
    price_rng = f"Рынок!$E$2:$E${market_last}"
    reviews_rng = f"Рынок!$G$2:$G${market_last}"
    pieces_rng = f"Рынок!$I$2:$I${market_last}"
    per_piece_rng = f"Рынок!$J$2:$J${market_last}"
    kind_rng = f"Рынок!$K$2:$K${market_last}"

    row = 4
    for _, item in matrix.iterrows():
        colors = str(item.get("colors", ""))
        size = str(item.get("size", ""))
        # у халатов в поле размера лежит ростовка («S/M, L/XL, XXL»), и каждая
        # ростовка — отдельная товарная позиция, поэтому цвета умножаются на них
        runs = [s for s in size.split(",") if s.strip()]
        size_runs = len(runs) if len(runs) > 1 else 1
        positions = len([c for c in colors.split(",") if c.strip()]) * size_runs
        bench = str(item.get("benchmark_size", "")).strip()

        values = [
            item.get("code", ""), item.get("category", ""), item.get("size", ""),
            item.get("pack", ""), colors, positions, item.get("density", ""),
            item.get("composition", ""), item.get("absorbency", ""), item.get("finish", ""),
        ]
        for j, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=j, value=value)
            cell.font = BODY
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=j in (5, 8, 3), vertical="top")

        # Бенчмарк сравнивает сопоставимое: одиночный товар — с одиночными,
        # набор — с наборами. Без этого медиана 50х80 мешала одиночные по 1416 ₽
        # с наборами по 501 ₽ и не значила ничего.
        pieces = max(int(float(item.get("pieces", 1) or 1)), 1)
        pack_cond = f"({pieces_rng}=1)" if pieces == 1 else f"({pieces_rng}>1)"
        pack_criteria = "1" if pieces == 1 else ">1"
        # тип товара тоже важен: в размере 50х80 полно ковриков, и без этого
        # фильтра они попадали в медиану полотенец
        kind = str(item.get("benchmark_kind", "") or "Полотенце").strip()
        # У халатов размер в названии не пишут (там ростовки), поэтому для них
        # бенчмарк строится по типу товара и комплектности, без размера.
        size_cond = f'({size_rng}="{bench}")*' if bench else ""
        mask = f'{size_cond}{pack_cond}*({kind_rng}="{kind}")'

        ws.cell(row=row, column=11, value=pieces).font = BODY
        ws.cell(row=row, column=11).number_format = COUNT

        size_criteria = ([size_rng, f'"{bench}"'] if bench else [])
        count_args = ", ".join(size_criteria + [pieces_rng, f'"{pack_criteria}"',
                                               kind_rng, f'"{kind}"'])
        ws.cell(row=row, column=12, value=f"=COUNTIFS({count_args})")
        for col, formula in (
            (13, f'=IFERROR(MEDIAN(IF({mask},{price_rng})),"нет данных")'),
            (14, f'=IFERROR(QUARTILE(IF({mask},{price_rng}),1),"нет данных")'),
            (15, f'=IFERROR(QUARTILE(IF({mask},{price_rng}),3),"нет данных")'),
            (16, f'=IFERROR(MEDIAN(IF({mask},{per_piece_rng})),"нет данных")'),
        ):
            letter = get_column_letter(col)
            ws.cell(row=row, column=col, value=ArrayFormula(f"{letter}{row}", formula))
        ws.cell(row=row, column=17,
                value=f'=IFERROR(AVERAGEIFS({reviews_rng}, {count_args}),"нет данных")')

        target = ws.cell(row=row, column=18, value=float(item.get("target_price", 0) or 0))
        target.font = INPUT_FONT
        target.fill = INPUT_FILL
        ws.cell(row=row, column=19, value=f'=IFERROR($R{row}/$K{row},"")')
        ws.cell(row=row, column=20, value=f'=IFERROR($R{row}/$M{row}-1,"—")')
        ws.cell(row=row, column=21, value=f'=IFERROR($S{row}/$P{row}-1,"—")')

        for col in range(12, 22):
            cell = ws.cell(row=row, column=col)
            cell.border = BORDER
            if col != 18:
                cell.font = BODY
        ws.cell(row=row, column=12).number_format = COUNT
        for col in (13, 14, 15, 16, 18, 19):
            ws.cell(row=row, column=col).number_format = MONEY
        ws.cell(row=row, column=17).number_format = COUNT
        for col in (20, 21):
            ws.cell(row=row, column=col).number_format = PERCENT

        for col, value in ((22, item.get("role", "")), (23, item.get("usp", ""))):
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = BODY
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    last = row - 1
    ws.cell(row=row, column=1, value="Итого").font = BOLD
    ws.cell(row=row, column=6, value=f"=SUM($F$4:$F${last})").font = BOLD
    ws.cell(row=row, column=18, value=f"=AVERAGE($R$4:$R${last})").font = BOLD
    ws.cell(row=row, column=18).number_format = MONEY
    ws.cell(row=row, column=1).fill = CALC_FILL

    ws.cell(row=row + 2, column=1,
            value="Жёлтые ячейки — ваши значения. Бенчмарк берёт только сопоставимые "
                  "карточки: тот же тип товара, тот же размер, та же комплектность — "
                  "одиночный товар с одиночными, набор с наборами. "
                  "Колонки 25% и 75% — коридор, где лежит основная масса; «за 1 шт» "
                  "приводит наборы разной комплектности к одной базе.").font = NOTE
    ws.cell(row=row + 3, column=1,
            value="«нет данных» в бенчмарке означает, что размер не попал в сбор: "
                  "запустите анализ заново — запросы по коврикам и салфеткам уже добавлены."
            ).font = NOTE
    ws.freeze_panes = "B4"
    return last


def build_colors(ws, reviews: pd.DataFrame) -> None:
    ws.cell(row=1, column=1, value="Цвета: что показывают отзывы конкурентов").font = TITLE_FONT
    if reviews.empty:
        ws.cell(row=3, column=1,
                value="Отзывы не собраны. Запустите scripts/collect_reviews.py, "
                      "и лист заполнится.").font = BODY
        return

    ws.cell(row=2, column=1,
            value="Только карточки с одним цветом — иначе в поле цвета попадает весь "
                  "список оттенков товара и сигнал смазывается.").font = NOTE
    headers = [("Цвет", 22), ("Отзывов", 10), ("Средний рейтинг", 14),
               ("Негатив, %", 12), ("Жалобы на линьку, %", 16),
               ("Жалобы «тонкое», %", 16), ("Вывод", 34)]
    write_header(ws, headers, row=3)

    single = reviews[(reviews["color_clean"] != "") & (~reviews["color_clean"].str.contains(","))]
    row = 4
    for color, group in single.groupby("color_clean"):
        if len(group) < 25:
            continue
        negative = (group["rating"] <= 3).mean() * 100
        fade = group["text_all"].str.contains("линя|красит|полиня|выцвет").mean() * 100
        thin = group["text_all"].str.contains("тонк|просвечив").mean() * 100
        verdict = ("брать в линейку" if negative < 13 and fade < 2
                   else "не брать" if negative > 25 or fade > 5
                   else "можно, с контролем качества")

        for j, value in enumerate([color.capitalize(), len(group),
                                   round(float(group["rating"].mean()), 2),
                                   round(negative, 1) / 100, round(fade, 1) / 100,
                                   round(thin, 1) / 100, verdict], start=1):
            cell = ws.cell(row=row, column=j, value=value)
            cell.font = BODY
            cell.border = BORDER
        ws.cell(row=row, column=3).number_format = RATING
        for col in (4, 5, 6):
            ws.cell(row=row, column=col).number_format = PERCENT1
        row += 1

    ws.cell(row=row + 1, column=1,
            value=f"Источник: {len(single)} отзывов на одноцветных карточках "
                  f"из {len(reviews)} собранных.").font = NOTE
    ws.auto_filter.ref = f"A3:G{row - 1}"


def build_unit(ws, matrix: pd.DataFrame, matrix_last: int) -> None:
    ws.cell(row=1, column=1, value="Юнит-экономика: заготовка").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="Заполните жёлтые ячейки — остальное посчитается. Значения-ставки "
                  "внизу тоже правятся.").font = NOTE

    # блок общих ставок
    ws.cell(row=4, column=1, value="Общие ставки").font = BOLD
    rates = [("Комиссия WB, %", 0.23), ("Логистика WB за единицу, ₽", 60),
             ("Доля возвратов, %", 0.08), ("Налог с оборота, %", 0.06)]
    for i, (label, value) in enumerate(rates, start=5):
        ws.cell(row=i, column=1, value=label).font = BODY
        cell = ws.cell(row=i, column=2, value=value)
        cell.font = INPUT_FONT
        cell.fill = INPUT_FILL
        cell.border = BORDER
        cell.number_format = PERCENT if "%" in label else MONEY
    ws.cell(row=9, column=3,
            value="Ставки проверьте в личном кабинете WB — они зависят от категории "
                  "и склада.").font = NOTE

    header_row = 11
    write_header(ws, UNIT_HEADERS, row=header_row)

    row = header_row + 1
    for i, (_, item) in enumerate(matrix.iterrows()):
        matrix_row = 4 + i
        ws.cell(row=row, column=1, value=f"=Матрица!$A${matrix_row}").font = BODY
        ws.cell(row=row, column=2, value=f"=Матрица!$R${matrix_row}").font = BODY
        ws.cell(row=row, column=2).number_format = MONEY

        for col in (3, 4, 5):  # закупка, доставка, упаковка — вводит пользователь
            cell = ws.cell(row=row, column=col, value=0)
            cell.font = INPUT_FONT
            cell.fill = INPUT_FILL
            cell.number_format = MONEY

        ws.cell(row=row, column=6, value=f"=SUM($C{row}:$E{row})").number_format = MONEY
        ws.cell(row=row, column=7, value=f"=$B{row}*$B$5").number_format = MONEY
        ws.cell(row=row, column=8, value="=$B$6").number_format = MONEY
        ws.cell(row=row, column=9, value=f"=($F{row}+$H{row})*$B$7").number_format = MONEY
        ws.cell(row=row, column=10, value=f"=$B{row}*(1-$B$8)").number_format = MONEY
        ws.cell(row=row, column=11,
                value=f"=$J{row}-$F{row}-$G{row}-$H{row}-$I{row}").number_format = MONEY
        ws.cell(row=row, column=12, value=f'=IFERROR($K{row}/$B{row},"")').number_format = PERCENT
        ws.cell(row=row, column=13,
                value=f'=IFERROR(($F{row}+$H{row}+($F{row}+$H{row})*$B$7)'
                      f'/((1-$B$8)-$B$5),"")').number_format = MONEY

        for col in range(1, len(UNIT_HEADERS) + 1):
            ws.cell(row=row, column=col).border = BORDER
            if col not in (3, 4, 5):
                ws.cell(row=row, column=col).font = BODY
        row += 1

    last = row - 1
    ws.cell(row=row, column=1, value="Итого / среднее").font = BOLD
    ws.cell(row=row, column=11, value=f"=SUM($K${header_row + 1}:$K${last})").font = BOLD
    ws.cell(row=row, column=11).number_format = MONEY
    ws.cell(row=row, column=12, value=f"=AVERAGE($L${header_row + 1}:$L${last})").font = BOLD
    ws.cell(row=row, column=12).number_format = PERCENT

    ws.cell(row=row + 2, column=1, value="Как читать:").font = BOLD
    for i, text in enumerate([
        "Себестоимость — ваши деньги за товар до маркетплейса (закупка + доставка + упаковка).",
        "Потери на возвратах — доля возвратов от себестоимости и логистики: товар вернулся, "
        "расходы остались.",
        "Выручка нетто — цена минус налог с оборота.",
        "Мин. цена без убытка — ниже неё продажа уходит в минус при текущих ставках.",
        "Реклама и продвижение сюда не заложены: добавьте их отдельной строкой, когда "
        "появятся ставки.",
    ], start=row + 3):
        ws.cell(row=i, column=1, value=text).font = NOTE


def prepare_reviews(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    if df.empty:
        return df
    if "review_id" in df.columns:
        df = df.drop_duplicates("review_id")
    df["rating"] = pd.to_numeric(df.get("rating"), errors="coerce")
    for col in ("text", "pros", "cons", "color"):
        if col not in df.columns:
            df[col] = ""
    df["text_all"] = (df["text"] + " " + df["pros"] + " " + df["cons"]).str.lower()
    # в поле цвета WB держит характеристику карточки: "цвета · размеры"
    df["color_clean"] = df["color"].str.split("·").str[0].str.strip().str.lower()
    return df


def main() -> None:
    parser = argparse.ArgumentParser(description="Матрица SKU в Excel")
    parser.add_argument("--config", default="config.yaml")
    parser.add_argument("--out", default=None)
    args = parser.parse_args()

    cfg = load_config(args.config)
    out_cfg = cfg.get("output") or {}
    reports_dir = resolve(out_cfg.get("reports_dir", "reports"))
    products_path = resolve(out_cfg.get("products_file", "data/products.csv"))
    reviews_path = resolve(out_cfg.get("reviews_file", "data/reviews.csv"))
    matrix_path = resolve(out_cfg.get("matrix_file", "data/sku_matrix.csv"))
    out_path = Path(args.out) if args.out else reports_dir / "Матрица_SKU.xlsx"

    if not matrix_path.exists():
        sys.exit(f"Нет файла матрицы: {matrix_path}")
    matrix = pd.read_csv(matrix_path, dtype=str, keep_default_na=False)
    if matrix.empty:
        sys.exit("Матрица пуста — добавьте строки в data/sku_matrix.csv")

    products = pd.DataFrame()
    if products_path.exists():
        products = pd.read_csv(products_path, dtype=str, keep_default_na=False)
        for col in ("price", "rating", "review_count"):
            products[col] = pd.to_numeric(products.get(col, ""), errors="coerce")

    reviews = prepare_reviews(reviews_path)

    wb = Workbook()
    ws_matrix = wb.active
    ws_matrix.title = "Матрица"
    ws_market = wb.create_sheet("Рынок")
    ws_colors = wb.create_sheet("Цвета")
    ws_unit = wb.create_sheet("Юнитка")

    market_last = build_market(ws_market, products) if not products.empty else 1
    matrix_last = build_matrix(ws_matrix, matrix, market_last)
    build_colors(ws_colors, reviews)
    build_unit(ws_unit, matrix, matrix_last)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    print(f"Готово: {out_path}")
    print(f"  SKU в матрице: {len(matrix)}")
    print(f"  Карточек рынка для бенчмарка: {max(market_last - 1, 0)}")
    print(f"  Отзывов для среза по цветам: {len(reviews)}")


if __name__ == "__main__":
    main()
