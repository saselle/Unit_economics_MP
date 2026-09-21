"""Разбор собранных отзывов: за что хвалят и на что жалуются.

Вход:  data/reviews.csv
Выход: reports/review_themes.csv     — темы с частотой, средним рейтингом и цитатой;
       reports/review_by_product.csv — срез по каждой карточке;
       reports/reviews_summary.md    — что с этим делать при запуске.

Темы и ключевые слова лежат в config.yaml (reviews.themes) — правятся без кода.

Запуск:
    python scripts/analyze_reviews.py
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import load_config, resolve, save_document  # noqa: E402

THEME_COLUMNS = ["тема", "тип", "упоминаний", "доля_отзывов_%", "средний_рейтинг",
                 "негативных_%", "пример"]
PRODUCT_COLUMNS = ["product_id", "brand", "product_title", "отзывов", "средний_рейтинг",
                   "негативных_%", "главные_жалобы", "главные_плюсы"]

STOPWORDS = {
    "очень", "полотенце", "полотенца", "полотенец", "полотенца", "просто", "всё", "все",
    "это", "как", "что", "для", "или", "тоже", "меня", "мне", "они", "оно", "была", "были",
    "буду", "ещё", "еще", "себе", "свой", "свои", "после", "деньги", "цена", "заказ",
    "товар", "спасибо", "магазин", "продавец", "доставка", "качество", "хорошо", "отлично",
}
WORD_RE = re.compile(r"[а-яёa-z]{4,}", re.IGNORECASE)


def load_reviews(path: Path) -> pd.DataFrame:
    if not path.exists():
        print(f"Нет файла {path}. Сначала запустите scripts/collect_reviews.py")
        return pd.DataFrame()

    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    if df.empty:
        return df

    df["rating"] = pd.to_numeric(df.get("rating"), errors="coerce")
    for col in ("text", "pros", "cons", "brand", "product_title", "product_id"):
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str)

    # в одном поле удобнее искать: WB часто кладёт суть в «минусы», а не в текст
    df["full_text"] = (df["text"] + " " + df["pros"] + " " + df["cons"]).str.lower()

    df = drop_foreign(df)
    df = drop_duplicates(df)
    return df


def drop_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Убирает повторы: один отзыв мог прийти по нескольким артикулам сразу."""
    if "review_id" not in df.columns:
        return df
    before = len(df)
    df = df.drop_duplicates(subset=["review_id"])
    if len(df) < before:
        print(f"  убрано повторов отзыва: {before - len(df)} "
              "(один товар продаётся под несколькими артикулами)")
    return df


def drop_foreign(df: pd.DataFrame) -> pd.DataFrame:
    """Страховка от чужих отзывов: артикул в отзыве должен совпадать с карточкой.

    Проверяем ещё раз на этапе анализа — вдруг данные собраны старой версией
    скрипта, которая такие отзывы пропускала.
    """
    if "review_nm_id" not in df.columns:
        print("  ВНИМАНИЕ: в файле нет колонки review_nm_id — он собран старой версией")
        print("  collect_reviews.py. Удалите data/reviews.csv и соберите отзывы заново.")
        return df

    known = df["review_nm_id"].astype(str).str.strip() != ""
    same = df["review_nm_id"].astype(str).str.strip() == df["product_id"].astype(str).str.strip()
    foreign = known & ~same
    if foreign.any():
        print(f"  отброшено чужих отзывов: {int(foreign.sum())}")
        df = df[~foreign]
    return df


def match_themes(text: str, themes: dict) -> list[tuple[str, str]]:
    """Возвращает пары (тема, тип) для всех сработавших ключевых слов."""
    found = []
    for kind, group in themes.items():
        for theme, keywords in (group or {}).items():
            if any(str(k).lower() in text for k in (keywords or [])):
                found.append((theme, kind))
    return found


def build_themes(df: pd.DataFrame, themes: dict) -> pd.DataFrame:
    rows = []
    total = len(df)
    for kind, group in themes.items():
        for theme, keywords in (group or {}).items():
            mask = df["full_text"].apply(
                lambda t, k=keywords: any(str(w).lower() in t for w in (k or []))
            )
            hits = df[mask]
            if hits.empty:
                continue
            negative = (hits["rating"] <= 3).sum()
            example = ""
            candidates = hits[hits["full_text"].str.len() > 40]
            if not candidates.empty:
                example = candidates.iloc[0]["full_text"][:180].strip()
            rows.append({
                "тема": theme,
                "тип": kind,
                "упоминаний": len(hits),
                "доля_отзывов_%": round(len(hits) / total * 100, 1),
                "средний_рейтинг": round(float(hits["rating"].mean()), 2)
                if hits["rating"].notna().any() else "",
                "негативных_%": round(negative / len(hits) * 100, 1),
                "пример": example,
            })
    result = pd.DataFrame(rows, columns=THEME_COLUMNS)
    if not result.empty:
        result = result.sort_values(["тип", "упоминаний"], ascending=[True, False])
    return result


def build_by_product(df: pd.DataFrame, themes: dict) -> pd.DataFrame:
    rows = []
    for product_id, group in df.groupby("product_id", sort=False):
        complaints: Counter[str] = Counter()
        praises: Counter[str] = Counter()
        for text in group["full_text"]:
            for theme, kind in match_themes(text, themes):
                if kind == "жалобы":
                    complaints[theme] += 1
                elif kind == "похвала":
                    praises[theme] += 1
        rows.append({
            "product_id": product_id,
            "brand": group["brand"].iloc[0],
            "product_title": group["product_title"].iloc[0],
            "отзывов": len(group),
            "средний_рейтинг": round(float(group["rating"].mean()), 2)
            if group["rating"].notna().any() else "",
            "негативных_%": round((group["rating"] <= 3).sum() / len(group) * 100, 1),
            "главные_жалобы": ", ".join(f"{t} ({n})" for t, n in complaints.most_common(3)),
            "главные_плюсы": ", ".join(f"{t} ({n})" for t, n in praises.most_common(3)),
        })
    return pd.DataFrame(rows, columns=PRODUCT_COLUMNS)


def frequent_words(texts, limit: int = 15) -> list[tuple[str, int]]:
    counter: Counter[str] = Counter()
    for text in texts:
        counter.update({w for w in WORD_RE.findall(str(text)) if w not in STOPWORDS})
    return counter.most_common(limit)


def build_summary(df: pd.DataFrame, themes_df: pd.DataFrame, products_df: pd.DataFrame) -> str:
    negative = df[df["rating"] <= 3]
    positive = df[df["rating"] >= 4]
    complaints = themes_df[themes_df["тип"] == "жалобы"].head(8)
    praises = themes_df[themes_df["тип"] == "похвала"].head(8)

    out = [
        "# Что говорят покупатели",
        "",
        f"- Отзывов разобрано: **{len(df)}** по {df['product_id'].nunique()} карточкам",
        f"- Средний рейтинг: **{df['rating'].mean():.2f}**",
        f"- Негативных (1–3 звезды): **{len(negative)}** ({len(negative) / max(len(df), 1) * 100:.1f}%)",
        "",
        "## На что жалуются",
        "",
    ]

    if complaints.empty:
        out.append("_Жалоб по словарю не нашлось — проверьте ключевые слова в config.yaml._")
    else:
        out.append("| Тема | Упоминаний | Доля отзывов | Средний рейтинг |")
        out.append("|---|---|---|---|")
        for _, row in complaints.iterrows():
            out.append(f"| {row['тема']} | {row['упоминаний']} | {row['доля_отзывов_%']}% | "
                       f"{row['средний_рейтинг']} |")
    out += ["", "## За что хвалят", ""]

    if praises.empty:
        out.append("_Похвалы по словарю не нашлось._")
    else:
        out.append("| Тема | Упоминаний | Доля отзывов | Средний рейтинг |")
        out.append("|---|---|---|---|")
        for _, row in praises.iterrows():
            out.append(f"| {row['тема']} | {row['упоминаний']} | {row['доля_отзывов_%']}% | "
                       f"{row['средний_рейтинг']} |")

    out += ["", "## Частые слова в негативных отзывах", ""]
    words = frequent_words(negative["full_text"]) if not negative.empty else []
    out.append(", ".join(f"`{w}` ({n})" for w, n in words) if words else "_нет данных_")

    out += ["", "## Частые слова в позитивных отзывах", ""]
    words = frequent_words(positive["full_text"]) if not positive.empty else []
    out.append(", ".join(f"`{w}` ({n})" for w, n in words) if words else "_нет данных_")

    out += ["", "## Карточки с самой высокой долей негатива", "",
            "Смотреть, что там пошло не так, и не повторять:", ""]
    if not products_df.empty:
        worst = products_df.sort_values("негативных_%", ascending=False).head(5)
        out.append("| Бренд | Товар | Отзывов | Негатива | Главные жалобы |")
        out.append("|---|---|---|---|---|")
        for _, row in worst.iterrows():
            out.append(f"| {row['brand']} | {str(row['product_title'])[:44]} | {row['отзывов']} | "
                       f"{row['негативных_%']}% | {row['главные_жалобы'] or '—'} |")

    out += [
        "",
        "## Как этим пользоваться",
        "",
        "1. **Топ жалоб — это ТЗ поставщику.** Проверяйте каждый пункт на образце до заказа партии.",
        "2. **Топ жалоб — это и текст карточки.** Прямо снимайте возражения: «не линяет после "
        "50 стирок», «плотность 500 г/м²».",
        "3. **Топ похвал — слова покупателя.** Их стоит использовать в описании дословно.",
        "4. **Частые слова** годятся в характеристики и в инфографику на фото.",
        "",
        "---",
        "",
        "_Отчёт собран автоматически: `python scripts/analyze_reviews.py`._",
    ]
    return "\n".join(out) + "\n"


def save(df: pd.DataFrame, path: Path, columns: list[str]) -> None:
    out = df if not df.empty else pd.DataFrame(columns=columns)
    path = save_document(
        path, lambda p: out.to_csv(p, index=False, encoding="utf-8-sig"))
    print(f"  {path}  ({len(df)} строк)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Анализ отзывов")
    parser.add_argument("--config", default="config.yaml")
    args = parser.parse_args()

    cfg = load_config(args.config)
    out_cfg = cfg.get("output") or {}
    themes = (cfg.get("reviews") or {}).get("themes") or {}
    reviews_path = resolve(out_cfg.get("reviews_file", "data/reviews.csv"))
    reports_dir = resolve(out_cfg.get("reports_dir", "reports"))

    df = load_reviews(reviews_path)
    if df.empty:
        print("Отзывов нет — создаю пустые отчёты со схемой.")
        save(pd.DataFrame(), reports_dir / "review_themes.csv", THEME_COLUMNS)
        save(pd.DataFrame(), reports_dir / "review_by_product.csv", PRODUCT_COLUMNS)
        return

    print(f"Отзывов: {len(df)}, карточек: {df['product_id'].nunique()}")
    themes_df = build_themes(df, themes)
    products_df = build_by_product(df, themes)

    print("Сохраняю:")
    save(themes_df, reports_dir / "review_themes.csv", THEME_COLUMNS)
    save(products_df, reports_dir / "review_by_product.csv", PRODUCT_COLUMNS)

    summary = build_summary(df, themes_df, products_df)
    summary_path = save_document(reports_dir / "reviews_summary.md",
                                 lambda p: p.write_text(summary, encoding="utf-8"))
    print(f"  {summary_path}")


if __name__ == "__main__":
    main()
