"""Полный прогон: сбор -> анализ -> отчёты -> Excel и Word.

Запускает все скрипты по очереди и в конце показывает, что получилось.
Если какой-то шаг падает, следующие всё равно выполняются: например, когда
отзывы собрать не удалось, отчёты по выдаче всё равно соберутся.

Запуск:
    python scripts/run_all.py
    python scripts/run_all.py --skip-reviews   # без сбора отзывов (быстрее)
    python scripts/run_all.py --skip-collect   # не ходить в сеть, пересчитать на том, что есть
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import PROJECT_ROOT, load_config, resolve  # noqa: E402

# модуль -> что ставить, если его нет
REQUIRED_MODULES = {
    "pandas": "pandas",
    "yaml": "pyyaml",
    "requests": "requests",
    "openpyxl": "openpyxl",
    "docx": "python-docx",
    "curl_cffi": "curl_cffi",
}

LINE = "=" * 62


def ensure_packages() -> None:
    """Доставляет недостающие библиотеки, чтобы не ловить ошибку на середине."""
    missing = []
    for module, package in REQUIRED_MODULES.items():
        try:
            __import__(module)
        except ImportError:
            missing.append(package)

    if not missing:
        return

    print(f"Не хватает библиотек: {', '.join(missing)}. Ставлю, это разовая операция...")
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "--quiet", *missing],
        cwd=PROJECT_ROOT,
    )
    if result.returncode == 0:
        print("Библиотеки установлены.\n")
    else:
        print("Установить не удалось. Выполните вручную:")
        print(f"  {Path(sys.executable).name} -m pip install -r requirements.txt\n")


def run_step(number: int, total: int, title: str, script: str,
             args: list[str] | None = None) -> bool:
    """Выполняет один скрипт. Возвращает True, если он отработал без ошибки."""
    print(f"\n{LINE}\n[{number}/{total}] {title}\n{LINE}")
    started = time.time()
    result = subprocess.run(
        [sys.executable, str(PROJECT_ROOT / "scripts" / script), *(args or [])],
        cwd=PROJECT_ROOT,
    )
    took = time.time() - started
    if result.returncode == 0:
        print(f"-- готово за {took:.0f} сек")
        return True
    print(f"-- шаг завершился с ошибкой (код {result.returncode}), иду дальше")
    return False


def report_files(cfg: dict) -> list[tuple[str, Path]]:
    out_cfg = cfg.get("output") or {}
    reports_dir = resolve(out_cfg.get("reports_dir", "reports"))
    return [
        ("Товары одной таблицей", resolve(out_cfg.get("products_file", "data/products.csv"))),
        ("Отзывы конкурентов", resolve(out_cfg.get("reviews_file", "data/reviews.csv"))),
        ("Excel со всеми срезами", reports_dir / "marketplace_report.xlsx"),
        ("Word с выводами", reports_dir / "Отчёт_по_рынку.docx"),
        ("Матрица SKU", reports_dir / "Матрица_SKU.xlsx"),
        ("Выводы под запуск", reports_dir / "launch_recommendations.md"),
        ("Что говорят покупатели", reports_dir / "reviews_summary.md"),
    ]


def main() -> None:
    parser = argparse.ArgumentParser(description="Полный прогон пайплайна")
    parser.add_argument("--skip-collect", action="store_true",
                        help="не собирать выдачу заново")
    parser.add_argument("--skip-reviews", action="store_true",
                        help="не собирать отзывы")
    parser.add_argument("--top-n", type=int, default=None,
                        help="сколько карточек брать по запросу")
    parser.add_argument("--reviews-top", type=int, default=None,
                        help="по скольким карточкам собирать отзывы")
    args = parser.parse_args()

    print(f"{LINE}\nАНАЛИЗ МАРКЕТПЛЕЙСА — ПОЛНЫЙ ПРОГОН\n{LINE}")
    ensure_packages()

    steps: list[tuple[str, str, list[str]]] = []
    if not args.skip_collect:
        steps.append(("Сбор карточек с Wildberries", "collect_wb.py",
                      ["--top-n", str(args.top_n)] if args.top_n else []))
    if not args.skip_reviews:
        steps.append(("Сбор отзывов конкурентов", "collect_reviews.py",
                      ["--top", str(args.reviews_top)] if args.reviews_top else []))
    steps += [
        ("Аналитика по товарам", "analyze_products.py", []),
        ("Разбор отзывов", "analyze_reviews.py", []),
        ("Отчёты в Markdown", "generate_report.py", []),
        ("Выгрузка в Excel", "export_excel.py", []),
        ("Выгрузка в Word", "export_docx.py", []),
        ("Матрица SKU и заготовка юнитки", "build_matrix.py", []),
    ]

    failed = []
    for i, (title, script, extra) in enumerate(steps, start=1):
        if not run_step(i, len(steps), title, script, extra):
            failed.append(title)

    cfg = load_config()
    print(f"\n{LINE}\nИТОГ\n{LINE}")
    missing = []
    for label, path in report_files(cfg):
        exists = path.exists()
        print(f"  [{'есть ' if exists else 'нет  '}] {label}: {path.name}")
        if not exists:
            missing.append(label)

    if failed:
        print("\nШаги с ошибками: " + ", ".join(failed))
        print("Чаще всего дело в доступе к Wildberries — проверьте:")
        print(f"  {Path(sys.executable).name} scripts/check_wb.py")
    elif missing:
        print("\nОшибок не было, но часть файлов не собралась: " + ", ".join(missing) + ".")
    else:
        print("\nВсё собралось.")

    print("\nСмотреть результат: папка reports (файлы .xlsx и .docx).")


if __name__ == "__main__":
    main()
