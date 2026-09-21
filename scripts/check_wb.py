"""Диагностика доступа к Wildberries.

Скрипт перебирает версии поискового адреса, наборы параметров и заголовков,
и показывает, какой вариант отвечает. Нужен, когда collect_wb.py получает 403
и уходит в fallback на мок-данные.

Запуск:
    python scripts/check_wb.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import curl_requests, make_session  # noqa: E402

QUERY = "полотенце"

VERSIONS = ["v18", "v17", "v16", "v15", "v14", "v13", "v5", "v4"]

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "*/*",
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Referer": "https://www.wildberries.ru/",
}

HEADER_SETS = {
    "браузерные": BROWSER_HEADERS,
    "пустые": {},
}

PARAM_SETS = {
    "новый": {
        "ab_testing": "false",
        "appType": 1,
        "curr": "rub",
        "dest": -1257786,
        "hide_dtype": 13,
        "lang": "ru",
        "query": QUERY,
        "resultset": "catalog",
        "sort": "popular",
        "spp": 30,
        "suppressSpellcheck": "false",
    },
    "минимум": {
        "appType": 1,
        "curr": "rub",
        "dest": -1257786,
        "query": QUERY,
        "resultset": "catalog",
        "sort": "popular",
    },
}


def count_products(payload) -> int:
    if not isinstance(payload, dict):
        return 0
    data = payload.get("data")
    if isinstance(data, dict) and isinstance(data.get("products"), list):
        return len(data["products"])
    if isinstance(payload.get("products"), list):
        return len(payload["products"])
    return 0


def check_site(session) -> None:
    """Открывается ли сам сайт WB — отличает блокировку по IP от смены адреса API."""
    try:
        response = session.get(
            "https://www.wildberries.ru/", headers=BROWSER_HEADERS, timeout=15
        )
        print(f"Доступ к сайту wildberries.ru: {response.status_code}")
        if response.status_code == 403:
            print("  !! Сайт отдаёт 403 — скорее всего включён VPN или зарубежный IP.")
            print("     Выключите VPN и запустите скрипт заново.")
    except Exception as exc:
        print(f"Доступ к сайту wildberries.ru: не открылся ({type(exc).__name__})")
        print("  !! Похоже, интернет не пускает к WB вообще (фаервол, прокси, VPN).")
    print()


def main() -> None:
    print("Проверка доступа к Wildberries\n" + "=" * 46)
    session, engine = make_session()
    print(f"Движок запросов: {engine}")
    if curl_requests is None:
        print("  Подсказка: маскировка под Chrome не установлена.")
        print("  Поставьте её командой:  py -m pip install curl_cffi")
        print("  и запустите этот скрипт заново — WB часто пропускает только её.\n")
    check_site(session)

    total = len(VERSIONS) * len(PARAM_SETS) * len(HEADER_SETS)
    print(f"Перебираю {total} вариантов запроса, это примерно минуту.\n")

    winner = None
    for version in VERSIONS:
        url = f"https://search.wb.ru/exactmatch/ru/common/{version}/search"
        for params_name, params in PARAM_SETS.items():
            for headers_name, headers in HEADER_SETS.items():
                label = f"{version:<4} | параметры {params_name:<8} | заголовки {headers_name:<10}"
                try:
                    response = session.get(url, params=params, headers=headers, timeout=15)
                    status = response.status_code
                    if status == 200:
                        try:
                            found = count_products(response.json())
                        except ValueError:
                            found = 0
                        print(f"{label} -> 200, товаров: {found}")
                        if found and winner is None:
                            winner = (version, params_name, headers_name, found)
                    else:
                        print(f"{label} -> {status}")
                except Exception as exc:
                    print(f"{label} -> ошибка {type(exc).__name__}")
                time.sleep(0.3)
        if winner:
            break

    print("\n" + "=" * 46)
    if winner:
        version, params_name, headers_name, found = winner
        print(f"РАБОТАЕТ: {version}, параметры «{params_name}», заголовки «{headers_name}», "
              f"товаров {found}")
        print("Пришлите эту строку — я поправлю сбор под неё.")
    else:
        print("Ни один вариант не сработал.")
        if curl_requests is None:
            print("1) Поставьте маскировку под Chrome:  py -m pip install curl_cffi")
            print("   и запустите  py scripts\\check_wb.py  ещё раз — это чаще всего помогает.")
        else:
            print("1) Маскировка под Chrome уже стоит, но WB всё равно отказывает.")
            print("   Выключите VPN, если он включён, и попробуйте мобильный интернет.")
        print("2) Пришлите скриншот этого окна целиком, включая самые верхние строки.")
    sys.stdout.flush()


if __name__ == "__main__":
    main()
