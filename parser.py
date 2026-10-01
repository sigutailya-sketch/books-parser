"""
Парсер каталога books.toscrape.com с выгрузкой в Excel и Google Таблицу.

Примеры запуска (из папки проекта, с активированным .venv):
    python parser.py                              # весь каталог, все 50 категорий
    python parser.py --list-categories            # показать список категорий
    python parser.py -c Mystery -c Poetry         # только выбранные категории
    python parser.py -c Travel --details          # + количество на складе, UPC и описание
    python parser.py --pages 1 --delay 1          # по одной странице в категории, пауза 1 с
    python parser.py --gsheet                     # дополнительно записать в Google Таблицу

Все параметры: python parser.py --help
"""

import argparse
import difflib
import logging
import os
import sys
import time
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

from dotenv import load_dotenv

from exporters import save_excel, save_gsheet
from scraper import BooksScraper, SiteUnavailableError, get_gbp_rate

PROJECT_DIR = Path(__file__).parent
log = logging.getLogger("parser")


def setup_logging(verbose: bool) -> Path:
    """Лог в консоль (коротко) и в файл logs/parser.log (подробно, с датой). Файл не растёт бесконечно."""
    log_dir = PROJECT_DIR / "logs"
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / "parser.log"

    log.setLevel(logging.DEBUG)
    # Если вывод перенаправлен в файл, Windows-кодировка может не знать какой-то символ — не падаем, заменяем на «?»
    sys.stdout.reconfigure(errors="replace")
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.DEBUG if verbose else logging.INFO)
    console.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s", "%H:%M:%S"))

    file = RotatingFileHandler(log_file, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    file.setLevel(logging.DEBUG)
    file.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s"))

    log.addHandler(console)
    log.addHandler(file)
    return log_file


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Сбор каталога книг с books.toscrape.com в Excel / Google Таблицу.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Пример: python parser.py -c Mystery -c Poetry --details",
    )
    p.add_argument("-c", "--category", action="append", metavar="НАЗВАНИЕ",
                   help="категория (можно указать несколько раз); по умолчанию — все")
    p.add_argument("-p", "--pages", type=int, metavar="N",
                   help="максимум страниц в каждой категории (20 книг на странице)")
    p.add_argument("-d", "--delay", type=float, default=0.5, metavar="СЕК",
                   help="пауза между запросами, по умолчанию 0.5 с")
    p.add_argument("--details", action="store_true",
                   help="заходить на страницу каждой книги: количество, UPC, описание (дольше)")
    p.add_argument("-o", "--output", type=Path, metavar="ФАЙЛ",
                   help="куда сохранить .xlsx (по умолчанию output/books_ДАТА.xlsx)")
    p.add_argument("--gsheet", action="store_true",
                   help="записать результат в Google Таблицу (настройки в .env)")
    p.add_argument("--list-categories", action="store_true", help="показать категории и выйти")
    p.add_argument("-v", "--verbose", action="store_true", help="подробный вывод в консоль")
    args = p.parse_args()

    if args.pages is not None and args.pages < 1:
        p.error("--pages должно быть не меньше 1")
    if args.delay < 0:
        p.error("--delay не может быть отрицательным")
    return args


def pick_categories(all_categories: dict[str, str], wanted: list[str] | None) -> dict[str, str]:
    """Выбирает категории по названию без учёта регистра. Опечатка — подскажем похожие и выйдем."""
    if not wanted:
        return all_categories

    by_lower = {name.lower(): name for name in all_categories}
    selected = {}
    for name in wanted:
        real = by_lower.get(name.strip().lower())
        if not real:
            hints = difflib.get_close_matches(name, all_categories.keys(), n=3, cutoff=0.5)
            hint = f" Возможно, вы имели в виду: {', '.join(hints)}." if hints else ""
            log.error("Категория «%s» не найдена.%s Список: python parser.py --list-categories", name, hint)
            sys.exit(2)
        selected[real] = all_categories[real]
    return selected


def main() -> int:
    args = parse_args()
    load_dotenv(PROJECT_DIR / ".env")
    log_file = setup_logging(args.verbose)
    started = time.monotonic()

    scraper = BooksScraper(delay=args.delay)

    try:
        categories = scraper.get_categories()
    except SiteUnavailableError as e:
        log.error("%s\nПроверьте интернет и попробуйте позже.", e)
        return 1

    if args.list_categories:
        print(f"Категорий: {len(categories)}")
        for name in categories:
            print(" ", name)
        return 0

    selected = pick_categories(categories, args.category)
    log.info("Старт: категорий %d, страниц на категорию: %s, пауза %.1f с%s",
             len(selected), args.pages or "все", args.delay, ", с деталями" if args.details else "")

    # Сбор. Ctrl+C не теряет данные: сохраним всё, что успели собрать.
    books = []
    interrupted = False
    try:
        for name, url in selected.items():
            for book in scraper.scrape_category(name, url, args.pages):
                if args.details:
                    scraper.fill_details(book)
                books.append(book)
    except KeyboardInterrupt:
        interrupted = True
        log.warning("Остановлено пользователем — сохраняю собранное (%d книг)", len(books))

    if not books:
        log.error("Не собрано ни одной книги — файл не создан.")
        return 1

    rate = get_gbp_rate()
    elapsed = time.monotonic() - started

    output = args.output or PROJECT_DIR / "output" / f"books_{datetime.now():%Y-%m-%d_%H-%M}.xlsx"
    run_info = {
        "Дата сбора": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "Источник": "https://books.toscrape.com/",
        "Категории": ", ".join(selected) if args.category else f"все ({len(selected)})",
        "Страниц на категорию": args.pages or "все",
        "Режим": "с деталями (страница каждой книги)" if args.details else "каталог",
        "Книг собрано": len(books),
        "Страниц каталога": scraper.stats.pages,
        "HTTP-запросов": scraper.stats.requests,
        "Ошибок": len(scraper.stats.errors),
        "Курс £ (ЦБ РФ)": round(rate, 4) if rate else "не получен",
        "Время работы, с": round(elapsed, 1),
        "Статус": "прерван пользователем" if interrupted else "завершён",
    }

    path = save_excel(books, output, rate, run_info)
    log.info("Excel сохранён: %s", path.resolve())

    if args.gsheet:
        creds = os.getenv("GOOGLE_CREDENTIALS", "credentials.json")
        sheet_id = os.getenv("GSHEET_ID")
        if not sheet_id:
            log.error("Не задан GSHEET_ID в файле .env — выгрузка в Google Таблицу пропущена")
        elif not (PROJECT_DIR / creds).exists() and not Path(creds).exists():
            log.error("Не найден файл ключа сервисного аккаунта «%s» — см. README", creds)
        else:
            try:
                url = save_gsheet(books, rate, str(PROJECT_DIR / creds if not Path(creds).is_absolute() else creds), sheet_id)
                log.info("Google Таблица обновлена: %s", url)
            except Exception as e:  # ошибки Google API разнообразны — показываем понятное сообщение
                log.error("Не удалось записать в Google Таблицу: %s", e)

    log.info("Готово: %d книг за %.0f с, ошибок: %d. Подробный лог: %s",
             len(books), elapsed, len(scraper.stats.errors), log_file)
    for err in scraper.stats.errors[:10]:
        log.info("  ошибка: %s", err)
    return 0


if __name__ == "__main__":
    sys.exit(main())
