"""
Сбор данных с books.toscrape.com — учебного сайта, созданного специально для практики парсинга.

Логика:
  1. С главной страницы берём список категорий (50 штук).
  2. Для каждой нужной категории проходим страницы каталога (пагинация через кнопку «next»).
  3. С карточки в каталоге берём название, цену, наличие, рейтинг, ссылку и обложку.
  4. С флагом details дополнительно открываем страницу каждой книги:
     там есть точное количество на складе, артикул (UPC) и описание.
"""

import logging
import re
import time
from dataclasses import dataclass, field
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/"

# Рейтинг на сайте записан словом в классе: <p class="star-rating Three">
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

log = logging.getLogger("parser")


class SiteUnavailableError(Exception):
    """Сайт не отвечает — дальше работать бессмысленно."""


@dataclass
class Book:
    """Одна книга — одна строка будущей таблицы."""
    title: str
    category: str
    price_gbp: float
    in_stock: bool
    rating: int
    url: str
    image_url: str
    # Поля ниже заполняются только в режиме details
    quantity: int | None = None
    upc: str | None = None
    description: str | None = None


@dataclass
class ScrapeStats:
    """Статистика запуска — попадает в лог и на лист «Параметры» в Excel."""
    pages: int = 0
    requests: int = 0
    errors: list[str] = field(default_factory=list)


class BooksScraper:
    def __init__(self, delay: float = 0.5, timeout: float = 20):
        self.delay = delay          # пауза между запросами, чтобы не нагружать сайт
        self.timeout = timeout      # сколько секунд ждать ответа
        self.stats = ScrapeStats()
        self.session = self._make_session()

    @staticmethod
    def _make_session() -> requests.Session:
        """Сессия с автоповтором: при обрыве или ошибке 5xx запрос повторится до 3 раз с растущей паузой."""
        session = requests.Session()
        retry = Retry(
            total=3,
            backoff_factor=1,                       # паузы 1, 2, 4 секунды
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=("GET",),
        )
        session.mount("https://", HTTPAdapter(max_retries=retry))
        session.mount("http://", HTTPAdapter(max_retries=retry))
        session.headers["User-Agent"] = "Mozilla/5.0 (books-parser demo; +https://github.com/sigutailya-sketch)"
        return session

    # ------------------------------------------------------------------ загрузка страниц

    def _get(self, url: str) -> BeautifulSoup:
        """Скачивает страницу и возвращает разобранный HTML. Между запросами — пауза."""
        if self.stats.requests:
            time.sleep(self.delay)
        self.stats.requests += 1
        response = self.session.get(url, timeout=self.timeout)
        response.raise_for_status()
        response.encoding = "utf-8"  # иначе символ фунта превращается в «Â£»
        return BeautifulSoup(response.text, "lxml")

    # ------------------------------------------------------------------ категории

    def get_categories(self) -> dict[str, str]:
        """Возвращает {название категории: ссылка на первую страницу}."""
        try:
            soup = self._get(BASE_URL)
        except requests.RequestException as e:
            raise SiteUnavailableError(f"Не удалось открыть {BASE_URL}: {e}") from e

        links = soup.select("div.side_categories ul li ul li a")
        return {a.get_text(strip=True): urljoin(BASE_URL, a["href"]) for a in links}

    # ------------------------------------------------------------------ каталог

    def scrape_category(self, name: str, url: str, max_pages: int | None = None):
        """Проходит страницы категории и по одной отдаёт найденные книги (генератор)."""
        page_num = 0
        while url:
            page_num += 1
            try:
                soup = self._get(url)
            except requests.RequestException as e:
                # Одна битая страница не должна ронять весь сбор — записываем и идём дальше
                msg = f"{name}, страница {page_num}: {e}"
                log.error("Ошибка загрузки — %s", msg)
                self.stats.errors.append(msg)
                return

            self.stats.pages += 1
            cards = soup.select("article.product_pod")
            total_pages = self._total_pages(soup)
            log.info("%s: страница %d/%d, книг на странице: %d", name, page_num, total_pages, len(cards))

            for card in cards:
                try:
                    yield self._parse_card(card, name, url)
                except (AttributeError, KeyError, ValueError) as e:
                    # Разметка карточки не такая, как ожидали — пропускаем её, но фиксируем
                    msg = f"{name}, страница {page_num}: не разобрана карточка ({e})"
                    log.warning(msg)
                    self.stats.errors.append(msg)

            if max_pages and page_num >= max_pages:
                break
            next_link = soup.select_one("li.next a")
            url = urljoin(url, next_link["href"]) if next_link else None

    @staticmethod
    def _total_pages(soup: BeautifulSoup) -> int:
        """Из «Page 1 of 50» достаёт 50. Если категория в одну страницу — счётчика нет."""
        current = soup.select_one("li.current")
        match = re.search(r"of\s+(\d+)", current.get_text()) if current else None
        return int(match.group(1)) if match else 1

    @staticmethod
    def _parse_card(card, category: str, page_url: str) -> Book:
        link = card.select_one("h3 a")
        rating_classes = card.select_one("p.star-rating")["class"]
        rating = next((RATING_WORDS[c] for c in rating_classes if c in RATING_WORDS), 0)
        price_text = card.select_one("p.price_color").get_text()

        return Book(
            title=link["title"],  # в тексте ссылки название обрезано «...», полное — в title
            category=category,
            price_gbp=parse_price(price_text),
            in_stock="In stock" in card.select_one("p.availability").get_text(),
            rating=rating,
            url=urljoin(page_url, link["href"]),
            image_url=urljoin(page_url, card.select_one("img")["src"]),
        )

    # ------------------------------------------------------------------ страница книги

    def fill_details(self, book: Book) -> None:
        """Открывает страницу книги и дописывает количество, UPC и описание."""
        try:
            soup = self._get(book.url)
        except requests.RequestException as e:
            msg = f"«{book.title}»: не открылась страница книги ({e})"
            log.error(msg)
            self.stats.errors.append(msg)
            return

        table = {tr.th.get_text(strip=True): tr.td.get_text(strip=True) for tr in soup.select("table.table-striped tr")}
        book.upc = table.get("UPC")
        qty = re.search(r"\((\d+) available\)", table.get("Availability", ""))
        book.quantity = int(qty.group(1)) if qty else 0

        # Описание лежит в первом <p> после блока с id="product_description"
        desc = soup.select_one("#product_description + p")
        book.description = desc.get_text(strip=True) if desc else ""


def parse_price(text: str) -> float:
    """'£51.77' → 51.77"""
    return float(re.sub(r"[^\d.]", "", text))


def get_gbp_rate() -> float | None:
    """Курс фунта по данным ЦБ РФ (через открытое зеркало cbr-xml-daily.ru). Нет связи — вернёт None."""
    try:
        data = requests.get("https://www.cbr-xml-daily.ru/daily_json.js", timeout=10).json()
        gbp = data["Valute"]["GBP"]
        return gbp["Value"] / gbp["Nominal"]
    except (requests.RequestException, KeyError, ValueError) as e:
        log.warning("Курс фунта не получен (%s) — колонка в рублях останется пустой", e)
        return None
