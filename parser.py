#!/usr/bin/env python3
"""
Парсер каталога books.toscrape.com — учебного сайта для тренировки
парсеров (скрапинг там официально разрешён без ограничений).

По структуре и логике скрипт максимально близок к типовому заказу
"собрать каталог интернет-магазина в Excel": обход категорий с
пагинацией, сбор карточек товаров, сохранение в .xlsx.

Пример запуска:
    python parser.py
    python parser.py --category "Travel" --delay 1.0
    python parser.py --output output/mystery.xlsx --category Mystery
"""

import argparse
import os
import sys
import time
from urllib.parse import urljoin

import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.utils import get_column_letter

BASE_URL = "http://books.toscrape.com/"
USER_AGENT = "Mozilla/5.0 (compatible; PortfolioBookParser/1.0; +https://kwork.ru)"
REQUEST_TIMEOUT = 10        # секунд на один запрос
MAX_RETRIES = 3              # попыток на одну страницу при сетевой ошибке
RETRY_PAUSE = 2              # пауза перед повторной попыткой, секунд

RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
NOT_SPECIFIED = "не указано"


# =====================================================================
# Получение страницы
# =====================================================================

def get_page(url, delay):
    """Скачивает страницу с вежливой задержкой и ретраями при сетевой
    ошибке. Возвращает BeautifulSoup или None, если все попытки неудачны."""
    headers = {"User-Agent": USER_AGENT}

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            # сервер не указывает charset в заголовке Content-Type, из-за
            # этого requests по умолчанию берёт ISO-8859-1 и ломает кириллицу
            # и спецсимволы (например, £) — определяем кодировку по содержимому
            response.encoding = response.apparent_encoding
            time.sleep(delay)  # вежливая задержка перед следующим запросом
            return BeautifulSoup(response.text, "html.parser")
        except requests.RequestException as e:
            print(f"    Попытка {attempt}/{MAX_RETRIES} не удалась ({url}): {e}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_PAUSE)

    return None


# =====================================================================
# Разбор страниц
# =====================================================================

def get_categories(homepage_soup):
    """Возвращает список категорий каталога как [(название, url), ...]."""
    links = homepage_soup.select("div.side_categories ul.nav-list > li > ul > li > a")
    categories = []
    for a in links:
        name = a.get_text(strip=True)
        url = urljoin(BASE_URL, a.get("href", ""))
        categories.append((name, url))
    return categories


def get_next_page_url(soup, current_url):
    """Возвращает адрес следующей страницы пагинации или None, если
    текущая страница последняя."""
    next_tag = soup.select_one("li.next a")
    if next_tag and next_tag.get("href"):
        return urljoin(current_url, next_tag["href"])
    return None


def parse_product_card(card, page_url, category_name):
    """Извлекает данные о товаре из карточки на странице списка.
    Если поле не найдено на странице — подставляет "не указано" и
    не прерывает работу."""
    title_tag = card.select_one("h3 a")
    title = title_tag.get("title", "").strip() if title_tag and title_tag.get("title") else NOT_SPECIFIED
    link = urljoin(page_url, title_tag["href"]) if title_tag and title_tag.get("href") else NOT_SPECIFIED

    price_tag = card.select_one("p.price_color")
    price = price_tag.get_text(strip=True) if price_tag else NOT_SPECIFIED

    availability_tag = card.select_one("p.availability")
    if availability_tag:
        text = availability_tag.get_text(strip=True).lower()
        availability = "В наличии" if "in stock" in text else "Нет в наличии"
    else:
        availability = NOT_SPECIFIED

    rating_tag = card.select_one("p.star-rating")
    rating = NOT_SPECIFIED
    if rating_tag:
        for css_class in rating_tag.get("class", []):
            if css_class in RATING_WORDS:
                rating = RATING_WORDS[css_class]
                break

    return {
        "Название": title,
        "Цена": price,
        "Наличие": availability,
        "Категория": category_name,
        "Рейтинг": rating,
        "Ссылка": link,
    }


def scrape_category(name, start_url, delay):
    """Обходит все страницы категории (с пагинацией) и собирает товары.
    Возвращает (список товаров, количество ошибок/пропусков)."""
    products = []
    errors = 0
    url = start_url
    page_num = 1

    while url:
        print(f"  [{name}] страница {page_num}: {url}")
        soup = get_page(url, delay)

        if soup is None:
            print(f"  [{name}] страница {page_num} не загружена после {MAX_RETRIES} попыток — пропущена")
            errors += 1
            break

        cards = soup.select("article.product_pod")
        for card in cards:
            try:
                products.append(parse_product_card(card, url, name))
            except Exception as e:
                print(f"  [{name}] пропущен товар из-за ошибки разбора: {e}")
                errors += 1

        url = get_next_page_url(soup, url)
        page_num += 1

    return products, errors


# =====================================================================
# Сохранение результата
# =====================================================================

def save_to_excel(products, output_path):
    if not products:
        sys.exit("Ошибка: не собрано ни одного товара — файл не создан.")

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    columns = ["Название", "Цена", "Наличие", "Категория", "Рейтинг", "Ссылка"]
    df = pd.DataFrame(products, columns=columns)

    try:
        df.to_excel(output_path, index=False)
    except PermissionError:
        sys.exit(
            f"Ошибка: нет доступа для записи файла {output_path} "
            "(возможно, он открыт в другой программе)."
        )

    set_column_widths(output_path, df)


def set_column_widths(path, df):
    """Расширяет столбцы по самому длинному значению, чтобы файл сразу
    открывался читаемым, без ручной подгонки ширины в Excel."""
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    for i, col in enumerate(df.columns, start=1):
        longest = max([len(col)] + [len(str(v)) for v in df[col]])
        sheet.column_dimensions[get_column_letter(i)].width = min(longest + 2, 60)
    workbook.save(path)


# =====================================================================
# Основной сценарий
# =====================================================================

def main():
    parser = argparse.ArgumentParser(
        prog="parser.py",
        description="Парсер каталога books.toscrape.com в Excel.",
    )
    parser.add_argument(
        "--category",
        help="Название категории для парсинга (например, 'Travel'). "
             "По умолчанию — все категории каталога.",
    )
    parser.add_argument(
        "--output",
        default="output/books.xlsx",
        help="Путь к результату (.xlsx). По умолчанию: output/books.xlsx",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.75,
        help="Задержка между запросами в секундах (по умолчанию 0.75)",
    )
    args = parser.parse_args()

    print(f"Подключаюсь к {BASE_URL} ...")
    homepage_soup = get_page(BASE_URL, args.delay)
    if homepage_soup is None:
        sys.exit(f"Ошибка: не удалось подключиться к сайту {BASE_URL}")

    all_categories = get_categories(homepage_soup)
    if not all_categories:
        sys.exit("Ошибка: не удалось найти список категорий на сайте")

    if args.category:
        categories = [c for c in all_categories if c[0].lower() == args.category.strip().lower()]
        if not categories:
            available = ", ".join(name for name, _ in all_categories)
            sys.exit(f"Ошибка: категория '{args.category}' не найдена. Доступные категории: {available}")
    else:
        categories = all_categories

    print(f"К обработке: {len(categories)} категори(я/й)\n")

    all_products = []
    total_errors = 0

    for name, url in categories:
        products, errors = scrape_category(name, url, args.delay)
        all_products.extend(products)
        total_errors += errors
        print(f"  [{name}] собрано товаров: {len(products)}\n")

    save_to_excel(all_products, args.output)

    print("=== Итоговая статистика ===")
    print(f"Категорий обработано: {len(categories)}")
    print(f"Товаров собрано: {len(all_products)}")
    print(f"Ошибок/пропусков: {total_errors}")
    print(f"Результат сохранён в: {args.output}")


if __name__ == "__main__":
    main()
