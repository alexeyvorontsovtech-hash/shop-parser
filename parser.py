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
USER_AGENT = "Mozilla/5.0 (compatible; PortfolioBookParser/1.0; https://github.com/alexeyvorontsovtech-hash/shop-parser)"
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
    ошибке, таймауте, 5xx или 429. Ошибки 4xx (кроме 429) не повторяются —
    это означает проблему с самим запросом, а не временный сбой сервера.
    Возвращает BeautifulSoup или None, если страница так и не загружена."""
    headers = {"User-Agent": USER_AGENT}

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            time.sleep(delay)  # вежливая задержка перед следующим запросом
            return BeautifulSoup(response.content, "html.parser")
        except requests.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            if status is not None and status != 429 and status < 500:
                print(f"    Ошибка {status} ({url}): повторные попытки не выполняются")
                return None
            print(f"    Попытка {attempt}/{MAX_RETRIES} не удалась ({url}): {e}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_PAUSE)
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
    Возвращает (список товаров, ошибок разбора карточек, признак полноты
    категории, номер последней успешно собранной страницы)."""
    products = []
    errors = 0
    complete = True
    last_completed_page = 0
    url = start_url
    page_num = 1

    while url:
        print(f"  [{name}] страница {page_num}: {url}")
        soup = get_page(url, delay)

        if soup is None:
            print(f"  [{name}] страница {page_num} не загружена после {MAX_RETRIES} попыток — категория неполная")
            complete = False
            break

        cards = soup.select("article.product_pod")
        if not cards:
            print(f"  [{name}] страница {page_num} не содержит товаров — категория неполная")
            complete = False
            break

        for card in cards:
            try:
                products.append(parse_product_card(card, url, name))
            except Exception as e:
                print(f"  [{name}] пропущен товар из-за ошибки разбора: {e}")
                errors += 1
                complete = False

        last_completed_page = page_num
        url = get_next_page_url(soup, url)
        page_num += 1

    return products, errors, complete, last_completed_page


# =====================================================================
# Сохранение результата
# =====================================================================

def parse_price(value):
    """Превращает текст цены вида '£45.17' в число. Если распознать не
    удалось (например, значение — NOT_SPECIFIED), возвращает None."""
    if isinstance(value, str):
        cleaned = value.replace("£", "").replace(",", "").strip()
        try:
            return float(cleaned)
        except ValueError:
            return None
    return value


def save_to_excel(products, output_path):
    if not products:
        sys.exit("Ошибка: не собрано ни одного товара — файл не создан.")

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    columns = ["Название", "Цена", "Наличие", "Категория", "Рейтинг", "Ссылка"]
    df = pd.DataFrame(products, columns=columns)
    df["Цена"] = df["Цена"].apply(parse_price)
    df = df.rename(columns={"Цена": "Цена, £"})

    try:
        df.to_excel(output_path, index=False)
    except PermissionError:
        sys.exit(
            f"Ошибка: нет доступа для записи файла {output_path} "
            "(возможно, он открыт в другой программе)."
        )

    apply_excel_formatting(output_path, df)


def apply_excel_formatting(path, df):
    """Расширяет столбцы по самому длинному значению, задаёт числовой
    формат цене, делает ссылки кликабельными, включает автофильтр и
    закрепляет строку заголовка."""
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    for i, col in enumerate(df.columns, start=1):
        longest = max([len(str(col))] + [len(str(v)) for v in df[col]])
        sheet.column_dimensions[get_column_letter(i)].width = min(longest + 2, 60)

    price_col = df.columns.get_loc("Цена, £") + 1
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=price_col)
        if isinstance(cell.value, (int, float)):
            cell.number_format = "0.00"

    link_col = df.columns.get_loc("Ссылка") + 1
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=link_col)
        if cell.value and cell.value != NOT_SPECIFIED:
            cell.hyperlink = cell.value
            cell.style = "Hyperlink"

    sheet.auto_filter.ref = sheet.dimensions
    sheet.freeze_panes = "A2"

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

    print(f"Категорий к обработке: {len(categories)}\n")

    all_products = []
    total_errors = 0
    incomplete_categories = []

    for name, url in categories:
        products, errors, complete, last_page = scrape_category(name, url, args.delay)
        all_products.extend(products)
        total_errors += errors
        if not complete:
            incomplete_categories.append((name, last_page))
        print(f"  [{name}] собрано товаров: {len(products)}\n")

    save_to_excel(all_products, args.output)

    print("=== Итоговая статистика ===")
    print(f"Категорий обработано: {len(categories)}")
    print(f"Товаров собрано: {len(all_products)}")
    print(f"Ошибок/пропусков: {total_errors}")
    print(f"Результат сохранён в: {args.output}")

    if incomplete_categories:
        print("Неполные категории:")
        for name, last_page in incomplete_categories:
            print(f"  - {name}: собрано страниц — {last_page}")
        sys.exit(1)


if __name__ == "__main__":
    main()
