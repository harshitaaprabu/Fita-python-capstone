"""
Capstone Project 1 - Web scraping using a Python script

Target site : https://books.toscrape.com  (a practice site built for scraping)
Steps       : 1. identify the target website
              2. analyse page structure (each book is an <article class="product_pod">)
              3. extract title, price, rating, availability with requests + BeautifulSoup
              4. store the result in a structured CSV file

Install : pip install requests beautifulsoup4
Run     : python web_scraper.py            (scrapes 3 pages)
          python web_scraper.py --pages 10
"""
import argparse
import csv
import time
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://books.toscrape.com/catalogue/page-{}.html"
HEADERS = {"User-Agent": "Mozilla/5.0 (FITA capstone scraper; educational use)"}
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


def fetch(url):
    """Download a page; raise a clear error if something goes wrong."""
    response = requests.get(url, headers=HEADERS, timeout=15)
    response.raise_for_status()
    response.encoding = "utf-8"
    return response.text


def parse_books(html, page_url):
    """Extract one dict per book from a listing page."""
    soup = BeautifulSoup(html, "html.parser")
    books = []
    for card in soup.select("article.product_pod"):
        link = card.h3.a
        rating_word = card.select_one("p.star-rating")["class"][1]
        books.append({
            "title": link["title"],
            "price_gbp": float(card.select_one("p.price_color").text.strip().lstrip("£Â")),
            "rating": RATINGS.get(rating_word, 0),
            "in_stock": "In stock" in card.select_one("p.availability").text,
            "url": urljoin(page_url, link["href"]),
        })
    return books


def scrape(pages):
    all_books = []
    for page in range(1, pages + 1):
        url = BASE_URL.format(page)
        try:
            books = parse_books(fetch(url), url)
        except requests.RequestException as err:
            print(f"Page {page}: request failed ({err}); stopping.")
            break
        print(f"Page {page}: {len(books)} books")
        all_books.extend(books)
        time.sleep(1)  # be polite to the server
    return all_books


def save_csv(rows, path):
    if not rows:
        print("Nothing to save.")
        return
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    print(f"Saved {len(rows)} rows to {path}")


def summarise(rows):
    """Small analysis step: the kind of insight scraped data is used for."""
    if not rows:
        return
    avg = sum(r["price_gbp"] for r in rows) / len(rows)
    top = [r for r in rows if r["rating"] == 5]
    print(f"\nAverage price: £{avg:.2f} | 5-star books: {len(top)} of {len(rows)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--pages", type=int, default=3)
    parser.add_argument("--out", default="books.csv")
    args = parser.parse_args()

    data = scrape(args.pages)
    save_csv(data, args.out)
    summarise(data)
