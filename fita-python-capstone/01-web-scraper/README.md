# Capstone 1: Web Scraping with Python

Scrapes book data (title, price, rating, availability, URL) from
[books.toscrape.com](https://books.toscrape.com), a site built for scraping practice,
and stores it in a CSV file.

## Approach
1. Identify the target website
2. Analyse the page structure (each book is an `article.product_pod`)
3. Extract data with `requests` + `BeautifulSoup`
4. Save to a structured CSV and print a short summary

## Run
```bash
pip install -r requirements.txt
python web_scraper.py                 # 3 pages by default
python web_scraper.py --pages 10 --out books.csv
```

## Output
`books.csv` with columns: `title, price_gbp, rating, in_stock, url`

## Notes
- A 1-second delay between requests keeps the scraper polite.
- Always check a site's terms of service and `robots.txt` before scraping real websites.
