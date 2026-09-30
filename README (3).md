# 🐍 Python Capstone Projects

Two hands-on capstone projects built as part of the **FITA Academy Python Programming (Level 1)** syllabus. They cover web scraping and office automation, two of the most common real-world uses of Python.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Cross--platform-lightgrey)

---

## 📁 Projects

| # | Project | What it does | Tech used |
|---|---------|--------------|-----------|
| 1 | [**Web Scraper**](01-web-scraper) | Extracts book data from multiple web pages and saves it to CSV | `requests`, `BeautifulSoup`, `csv` |
| 2 | [**Excel Automation**](02-excel-automation) | Automates Microsoft Excel: edits data, adds formulas, formats and saves | `pywin32` (`win32com`) |

---

## 1️⃣ Web Scraping using a Python Script

Scrapes book **title, price, rating, availability and URL** from [books.toscrape.com](https://books.toscrape.com) (a practice site built for scraping) and exports everything to a structured CSV file.

**How it works**
1. Identify the target website
2. Analyse the page structure
3. Extract the data with `requests` and `BeautifulSoup`
4. Store it in a CSV and print a quick summary (average price, number of 5-star books)

**Run it**
```bash
cd 01-web-scraper
pip install -r requirements.txt
python web_scraper.py --pages 5
```

---

## 2️⃣ Excel Automation using Python

Controls the real Excel application through `win32com` to do work that would normally be done by hand.

**What it automates**
- Opens an existing workbook and navigates to a worksheet
- Updates values, inserts a new column and row, and writes formulas
- Formats cells (bold headers, colours, currency format, auto-fit columns)
- Saves the file with `SaveAs` and closes Excel with `Quit`

**Run it** (Windows + Microsoft Excel required)
```bash
cd 02-excel-automation
pip install -r requirements.txt
python excel_automation.py
```

---

## 🗂️ Repository Structure

```
fita-python-capstone/
├── README.md
├── .gitignore
├── 01-web-scraper/
│   ├── web_scraper.py
│   ├── requirements.txt
│   └── README.md
└── 02-excel-automation/
    ├── excel_automation.py
    ├── requirements.txt
    └── README.md
```

---

## 🧠 Skills Demonstrated

- Python fundamentals: functions, loops, conditionals, file I/O
- Exception handling and clean script structure
- HTTP requests and HTML parsing
- Working with CSV files
- COM automation of desktop applications
- Git and GitHub for version control

---

## ⚙️ Getting Started

```bash
git clone https://github.com/harshitaaprabu/fita-python-capstone.git
cd fita-python-capstone
```

Then follow the instructions in each project folder.

---

## ⚠️ Notes

- Only scrape websites you have permission to scrape. Check a site's terms of service and `robots.txt` first.
- The Excel project works only on **Windows** with desktop Excel installed.

---

## 👩‍💻 Author

**Harshita Prabu**
GitHub: [@harshitaaprabu](https://github.com/harshitaaprabu)
