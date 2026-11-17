\# Quotes and Authors Scraper



Python scraper that collects data from a list site and follows links to detail pages, then delivers a clean, formatted Excel report.



\## What it does



\- \*\*List pages:\*\* scrapes all 10 pages of quotes (100 quotes) with their authors and tags.

\- \*\*Detail pages:\*\* follows each quote's author link and collects birth date and birth place for all 50 authors, visiting each author only once.

\- \*\*Cleaning:\*\* removes the quote marks around each quote and the leading "in" from birth places.

\- \*\*Report:\*\* one Excel file with 3 sheets: Quotes, Authors (with a quote count per author) and Tag Summary (how often each tag appears).



\## Tools used



Python, requests, BeautifulSoup, pandas, openpyxl



\## How to run



&#x20;   pip install -r requirements.txt

&#x20;   python quotes.py



Output: `quotes\_report.xlsx`



\## Data source



\[quotes.toscrape.com](https://quotes.toscrape.com), a site built for practising web scraping.



\## Available for freelance work



I build scrapers that collect data from websites and deliver it as clean Excel or CSV files.

