import time
from collections import Counter

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

base = "https://quotes.toscrape.com"
quote_rows = []
authors = {}
page = 1

while True:
    response = requests.get(f"{base}/page/{page}/")
    soup = BeautifulSoup(response.text, "html.parser")
    quotes = soup.find_all("div", class_="quote")

    if len(quotes) == 0:
        break

    for q in quotes:
        text = q.find("span", class_="text").get_text(strip=True).strip("\u201c\u201d")
        name = q.find("small", class_="author").get_text(strip=True)
        link = q.find("a", href=True)["href"]
        tags = [t.get_text(strip=True) for t in q.find_all("a", class_="tag")]

        quote_rows.append({"Quote": text, "Author": name, "Tags": ", ".join(tags)})

        if link not in authors:
            a_resp = requests.get(base + link)
            a_soup = BeautifulSoup(a_resp.text, "html.parser")
            born_date = a_soup.find("span", class_="author-born-date").get_text(strip=True)
            born_place = a_soup.find("span", class_="author-born-location").get_text(strip=True)
            authors[link] = {
                "Author": name,
                "Born Date": born_date,
                "Born Place": born_place.removeprefix("in "),
            }
            time.sleep(0.5)

    print("Page", page, "done -", len(quotes), "quotes")
    page += 1
    time.sleep(1)

quotes_df = pd.DataFrame(quote_rows)
authors_df = pd.DataFrame(authors.values())
authors_df["Quotes"] = authors_df["Author"].map(quotes_df["Author"].value_counts())

counts = Counter(
    tag for row in quote_rows for tag in row["Tags"].split(", ") if tag
)
tags_df = pd.DataFrame(counts.most_common(), columns=["Tag", "Quotes"])

file = "quotes_report.xlsx"
with pd.ExcelWriter(file) as writer:
    quotes_df.to_excel(writer, sheet_name="Quotes", index=False)
    authors_df.to_excel(writer, sheet_name="Authors", index=False)
    tags_df.to_excel(writer, sheet_name="Tag Summary", index=False)

wb = load_workbook(file)
for ws in wb.worksheets:
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.freeze_panes = "A2"
    for col in ws.columns:
        longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(longest + 3, 80)
wb.save(file)

print("Saved", len(quotes_df), "quotes and", len(authors_df), "authors to", file)