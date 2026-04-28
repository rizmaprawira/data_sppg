import time
import re
import urllib.parse
from pathlib import Path

import requests
import pandas as pd
from bs4 import BeautifulSoup

BASE_URL = "https://www.bgn.go.id/operasional-sppg"

HEADERS = {
    "User-Agent": "Mozilla/5.0 compatible; data-export-script/1.0"
}

COLUMNS = [
    "No",
    "Provinsi SPPG",
    "Kab./Kota SPPG",
    "Kecamatan SPPG",
    "Kelurahan/Desa SPPG",
    "Alamat SPPG",
    "Nama SPPG",
]


def get_page(page: int) -> str:
    params = {"page": page, "search": ""}
    response = requests.get(BASE_URL, params=params, headers=HEADERS, timeout=30)
    response.raise_for_status()
    return response.text


def find_last_page(html: str) -> int:
    soup = BeautifulSoup(html, "html.parser")
    last_page = 1

    for a in soup.find_all("a", href=True):
        href = a["href"]
        parsed = urllib.parse.urlparse(href)
        query = urllib.parse.parse_qs(parsed.query)
        if "page" in query:
            try:
                last_page = max(last_page, int(query["page"][0]))
            except ValueError:
                pass

    return last_page


def parse_rows(html: str):
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")

    if not table:
        return []

    rows = []
    for tr in table.find_all("tr"):
        cells = [cell.get_text(" ", strip=True) for cell in tr.find_all(["td", "th"])]

        if not cells:
            continue

        if cells[0].lower() == "no":
            continue

        if len(cells) == len(COLUMNS):
            rows.append(cells)

    return rows


def main():
    first_html = get_page(1)
    last_page = find_last_page(first_html)

    print(f"Detected {last_page} pages")

    all_rows = []

    for page in range(1, last_page + 1):
        html = first_html if page == 1 else get_page(page)
        rows = parse_rows(html)
        all_rows.extend(rows)

        print(f"Page {page}/{last_page}: {len(rows)} rows, total {len(all_rows)}")

        time.sleep(0.2)  # polite delay

    df = pd.DataFrame(all_rows, columns=COLUMNS)

    df["No"] = pd.to_numeric(df["No"], errors="coerce").astype("Int64")

    csv_path = Path("bgn_sppg_operasional.csv")
    xlsx_path = Path("bgn_sppg_operasional.xlsx")

    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    df.to_excel(xlsx_path, index=False)

    print(f"Saved {len(df)} rows")
    print(f"CSV:  {csv_path.resolve()}")
    print(f"XLSX: {xlsx_path.resolve()}")


if __name__ == "__main__":
    main()
