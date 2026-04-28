import time
from pathlib import Path

import requests
import pandas as pd
from bs4 import BeautifulSoup

BASE_URL = "https://www.bgn.go.id/operasional-sppg"
MAX_PAGES = 5

HEADERS = {
    "User-Agent": "Mozilla/5.0 compatible; bgn-sppg-test-export/1.0"
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
    params = {
        "page": page,
        "search": "",
    }

    response = requests.get(
        BASE_URL,
        params=params,
        headers=HEADERS,
        timeout=30,
    )
    response.raise_for_status()
    return response.text


def parse_rows(html: str):
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")

    if table is None:
        return []

    rows = []

    for tr in table.find_all("tr"):
        cells = [
            cell.get_text(" ", strip=True)
            for cell in tr.find_all(["td", "th"])
        ]

        if not cells:
            continue

        # Skip header row
        if cells[0].strip().lower() == "no":
            continue

        if len(cells) == len(COLUMNS):
            rows.append(cells)

    return rows


def main():
    all_rows = []

    for page in range(1, MAX_PAGES + 1):
        html = get_page(page)
        rows = parse_rows(html)
        all_rows.extend(rows)

        print(f"Page {page}/{MAX_PAGES}: {len(rows)} rows, total {len(all_rows)}")

        time.sleep(0.2)

    df = pd.DataFrame(all_rows, columns=COLUMNS)

    df["No"] = pd.to_numeric(df["No"], errors="coerce").astype("Int64")

    output_path = Path("bgn_sppg_first_5_pages.csv")
    df.to_csv(output_path, index=False, encoding="utf-8-sig")

    print(f"Saved {len(df)} rows to {output_path.resolve()}")


if __name__ == "__main__":
    main()
