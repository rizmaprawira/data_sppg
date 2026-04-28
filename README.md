# SPPG Data Scraper

Scrapers for operational SPPG data from the BGN website.

## Author

Rizma Prawira  
rizmaprawira17@gmail.com

## Contents

- `download_sppg.py`: Scrapes all available pages from the BGN operational SPPG listing and exports the result to CSV and XLSX.
- `sppg50.py`: Scrapes the first 5 pages and exports the result to CSV.

## Requirements

- Python 3.10+
- `requests`
- `pandas`
- `beautifulsoup4`
- `openpyxl` for writing Excel files

## Usage

Install dependencies:

```bash
pip install requests pandas beautifulsoup4 openpyxl
```

Run the full scraper:

```bash
python download_sppg.py
```

Run the first-5-pages scraper:

```bash
python sppg50.py
```

## Output Files

- `bgn_sppg_operasional.csv`
- `bgn_sppg_operasional.xlsx`
- `bgn_sppg_first_5_pages.csv`

## Notes

- The scripts use a short delay between requests to be polite to the source website.
- Output files are written to the project root.
