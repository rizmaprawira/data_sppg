# SPPG Data Scraper

Scrapers for operational SPPG data from the BGN website.

## Author

Rizma Prawira  
rizmaprawira17@gmail.com

## Contents

- `download_sppg.py`: Scrapes all available pages from the BGN operational SPPG listing and exports the result to CSV and XLSX.
- `sppg50.py`: Scrapes the first 5 pages and exports the result to CSV.
- `geocode_sppg_osm.py`: Core geocoder that adds `latitude` and `longitude` to `bgn_sppg_operasional.xlsx` using the local OSM extracts.
- `geocode_sppg_osm_sumatra.py`, `geocode_sppg_osm_java.py`, `geocode_sppg_osm_kalimantan.py`, `geocode_sppg_osm_sulawesi.py`, `geocode_sppg_osm_nusa_tenggara.py`, `geocode_sppg_osm_maluku.py`, `geocode_sppg_osm_papua.py`: Island-specific entrypoints that can be run independently or in parallel.
- `geocode_sppg_osm_java_roads.py`: Standalone Java geocoder that reads the road shapefile at `/Users/rizzie/ClimateData/OSM_roads/gis_osm_roads_free_1.shp` in small bbox-scoped batches.
- `geocode_sppg_osm_sumatra_roads.py`: Standalone Sumatra geocoder that batches by `kab/kota` and reads the same road shapefile in city-sized bbox-scoped batches.
- `geocode_sppg_osm_sulawesi_roads.py`: Standalone Sulawesi geocoder that follows the same Sumatra batching/resume logic.
- `geocode_sppg_osm_kalimantan_roads.py`, `geocode_sppg_osm_nusa_tenggara_roads.py`, `geocode_sppg_osm_maluku_roads.py`, `geocode_sppg_osm_papua_roads.py`: Standalone road-based geocoders for the other island groups using the same city-sized batching pattern.
- `geocode_sppg_osm_java_v2_roads.py`: Standalone Java geocoder that also batches by `kab/kota` city bbox instead of kecamatan.
- `combine_national_geocoded.py`: Merges the regional road-based outputs into one national workbook.
- `road_match_fuzzy_v1/retry_national_fuzzy.py`: Builds a separate national workbook from the regional outputs and retries empty rows with a looser road match, writing to `road_match_fuzzy_v1/output/`.

## Requirements

- Python 3.10+
- `requests`
- `pandas`
- `beautifulsoup4`
- `openpyxl` for writing Excel files
- For geocoding: `geopandas`, `shapely`, `pyrosm`, `pyogrio`

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

Geocode the operational workbook from the local OSM extracts:

```bash
conda run -n osm python geocode_sppg_osm.py
```

To run one island at a time, use the island wrapper scripts, for example:

```bash
conda run -n osm python geocode_sppg_osm_sumatra.py
conda run -n osm python geocode_sppg_osm_java.py
python geocode_sppg_osm_sumatra_roads.py
python geocode_sppg_osm_sulawesi_roads.py
python geocode_sppg_osm_kalimantan_roads.py
python geocode_sppg_osm_nusa_tenggara_roads.py
python geocode_sppg_osm_maluku_roads.py
python geocode_sppg_osm_papua_roads.py
python geocode_sppg_osm_java_v2_roads.py
python combine_national_geocoded.py
python geocode_sppg_osm_java_roads.py
```

Both road-based scripts resume automatically if the output workbook already exists. Use `--restart` to ignore the partial workbook and start fresh.

Each island script writes its own workbook, such as `bgn_sppg_operasional_geocoded_sumatra.xlsx`, and adds an `unresolved_rows` review sheet for rows that could not be matched conservatively.

The Java script now uses the local kelurahan boundary shapefile to batch rows by kecamatan, loads a small OSM bounding box for each batch, and saves the workbook after every batch finishes so you can inspect partial output during long runs.

## Output Files

- `bgn_sppg_operasional.csv`
- `bgn_sppg_operasional.xlsx`
- `bgn_sppg_operasional_geocoded.xlsx`
- `bgn_sppg_operasional_geocoded_sumatra.xlsx`
- `bgn_sppg_operasional_geocoded_sumatra_roads.xlsx`
- `bgn_sppg_operasional_geocoded_sulawesi_roads.xlsx`
- `bgn_sppg_operasional_geocoded_kalimantan_roads.xlsx`
- `bgn_sppg_operasional_geocoded_nusa_tenggara_roads.xlsx`
- `bgn_sppg_operasional_geocoded_maluku_roads.xlsx`
- `bgn_sppg_operasional_geocoded_papua_roads.xlsx`
- `bgn_sppg_operasional_geocoded_java_v2_roads.xlsx`
- `bgn_sppg_operasional_geocoded_national.xlsx`
- `bgn_sppg_operasional_geocoded_java.xlsx`
- `bgn_sppg_operasional_geocoded_java_roads.xlsx`
- `bgn_sppg_operasional_geocoded_kalimantan.xlsx`
- `bgn_sppg_operasional_geocoded_sulawesi.xlsx`
- `bgn_sppg_operasional_geocoded_nusa_tenggara.xlsx`
- `bgn_sppg_operasional_geocoded_maluku.xlsx`
- `bgn_sppg_operasional_geocoded_papua.xlsx`
- `bgn_sppg_first_5_pages.csv`

## Notes

- The scripts use a short delay between requests to be polite to the source website.
- Output files are written to the project root.
