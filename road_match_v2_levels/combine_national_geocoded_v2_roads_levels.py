#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path

import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_FILES = [
    OUTPUT_DIR / "bgn_sppg_operasional_geocoded_sumatra_v2_roads_levels.xlsx",
    OUTPUT_DIR / "bgn_sppg_operasional_geocoded_java_v2_roads_levels.xlsx",
    OUTPUT_DIR / "bgn_sppg_operasional_geocoded_kalimantan_v2_roads_levels.xlsx",
    OUTPUT_DIR / "bgn_sppg_operasional_geocoded_sulawesi_v2_roads_levels.xlsx",
    OUTPUT_DIR / "bgn_sppg_operasional_geocoded_nusa-tenggara_v2_roads_levels.xlsx",
    OUTPUT_DIR / "bgn_sppg_operasional_geocoded_maluku_v2_roads_levels.xlsx",
    OUTPUT_DIR / "bgn_sppg_operasional_geocoded_papua_v2_roads_levels.xlsx",
]
OUTPUT_PATH = OUTPUT_DIR / "bgn_sppg_operasional_geocoded_national_v2_roads_levels.xlsx"
SOURCE_SHEET = "Sheet1"


def load_region_sheet(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Missing regional workbook: {path}")
    df = pd.read_excel(path, sheet_name=SOURCE_SHEET)
    df["source_workbook"] = path.name
    return df


def main() -> None:
    frames: list[pd.DataFrame] = []
    for path in OUTPUT_FILES:
        frames.append(load_region_sheet(path))

    combined = pd.concat(frames, ignore_index=True)
    columns = [column for column in combined.columns if column != "source_workbook"]
    columns.append("source_workbook")
    combined = combined.loc[:, columns]

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    combined.to_excel(OUTPUT_PATH, index=False)

    print(f"Combined {len(frames)} workbooks into {OUTPUT_PATH.resolve()}")
    print(f"Rows: {len(combined)}")


if __name__ == "__main__":
    main()
