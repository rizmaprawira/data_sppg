#!/usr/bin/env python3
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from shapely.geometry import box

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import geocode_sppg_osm_sumatra_roads as base


INPUT_FILES = [
    Path("bgn_sppg_operasional_geocoded_sumatra_roads.xlsx"),
    Path("bgn_sppg_operasional_geocoded_java_v2_roads.xlsx"),
    Path("bgn_sppg_operasional_geocoded_kalimantan_roads.xlsx"),
    Path("bgn_sppg_operasional_geocoded_sulawesi_roads.xlsx"),
    Path("bgn_sppg_operasional_geocoded_nusa_tenggara_roads.xlsx"),
    Path("bgn_sppg_operasional_geocoded_maluku_roads.xlsx"),
    Path("bgn_sppg_operasional_geocoded_papua_roads.xlsx"),
]

SOURCE_SHEET = "Sheet1"
OUTPUT_DIR = Path("road_match_fuzzy_v1/output")
OUTPUT_PATH = OUTPUT_DIR / "bgn_sppg_operasional_geocoded_national_fuzzy.xlsx"
DEFAULT_BOUNDARY_PATH = Path(
    "/Users/rizzie/Work/IndonesiaRe/data/batas_keldesa/batas_kabkota/BATAS KABUPATEN KOTA DESEMBER 2019 DUKCAPIL.shp"
)
DEFAULT_ROAD_SHP_PATH = Path("/Users/rizzie/ClimateData/OSM_roads/gis_osm_roads_free_1.shp")

FUZZY_MIN_MATCH_SCORE = 0.74
FUZZY_MIN_MATCH_MARGIN = 0.03


def load_region_sheet(path: Path, source_rank: int) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Missing regional workbook: {path}")

    df = pd.read_excel(path, sheet_name=SOURCE_SHEET)
    df["source_workbook"] = path.name
    df["source_rank"] = source_rank
    df["source_row_index"] = range(len(df))
    df["row_status"] = df["latitude"].notna() & df["longitude"].notna()
    return df


def fuzzy_geocode_with_path(
    roads: base.PreparedRoadLayer,
    path: base.CityPathResult,
    address: object,
) -> dict[str, object]:
    phrases = base.candidate_phrases(address)
    if not phrases:
        return {
            "success": False,
            "latitude": None,
            "longitude": None,
            "road_name": None,
            "candidate_phrase": None,
            "score": None,
            "candidate_count": 0,
            "reason": "no_address_candidate",
        }

    boundary_geometry = path.geometry
    best: tuple[str, str, float, int, object] | None = None
    candidate_total = 0

    for phrase in phrases:
        scored_candidates = base.road_name_candidates(roads, phrase)
        candidate_total += len(scored_candidates)
        if not scored_candidates:
            continue

        top_score = scored_candidates[0][1]
        second_score = scored_candidates[1][1] if len(scored_candidates) > 1 else 0.0
        if top_score < FUZZY_MIN_MATCH_SCORE or (top_score - second_score) < FUZZY_MIN_MATCH_MARGIN:
            continue

        road_name, score, overlap = scored_candidates[0]
        clipped = base.choose_road_geometry(roads, road_name, boundary_geometry)
        if clipped is None:
            continue
        best = (phrase, road_name, score, overlap, clipped)
        break

    if best is None:
        return {
            "success": False,
            "latitude": None,
            "longitude": None,
            "road_name": None,
            "candidate_phrase": None,
            "score": None,
            "candidate_count": candidate_total,
            "reason": "road_candidate_not_matched",
        }

    phrase, road_name, score, overlap, clipped = best
    point = clipped.representative_point()
    return {
        "success": True,
        "latitude": float(point.y),
        "longitude": float(point.x),
        "road_name": road_name,
        "candidate_phrase": phrase,
        "score": score,
        "candidate_count": candidate_total,
        "reason": "matched",
    }


def retry_row(
    roads: base.PreparedRoadLayer,
    row: pd.Series | dict[str, object],
    path: base.CityPathResult | None,
) -> dict[str, object]:
    if path is None:
        return {
            "latitude": None,
            "longitude": None,
            "status": "unresolved",
            "reason": "kabkota_boundary_not_found",
            "admin_level_used": None,
            "candidate_phrase": None,
            "matched_road_name": None,
            "match_score": None,
            "candidate_count": 0,
        }

    result = fuzzy_geocode_with_path(roads, path, row["Alamat SPPG"])
    if result["success"]:
        return {
            "latitude": result["latitude"],
            "longitude": result["longitude"],
            "status": "matched",
            "reason": "matched",
            "admin_level_used": path.admin_level_used,
            "candidate_phrase": result["candidate_phrase"],
            "matched_road_name": result["road_name"],
            "match_score": result["score"],
            "candidate_count": result["candidate_count"],
        }

    return {
        "latitude": None,
        "longitude": None,
        "status": "unresolved",
        "reason": result["reason"],
        "admin_level_used": path.admin_level_used,
        "candidate_phrase": None,
        "matched_road_name": None,
        "match_score": None,
        "candidate_count": result["candidate_count"],
    }


def main() -> None:
    catalog = base.build_kabkota_catalog(DEFAULT_BOUNDARY_PATH)

    frames: list[pd.DataFrame] = []
    for rank, path in enumerate(INPUT_FILES):
        frame = load_region_sheet(path, rank)
        frames.append(frame)

    combined = pd.concat(frames, ignore_index=True)
    combined["retry_status"] = "already_matched"
    combined["retry_reason"] = "existing_output"
    combined["retry_candidate_phrase"] = None
    combined["retry_matched_road_name"] = None
    combined["retry_match_score"] = None

    unresolved_mask = combined["latitude"].isna() | combined["longitude"].isna()
    unresolved = combined.loc[unresolved_mask].copy()

    if not unresolved.empty:
        unresolved["job_key"] = [
            (base.normalize_province_name(row["Provinsi SPPG"]), base.normalize_kabkota_name(row["Kab./Kota SPPG"]))
            for row in unresolved.to_dict(orient="records")
        ]

        job_order = list(dict.fromkeys(unresolved["job_key"].tolist()))
        for job_key in job_order:
            job_rows = unresolved.loc[unresolved["job_key"] == job_key]
            province_key, kabkota_key = job_key
            bounds = base.resolve_kabkota_bounds(catalog, kabkota_key)
            if bounds is None:
                combined.loc[job_rows.index, "retry_status"] = "unresolved"
                combined.loc[job_rows.index, "retry_reason"] = "kabkota_boundary_not_found"
                continue

            job_geometry = box(*bounds)
            bbox = tuple(job_geometry.bounds)
            roads = base.load_road_layer_for_bbox(DEFAULT_ROAD_SHP_PATH, bbox=bbox)

            for idx, row in job_rows.iterrows():
                city_path = base.CityPathResult(
                    province_name=str(row["Provinsi SPPG"]),
                    kabkota_name=str(row["Kab./Kota SPPG"]),
                    geometry=job_geometry,
                )
                result = retry_row(roads, row, city_path)
                if result["status"] == "matched":
                    combined.at[idx, "latitude"] = result["latitude"]
                    combined.at[idx, "longitude"] = result["longitude"]
                    combined.at[idx, "retry_status"] = "fuzzy_matched"
                    combined.at[idx, "retry_reason"] = "matched"
                    combined.at[idx, "retry_candidate_phrase"] = result["candidate_phrase"]
                    combined.at[idx, "retry_matched_road_name"] = result["matched_road_name"]
                    combined.at[idx, "retry_match_score"] = result["match_score"]
                else:
                    combined.at[idx, "retry_status"] = "still_unresolved"
                    combined.at[idx, "retry_reason"] = result["reason"]
                    combined.at[idx, "retry_candidate_phrase"] = result["candidate_phrase"]
                    combined.at[idx, "retry_matched_road_name"] = result["matched_road_name"]
                    combined.at[idx, "retry_match_score"] = result["match_score"]

            del roads

    combined["source_workbook"] = combined["source_workbook"]
    combined = combined.sort_values(["source_rank", "source_row_index"], kind="stable").reset_index(drop=True)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_columns = [
        "No",
        "Provinsi SPPG",
        "Kab./Kota SPPG",
        "Kecamatan SPPG",
        "Kelurahan/Desa SPPG",
        "latitude",
        "longitude",
        "Alamat SPPG",
        "Nama SPPG",
        "source_workbook",
        "source_rank",
        "source_row_index",
        "retry_status",
        "retry_reason",
        "retry_candidate_phrase",
        "retry_matched_road_name",
        "retry_match_score",
    ]
    main_df = combined.loc[:, output_columns]
    still_unresolved = main_df.loc[main_df["latitude"].isna() | main_df["longitude"].isna()].copy()

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        main_df.to_excel(writer, index=False, sheet_name="National")
        still_unresolved.to_excel(writer, index=False, sheet_name="still_unresolved")

    print(f"wrote {OUTPUT_PATH.resolve()}")
    print(f"rows={len(main_df)} still_unresolved={len(still_unresolved)}")


if __name__ == "__main__":
    main()
