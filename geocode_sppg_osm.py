#!/usr/bin/env python3
from __future__ import annotations

import argparse
import gzip
import pickle
import re
import unicodedata
from functools import lru_cache
from collections import Counter, defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

import geopandas as gpd
import pandas as pd
from openpyxl import Workbook, load_workbook
from pyrosm import OSM
from shapely.geometry.base import BaseGeometry
from shapely.ops import unary_union
from shapely.prepared import prep


SOURCE_WORKBOOK = Path("bgn_sppg_operasional.xlsx")
DEFAULT_OUTPUT_WORKBOOK = Path("bgn_sppg_operasional_geocoded.xlsx")
DEFAULT_CACHE_DIR = Path("/private/tmp/sppg_osm_cache")
NETWORK_TYPE = "driving+service"

INPUT_COLUMNS = [
    "No",
    "Provinsi SPPG",
    "Kab./Kota SPPG",
    "Kecamatan SPPG",
    "Kelurahan/Desa SPPG",
    "Alamat SPPG",
    "Nama SPPG",
]

ADMIN_LEVELS = {
    "province": {4},
    "kabkota": {5},
    "kecamatan": {6},
    "kelurahan": {7, 8, 9},
}

DROP_PREFIX_TOKENS = {
    "JL",
    "JLN",
    "JALAN",
    "JALN",
    "JLAN",
    "GG",
    "GANG",
    "KOMP",
    "KOMPLEK",
    "KOMPLEKS",
    "PERUM",
    "PERUMAHAN",
    "BTN",
    "BLK",
    "BLOK",
    "RUKO",
    "RUKAN",
}

CANDIDATE_TERMINATORS = {
    "NO",
    "NOMOR",
    "RT",
    "RW",
    "KEL",
    "KELURAHAN",
    "KEC",
    "KECAMATAN",
    "KAB",
    "KABUPATEN",
    "PROV",
    "PROVINSI",
    "DEPAN",
    "BELAKANG",
    "SAMPING",
    "SEBELAH",
    "DALAM",
    "LUAR",
    "LINGKUNGAN",
    "LINGK",
    "KM",
}

ROAD_BLACKLIST = {
    "",
    "UNNAMED ROAD",
    "NO NAME",
    "TANPA NAMA",
    "JALAN TANPA NAMA",
    "UNKNOWN",
}

PROVINCE_GROUPS = {
    "sumatra": {
        "ACEH",
        "SUMATERA UTARA",
        "SUMATERA BARAT",
        "RIAU",
        "JAMBI",
        "SUMATERA SELATAN",
        "BENGKULU",
        "LAMPUNG",
        "KEPULAUAN RIAU",
        "KEPULAUAN BANGKA BELITUNG",
    },
    "java": {
        "BANTEN",
        "DKI JAKARTA",
        "JAKARTA",
        "JAWA BARAT",
        "JAWA TENGAH",
        "JAWA TIMUR",
        "DAERAH ISTIMEWA YOGYAKARTA",
        "YOGYAKARTA",
        "BALI",
    },
    "kalimantan": {
        "KALIMANTAN BARAT",
        "KALIMANTAN TENGAH",
        "KALIMANTAN SELATAN",
        "KALIMANTAN TIMUR",
        "KALIMANTAN UTARA",
    },
    "sulawesi": {
        "SULAWESI UTARA",
        "SULAWESI TENGAH",
        "SULAWESI SELATAN",
        "SULAWESI TENGGARA",
        "SULAWESI BARAT",
        "GORONTALO",
    },
    "nusa-tenggara": {
        "NUSA TENGGARA BARAT",
        "NUSA TENGGARA TIMUR",
    },
    "maluku": {
        "MALUKU",
        "MALUKU UTARA",
    },
    "papua": {
        "PAPUA",
        "PAPUA BARAT",
        "PAPUA BARAT DAYA",
        "PAPUA TENGAH",
        "PAPUA SELATAN",
        "PAPUA PEGUNUNGAN",
    },
}

MAX_CANDIDATE_TOKENS = 8
MIN_MATCH_SCORE = 0.82
MIN_MATCH_MARGIN = 0.05


@dataclass(frozen=True)
class BoundaryCandidate:
    level: int
    source_index: int
    name_raw: str
    name_norm: str
    geometry: BaseGeometry
    score: float


@dataclass(frozen=True)
class AdminPathResult:
    province: BoundaryCandidate
    kabkota: BoundaryCandidate
    kecamatan: BoundaryCandidate | None = None
    kelurahan: BoundaryCandidate | None = None

    @property
    def boundary(self) -> BaseGeometry:
        if self.kelurahan is not None:
            return self.kelurahan.geometry
        if self.kecamatan is not None:
            return self.kecamatan.geometry
        return self.kabkota.geometry

    @property
    def admin_level_used(self) -> str:
        if self.kelurahan is not None:
            return "kelurahan"
        if self.kecamatan is not None:
            return "kecamatan"
        return "kabkota"


@dataclass(frozen=True)
class RoadMatchResult:
    success: bool
    latitude: float | None
    longitude: float | None
    road_name: str | None
    candidate_phrase: str | None
    score: float | None
    candidate_count: int
    reason: str


@dataclass
class PreparedExtract:
    stem: str
    boundaries: dict[int, gpd.GeoDataFrame]
    road_geoms: dict[str, BaseGeometry]
    road_tokens: dict[str, tuple[str, ...]]
    token_index: dict[str, set[str]]


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = text.replace("&", " AND ")
    text = re.sub(r"[/,;:()\-]", " ", text)
    text = re.sub(r"[.']", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def collapse_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def collapse_spaced_letters(text: str) -> str:
    tokens = text.split()
    if len(tokens) > 1 and all(len(token) == 1 for token in tokens):
        return "".join(tokens)
    return text


def normalize_admin_name(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    text = collapse_spaced_letters(text)
    text = re.sub(r"\bDAERAH ISTIMEWA\b", " ", text)
    text = re.sub(r"\bDAERAH KHUSUS IBUKOTA\b", " ", text)
    text = re.sub(r"\bPROVINSI\b", " ", text)
    text = re.sub(r"\bKABUPATEN\b", " ", text)
    text = re.sub(r"\bKAB\b", " ", text)
    text = re.sub(r"\bKOTA ADMINISTRASI\b", " ", text)
    text = re.sub(r"\bKOTA ADM\b", " ", text)
    text = re.sub(r"\bKOTA\b", " ", text)
    text = re.sub(r"\bKECAMATAN\b", " ", text)
    text = re.sub(r"\bKEC\b", " ", text)
    text = re.sub(r"\bKELURAHAN\b", " ", text)
    text = re.sub(r"\bKEL\b", " ", text)
    text = re.sub(r"\bDESA\b", " ", text)
    text = re.sub(r"\bDS\b", " ", text)
    text = re.sub(r"\bDSN\b", " ", text)
    text = collapse_spaces(text)
    return text


def detect_admin_kind(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return "unknown"
    if re.search(r"\bKOTA\b", text) or "KOTA ADM" in text:
        return "city"
    if re.search(r"\bKABUPATEN\b", text) or re.search(r"\bKAB\b", text):
        return "regency"
    return "unknown"


def parse_admin_level(value: object) -> int | None:
    if value is None:
        return None
    text = normalize_text(value)
    match = re.search(r"\d+", text)
    if not match:
        return None
    try:
        return int(match.group(0))
    except ValueError:
        return None


def road_prefix_tokens(tokens: list[str]) -> list[str]:
    if not tokens:
        return tokens
    start = 0
    while start < len(tokens) and tokens[start] in DROP_PREFIX_TOKENS:
        start += 1
    if start < len(tokens) and tokens[start] == "RAYA" and start > 0:
        start += 1
    return tokens[start:]


def normalize_road_name(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    tokens = road_prefix_tokens(text.split())
    text = collapse_spaces(" ".join(tokens))
    if text in ROAD_BLACKLIST:
        return ""
    return text


def tokenize_for_index(value: str) -> tuple[str, ...]:
    tokens = []
    for token in value.split():
        if token in DROP_PREFIX_TOKENS:
            continue
        if token in CANDIDATE_TERMINATORS:
            continue
        if len(token) < 3 and not token.isdigit():
            continue
        tokens.append(token)
    return tuple(tokens)


def aliases_for_text(value: object) -> list[str]:
    raw = normalize_text(value)
    if not raw:
        return []
    aliases = {raw, normalize_admin_name(raw)}
    return [alias for alias in aliases if alias]


def build_context_phrases(row: pd.Series) -> list[str]:
    phrases: list[str] = []
    for column in [
        "Provinsi SPPG",
        "Kab./Kota SPPG",
        "Kecamatan SPPG",
        "Kelurahan/Desa SPPG",
    ]:
        phrases.extend(aliases_for_text(row[column]))
    return list(dict.fromkeys(p for p in phrases if p))


def candidate_phrases(address: object) -> list[str]:
    text = normalize_text(address)
    if not text:
        return []

    segments = [segment for segment in re.split(r"[,\n;|]+", text) if segment.strip()]
    if not segments:
        segments = [text]

    phrases: list[str] = []
    for segment in segments:
        tokens = segment.split()
        if not tokens:
            continue

        variants = [tokens]
        if tokens and tokens[0] in DROP_PREFIX_TOKENS:
            variants.append(tokens[1:])
            if len(tokens) > 1 and tokens[1] == "RAYA":
                variants.append(tokens[2:])
        if tokens and tokens[0] == "RAYA":
            variants.append(tokens[1:])

        for variant in variants:
            variant = list(variant)
            if not variant:
                continue
            cutoff = len(variant)
            for idx, token in enumerate(variant):
                if token in CANDIDATE_TERMINATORS:
                    cutoff = idx
                    break
            variant = variant[:cutoff]
            if not variant:
                continue
            max_tokens = min(len(variant), MAX_CANDIDATE_TOKENS)
            for end in range(max_tokens, 0, -1):
                phrase = collapse_spaces(" ".join(variant[:end]))
                if phrase:
                    phrases.append(phrase)

    ordered = list(dict.fromkeys(phrases))
    ordered = [phrase for _, phrase in sorted(enumerate(ordered), key=lambda item: (-len(item[1].split()), item[0]))]
    return ordered


def compact_text(value: str) -> str:
    return value.replace(" ", "")


def road_similarity(candidate: str, road_name: str) -> float:
    candidate = normalize_road_name(candidate)
    road_name = normalize_road_name(road_name)
    if not candidate or not road_name:
        return 0.0

    if candidate == road_name:
        return 1.0

    candidate_compact = compact_text(candidate)
    road_compact = compact_text(road_name)
    if candidate_compact == road_compact:
        return 1.0
    if candidate_compact in road_compact or road_compact in candidate_compact:
        return 0.98

    candidate_tokens = set(tokenize_for_index(candidate))
    road_tokens = set(tokenize_for_index(road_name))
    if candidate_tokens and road_tokens:
        overlap = len(candidate_tokens & road_tokens)
        token_score = overlap / max(len(candidate_tokens), len(road_tokens))
    else:
        token_score = 0.0

    sequence_score = SequenceMatcher(None, candidate_compact, road_compact).ratio()
    return max(sequence_score, token_score)


def build_road_index(roads: gpd.GeoDataFrame) -> tuple[dict[str, BaseGeometry], dict[str, tuple[str, ...]], dict[str, set[str]]]:
    road_geoms: dict[str, BaseGeometry] = {}
    road_tokens: dict[str, tuple[str, ...]] = {}
    token_index: dict[str, set[str]] = defaultdict(set)

    grouped = roads.groupby("road_name_norm", sort=False)
    for road_name, group in grouped:
        geometries = [geom for geom in group.geometry if geom is not None and not geom.is_empty]
        if not geometries:
            continue
        merged = unary_union(geometries)
        if merged.is_empty:
            continue
        road_geoms[road_name] = merged
        tokens = tokenize_for_index(road_name)
        road_tokens[road_name] = tokens
        for token in tokens:
            token_index[token].add(road_name)

    return road_geoms, road_tokens, token_index


def write_cache(path: Path, payload: PreparedExtract) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with gzip.open(path, "wb") as fh:
        pickle.dump(payload, fh, protocol=pickle.HIGHEST_PROTOCOL)


def read_cache(path: Path) -> PreparedExtract:
    with gzip.open(path, "rb") as fh:
        payload = pickle.load(fh)
    if not isinstance(payload, PreparedExtract):
        raise TypeError(f"Unexpected cache payload in {path}")
    return payload


def cache_path_for_extract(cache_dir: Path, stem: str) -> Path:
    return cache_dir / stem / "prepared_extract.pkl.gz"


@lru_cache(maxsize=1)
def province_group_lookup() -> dict[str, str]:
    lookup: dict[str, str] = {}
    for group, provinces in PROVINCE_GROUPS.items():
        for province in provinces:
            lookup[normalize_admin_name(province)] = group
    return lookup


def determine_extract_group(province: object) -> str:
    province_norm = normalize_admin_name(province)
    return province_group_lookup().get(province_norm, "sumatra")


def extract_group_for_province(province: object) -> str:
    return determine_extract_group(province)


def load_or_build_extract(pbf_path: Path, cache_dir: Path, rebuild: bool = False) -> PreparedExtract:
    cache_path = cache_path_for_extract(cache_dir, pbf_path.stem)
    if cache_path.exists() and not rebuild:
        return read_cache(cache_path)

    print(f"[extract] building {pbf_path.name}")
    osm = OSM(str(pbf_path))

    boundaries = osm.get_boundaries(boundary_type="administrative", extra_attributes=["admin_level"])
    if not isinstance(boundaries, gpd.GeoDataFrame):
        boundaries = gpd.GeoDataFrame(boundaries)
    boundaries = boundaries[boundaries.geometry.notna()].copy()
    boundaries = boundaries[boundaries.geometry.geom_type.isin({"Polygon", "MultiPolygon"})].copy()
    boundaries = boundaries[boundaries["name"].notna()].copy()
    boundaries["admin_level_int"] = boundaries["admin_level"].map(parse_admin_level)
    boundaries["name_raw"] = boundaries["name"].astype(str)
    boundaries["name_norm"] = boundaries["name_raw"].map(normalize_admin_name)
    boundaries["kind"] = boundaries["name_raw"].map(detect_admin_kind)
    boundaries = boundaries[boundaries["admin_level_int"].isin({4, 5, 6, 7, 8, 9})].copy()
    boundaries = boundaries[boundaries["name_norm"].ne("")].copy()
    boundaries = boundaries.reset_index(drop=True)

    boundary_tables: dict[int, gpd.GeoDataFrame] = {}
    for level in sorted({4, 5, 6, 7, 8, 9}):
        table = boundaries.loc[
            boundaries["admin_level_int"] == level,
            ["name_raw", "name_norm", "kind", "admin_level_int", "geometry"],
        ].copy()
        table = table.reset_index(drop=True)
        boundary_tables[level] = table

    roads = osm.get_network(network_type=NETWORK_TYPE, nodes=False)
    if not isinstance(roads, gpd.GeoDataFrame):
        roads = gpd.GeoDataFrame(roads)
    roads = roads[roads.geometry.notna()].copy()
    roads = roads[roads["name"].notna()].copy()
    roads["name_raw"] = roads["name"].astype(str)
    roads["road_name_norm"] = roads["name_raw"].map(normalize_road_name)
    roads = roads[roads["road_name_norm"].ne("")].copy()
    roads = roads[~roads["road_name_norm"].isin(ROAD_BLACKLIST)].copy()
    roads = roads.reset_index(drop=True)

    road_geoms, road_tokens, token_index = build_road_index(roads)

    payload = PreparedExtract(
        stem=pbf_path.stem,
        boundaries=boundary_tables,
        road_geoms=road_geoms,
        road_tokens=road_tokens,
        token_index=token_index,
    )
    write_cache(cache_path, payload)
    return payload


def admin_candidates(
    table: gpd.GeoDataFrame,
    query: object,
    parent_geometry: BaseGeometry | None = None,
) -> list[BoundaryCandidate]:
    query_norm = normalize_admin_name(query)
    if not query_norm:
        return []

    compact_query = compact_text(query_norm)
    candidates = table.loc[table["name_norm"] == query_norm].copy()
    if candidates.empty:
        candidates = table.loc[
            table["name_norm"].map(lambda value: compact_text(value) == compact_query or compact_query in compact_text(value) or compact_text(value) in compact_query)
        ].copy()

    if candidates.empty:
        return []

    result: list[BoundaryCandidate] = []
    prepared_parent = prep(parent_geometry) if parent_geometry is not None else None
    for index, row in candidates.iterrows():
        geometry = row["geometry"]
        if geometry is None or geometry.is_empty:
            continue
        if prepared_parent is not None:
            try:
                if not prepared_parent.intersects(geometry):
                    continue
            except Exception:
                continue
        score = 1.0 if row["name_norm"] == query_norm else 0.95
        result.append(
            BoundaryCandidate(
                level=int(row["admin_level_int"]) if "admin_level_int" in row else 0,
                source_index=int(index),
                name_raw=str(row["name_raw"]),
                name_norm=str(row["name_norm"]),
                geometry=geometry,
                score=score,
            )
        )
    result.sort(key=lambda candidate: (-candidate.score, -candidate.geometry.area))
    return result


def road_name_candidates(
    extract: PreparedExtract,
    phrase: str,
    boundary_geometry: BaseGeometry,
) -> list[tuple[str, float, int]]:
    phrase_norm = normalize_road_name(phrase)
    if not phrase_norm:
        return []

    exact = []
    if phrase_norm in extract.road_geoms:
        exact.append((phrase_norm, 1.0, 10_000))

    if exact:
        return exact

    tokens = tokenize_for_index(phrase_norm)
    if not tokens:
        return []

    candidate_counts: Counter[str] = Counter()
    for token in tokens:
        for road_name in extract.token_index.get(token, set()):
            candidate_counts[road_name] += 1

    if not candidate_counts:
        return []

    min_overlap = 1 if len(tokens) == 1 else 2 if len(tokens) == 2 else 2
    candidates = [
        road_name
        for road_name, count in candidate_counts.items()
        if count >= min_overlap
    ]
    if not candidates:
        candidates = [name for name, _ in candidate_counts.most_common(100)]

    scored: list[tuple[str, float, int]] = []
    for road_name in candidates:
        score = road_similarity(phrase_norm, road_name)
        if score <= 0:
            continue
        scored.append((road_name, score, candidate_counts[road_name]))

    scored.sort(key=lambda item: (-item[1], -item[2], item[0]))
    return scored


def choose_road_geometry(
    extract: PreparedExtract,
    road_name: str,
    boundary_geometry: BaseGeometry,
) -> BaseGeometry | None:
    road_geometry = extract.road_geoms.get(road_name)
    if road_geometry is None or road_geometry.is_empty:
        return None
    try:
        clipped = road_geometry.intersection(boundary_geometry)
    except Exception:
        return None
    if clipped.is_empty:
        return None
    return clipped


def geocode_with_path(
    extract: PreparedExtract,
    path: AdminPathResult,
    address: object,
) -> RoadMatchResult:
    phrases = candidate_phrases(address)
    if not phrases:
        return RoadMatchResult(False, None, None, None, None, None, 0, "no_address_candidate")

    boundary_geometry = path.boundary
    best: tuple[str, str, float, int, BaseGeometry] | None = None
    candidate_total = 0

    for phrase in phrases:
        scored_candidates = road_name_candidates(extract, phrase, boundary_geometry)
        candidate_total += len(scored_candidates)
        if not scored_candidates:
            continue

        top_score = scored_candidates[0][1]
        second_score = scored_candidates[1][1] if len(scored_candidates) > 1 else 0.0
        if top_score < MIN_MATCH_SCORE or (top_score - second_score) < MIN_MATCH_MARGIN:
            continue

        road_name, score, overlap = scored_candidates[0]
        clipped = choose_road_geometry(extract, road_name, boundary_geometry)
        if clipped is None:
            continue
        best = (phrase, road_name, score, overlap, clipped)
        break

    if best is None:
        return RoadMatchResult(False, None, None, None, None, None, candidate_total, "road_candidate_not_matched")

    phrase, road_name, score, overlap, clipped = best
    point = clipped.representative_point()
    return RoadMatchResult(
        True,
        float(point.y),
        float(point.x),
        road_name,
        phrase,
        score,
        candidate_total,
        "matched",
    )


def boundary_search_order() -> list[int]:
    return [4, 5, 6, 7, 8, 9]


def resolve_admin_paths(extract: PreparedExtract, row: pd.Series) -> list[AdminPathResult]:
    province_table = extract.boundaries.get(4, gpd.GeoDataFrame())
    kab_table = extract.boundaries.get(5, gpd.GeoDataFrame())
    kec_table = extract.boundaries.get(6, gpd.GeoDataFrame())
    kel_tables = [extract.boundaries.get(level, gpd.GeoDataFrame()) for level in (7, 8, 9)]

    province_candidates = admin_candidates(province_table, row["Provinsi SPPG"])
    if not province_candidates:
        return []

    results: list[AdminPathResult] = []
    for province in province_candidates:
        kab_candidates = admin_candidates(kab_table, row["Kab./Kota SPPG"], province.geometry)
        if not kab_candidates:
            continue

        for kabkota in kab_candidates:
            kec_candidates = admin_candidates(kec_table, row["Kecamatan SPPG"], kabkota.geometry)
            if not kec_candidates:
                results.append(AdminPathResult(province=province, kabkota=kabkota))
                continue

            for kecamatan in kec_candidates:
                kel_candidates: list[BoundaryCandidate] = []
                for kel_table in kel_tables:
                    kel_candidates.extend(admin_candidates(kel_table, row["Kelurahan/Desa SPPG"], kecamatan.geometry))

                if kel_candidates:
                    for kelurahan in kel_candidates:
                        results.append(
                            AdminPathResult(
                                province=province,
                                kabkota=kabkota,
                                kecamatan=kecamatan,
                                kelurahan=kelurahan,
                            )
                        )
                else:
                    results.append(
                        AdminPathResult(
                            province=province,
                            kabkota=kabkota,
                            kecamatan=kecamatan,
                        )
                    )

    results.sort(
        key=lambda path: (
            1 if path.kelurahan is not None else 0,
            1 if path.kecamatan is not None else 0,
            -path.kabkota.geometry.area,
        ),
        reverse=True,
    )
    return results


def geocode_row(extract: PreparedExtract, row: pd.Series) -> dict[str, object]:
    path_candidates = resolve_admin_paths(extract, row)
    if not path_candidates:
        return {
            "latitude": None,
            "longitude": None,
            "status": "unresolved",
            "reason": "admin_path_not_found",
            "admin_level_used": None,
            "best_candidate": None,
            "matched_road_name": None,
            "candidate_phrase": None,
            "match_score": None,
            "candidate_count": 0,
        }

    best_unresolved_reason = "admin_path_not_found"
    for path in path_candidates:
        road_result = geocode_with_path(extract, path, row["Alamat SPPG"])
        if road_result.success:
            return {
                "latitude": road_result.latitude,
                "longitude": road_result.longitude,
                "status": "matched",
                "reason": "matched",
                "admin_level_used": path.admin_level_used,
                "best_candidate": road_result.road_name,
                "matched_road_name": road_result.road_name,
                "candidate_phrase": road_result.candidate_phrase,
                "match_score": road_result.score,
                "candidate_count": road_result.candidate_count,
            }
        best_unresolved_reason = road_result.reason

    return {
        "latitude": None,
        "longitude": None,
        "status": "unresolved",
        "reason": best_unresolved_reason,
        "admin_level_used": path_candidates[0].admin_level_used if path_candidates else None,
        "best_candidate": None,
        "matched_road_name": None,
        "candidate_phrase": None,
        "match_score": None,
        "candidate_count": 0,
    }


def load_input_rows(workbook_path: Path) -> tuple[str, pd.DataFrame]:
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            raise ValueError(f"Workbook {workbook_path} is empty")
        header = list(rows[0])
        data = pd.DataFrame(rows[1:], columns=header)
        data["source_excel_row"] = range(2, len(data) + 2)
        return ws.title, data
    finally:
        wb.close()


def write_output_workbook(
    rows: pd.DataFrame,
    results: list[dict[str, object]],
    output_path: Path,
    sheet_name: str,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    ws = workbook.active
    ws.title = sheet_name

    address_index = INPUT_COLUMNS.index("Alamat SPPG")
    output_columns = INPUT_COLUMNS[:address_index] + ["latitude", "longitude"] + INPUT_COLUMNS[address_index:]
    ws.append(output_columns)
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 12

    for row, result in zip(rows.to_dict(orient="records"), results):
        ws.append(
            [
                row["No"],
                row["Provinsi SPPG"],
                row["Kab./Kota SPPG"],
                row["Kecamatan SPPG"],
                row["Kelurahan/Desa SPPG"],
                result["latitude"],
                result["longitude"],
                row["Alamat SPPG"],
                row["Nama SPPG"],
            ]
        )

    unresolved_sheet_name = "unresolved_rows"
    review_ws = workbook.create_sheet(unresolved_sheet_name)

    review_headers = [
        "source_excel_row",
        "No",
        "Provinsi SPPG",
        "Kab./Kota SPPG",
        "Kecamatan SPPG",
        "Kelurahan/Desa SPPG",
        "Alamat SPPG",
        "Nama SPPG",
        "status",
        "reason",
        "admin_level_used",
        "candidate_phrase",
        "matched_road_name",
        "match_score",
        "candidate_count",
    ]
    review_ws.append(review_headers)

    for idx, result in enumerate(results, start=2):
        if result["status"] == "matched":
            continue
        source = result["source_row"]
        review_ws.append(
            [
                source["source_excel_row"],
                source["No"],
                source["Provinsi SPPG"],
                source["Kab./Kota SPPG"],
                source["Kecamatan SPPG"],
                source["Kelurahan/Desa SPPG"],
                source["Alamat SPPG"],
                source["Nama SPPG"],
                result["status"],
                result["reason"],
                result["admin_level_used"],
                result["candidate_phrase"],
                result["matched_road_name"],
                result["match_score"],
                result["candidate_count"],
            ]
        )

    workbook.save(output_path)


def default_output_path(group: str | None) -> Path:
    if group is None:
        return DEFAULT_OUTPUT_WORKBOOK
    suffix = group.replace("-", "_")
    return DEFAULT_OUTPUT_WORKBOOK.with_name(f"{DEFAULT_OUTPUT_WORKBOOK.stem}_{suffix}.xlsx")


def extract_path_for_group(group: str) -> Path:
    return {
        "sumatra": Path("/Users/rizzie/Work/IndonesiaRe/data/openstreetmap/sumatra-260409.osm.pbf"),
        "java": Path("/Users/rizzie/Work/IndonesiaRe/data/openstreetmap/java-260409.osm.pbf"),
        "kalimantan": Path("/Users/rizzie/Work/IndonesiaRe/data/openstreetmap/kalimantan-260409.osm.pbf"),
        "sulawesi": Path("/Users/rizzie/Work/IndonesiaRe/data/openstreetmap/sulawesi-260409.osm.pbf"),
        "nusa-tenggara": Path("/Users/rizzie/Work/IndonesiaRe/data/openstreetmap/nusa-tenggara-260409.osm.pbf"),
        "maluku": Path("/Users/rizzie/Work/IndonesiaRe/data/openstreetmap/maluku-260409.osm.pbf"),
        "papua": Path("/Users/rizzie/Work/IndonesiaRe/data/openstreetmap/papua-260409.osm.pbf"),
    }[group]


def add_source_row_refs(rows: pd.DataFrame) -> pd.DataFrame:
    enriched = rows.copy()
    enriched["source_row"] = enriched.to_dict(orient="records")
    return enriched


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Geocode SPPG workbook using local OSM extracts.")
    parser.add_argument(
        "--input",
        type=Path,
        default=SOURCE_WORKBOOK,
        help="Input XLSX workbook path.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output XLSX workbook path.",
    )
    parser.add_argument(
        "--group",
        choices=list(PROVINCE_GROUPS.keys()),
        default=None,
        help="Process only one island extract and write one workbook for that group.",
    )
    parser.add_argument(
        "--cache-dir",
        type=Path,
        default=DEFAULT_CACHE_DIR,
        help="Directory for cached OSM extracts.",
    )
    parser.add_argument(
        "--rebuild-cache",
        action="store_true",
        help="Ignore cached extract files and rebuild them.",
    )
    args = parser.parse_args(argv)

    sheet_name, df = load_input_rows(args.input)
    missing = [column for column in INPUT_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"Workbook {args.input} is missing expected columns: {missing}")

    df = df.copy()
    df["extract_group"] = df["Provinsi SPPG"].map(determine_extract_group)
    if args.group is not None:
        df = df.loc[df["extract_group"] == args.group].copy()
        if df.empty:
            raise ValueError(f"No rows matched the requested group: {args.group}")

    if args.output is None:
        args.output = default_output_path(args.group)

    df["source_row"] = df.to_dict(orient="records")

    unique_groups = [args.group] if args.group is not None else [group for group in df["extract_group"].dropna().unique().tolist()]
    prepared_extracts: dict[str, PreparedExtract] = {}
    for group in unique_groups:
        pbf_path = extract_path_for_group(group)
        prepared_extracts[group] = load_or_build_extract(pbf_path, args.cache_dir, rebuild=args.rebuild_cache)

    results: list[dict[str, object]] = []
    total = len(df)
    for idx, row in df.iterrows():
        group = row["extract_group"]
        extract = prepared_extracts[group]
        geocode_result = geocode_row(extract, row)
        geocode_result["source_row"] = row["source_row"]
        results.append(geocode_result)
        if (idx + 1) % 500 == 0 or idx + 1 == total:
            matched = sum(1 for item in results if item["status"] == "matched")
            unresolved = len(results) - matched
            print(f"[rows] {idx + 1}/{total} matched={matched} unresolved={unresolved}")

    write_output_workbook(df, results, args.output, sheet_name)
    matched = sum(1 for item in results if item["status"] == "matched")
    unresolved = len(results) - matched
    print(f"[done] matched={matched} unresolved={unresolved}")
    print(f"[done] wrote {args.output.resolve()}")


if __name__ == "__main__":
    main()
