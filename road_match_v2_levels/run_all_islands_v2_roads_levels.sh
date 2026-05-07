#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

python "$ROOT_DIR/geocode_sppg_osm_sumatra_v2_roads_levels.py" "$@"
python "$ROOT_DIR/geocode_sppg_osm_java_v2_roads_levels.py" "$@"
python "$ROOT_DIR/geocode_sppg_osm_kalimantan_v2_roads_levels.py" "$@"
python "$ROOT_DIR/geocode_sppg_osm_sulawesi_v2_roads_levels.py" "$@"
python "$ROOT_DIR/geocode_sppg_osm_nusa_tenggara_v2_roads_levels.py" "$@"
python "$ROOT_DIR/geocode_sppg_osm_maluku_v2_roads_levels.py" "$@"
python "$ROOT_DIR/geocode_sppg_osm_papua_v2_roads_levels.py" "$@"
python "$ROOT_DIR/combine_national_geocoded_v2_roads_levels.py"
