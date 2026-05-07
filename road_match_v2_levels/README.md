# Road Match V2 Levels

Workflow baru untuk geocoding SPPG berbasis geometri jalan OSM, tetap diproses per pulau dan per `kab/kota`.

Level `Match_By`:

- `jalan`: nama jalan dari alamat berhasil dicocokkan di dalam boundary `kabupaten/kota`.
- `kelurahan/desa`: jika level jalan gagal, ambil titik acak dari geometri jalan di dalam boundary `kelurahan/desa`.
- `kecamatan`: jika level kelurahan/desa gagal, ambil titik acak dari geometri jalan di dalam boundary `kecamatan`.
- `kabupaten/kota`: fallback terakhir, ambil titik acak dari geometri jalan di dalam boundary `kabupaten/kota`.

Semua titik diambil dari geometri `jalan`, bukan centroid polygon.

## File utama

- `sppg_road_geocoder.py`: core geocoder generik.
- `geocode_sppg_osm_*_v2_roads_levels.py`: wrapper per pulau.
- `combine_national_geocoded_v2_roads_levels.py`: gabung hasil regional.
- `output/`: workbook hasil.

## Jalankan

```bash
python road_match_v2_levels/geocode_sppg_osm_java_v2_roads_levels.py --restart
bash road_match_v2_levels/run_all_islands_v2_roads_levels.sh --restart
```
