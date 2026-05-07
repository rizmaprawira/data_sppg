#!/usr/bin/env python3
from __future__ import annotations

import sys

from sppg_road_geocoder import main


if __name__ == "__main__":
    main(["--island", "nusa-tenggara", *sys.argv[1:]])
