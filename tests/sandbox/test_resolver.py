#
# Copyright (C) 2026 Aguirre MAURIN
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

import sys
import os

osgeo4w_root = r"C:\Program Files\QGIS 3.44.8"
sys.path.insert(0, os.path.join(osgeo4w_root, r"apps\qgis-ltr\python"))
sys.path.insert(0, os.path.join(osgeo4w_root, r"apps\Python312\Lib\site-packages"))

sys.path.insert(0, r"C:\Users\aguirre.maurin\Documents\GitHub\Bilans_production\src\bilans\cartographie")

from layer_resolver import resolve_layer_names

available_names = [
    'point_ctrl_2024_wgs84',
    'point_ctrl_20251231_wgs84',
    'point_ctrl_20260505_wgs84',
    'localisation_infrac_FAITS_20260505'
]

# Test 1: 2025 only
print("Test 1: 2025")
res1 = resolve_layer_names(
    configured_name="point_ctrl_20260205_wgs84",
    layer_role="point_controles",
    available_names=available_names,
    date_deb="2025-01-01",
    date_fin="2025-12-31"
)
print("->", res1)

# Test 2: 2024 to 2026
print("\nTest 2: 2024-2026")
res2 = resolve_layer_names(
    configured_name="point_ctrl_20260205_wgs84",
    layer_role="point_controles",
    available_names=available_names,
    date_deb="2024-01-01",
    date_fin="2026-12-31"
)
print("->", res2)

# Test 3: No dates
print("\nTest 3: No dates")
res3 = resolve_layer_names(
    configured_name="point_ctrl_20260205_wgs84",
    layer_role="point_controles",
    available_names=available_names,
    date_deb=None,
    date_fin=None
)
print("->", res3)
