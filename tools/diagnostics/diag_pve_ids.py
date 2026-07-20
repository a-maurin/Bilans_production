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

import pandas as pd
from pathlib import Path
from ofbilan.common.chargeurs_donnees import load_pve, load_point_ctrl
from ofbilan.engine.orchestrateur_profils import _filter_pve, load_tub_pnf_codes, _apply_restrict_geo_tub
import yaml
import logging

root = Path('.')
tub_codes, _ = load_tub_pnf_codes(root)

with open('config/profils_bilan/tub.yaml', 'r', encoding='utf-8') as f:
    tub_cfg = yaml.safe_load(f)

pve = load_pve(root, dept_code='21', date_deb='2025-07-01', date_fin='2026-06-01')
pve_filtered = _filter_pve(pve, tub_cfg)

import geopandas as gpd
shp = root / 'ref' / 'programme' / 'sig' / 'communes_pnf' / 'communes_PNF_centroides.shp'
cen = gpd.read_file(shp)
cen_insee = cen['INSEE_COM'].astype(str).str.zfill(5).unique()

pve_filtered['_insee_pve'] = pve_filtered['INF-INSEE'].astype(str).str.zfill(5)

pve_tub = _apply_restrict_geo_tub(pve_filtered, pd.DataFrame(), root, tub_codes, 'PVE', log=logging.getLogger())

print('--- PVe forcés au centroïde ---')
count = 0
for _, r in pve_tub.iterrows():
    insee = r.get('_insee_pve')
    is_cen = insee in cen_insee
    if is_cen:
        count += 1
        ident = r.get('INF-ID', r.get('INF-NUM', r.get('DC_ID', 'Inconnu')))
        print(f"ID: {ident} | INSEE: {insee} | NATINF: {r.get('INF-NATINF')}")
print(f'Total: {count}')
