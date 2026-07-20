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
from pathlib import Path
sys.path.insert(0, 'src')
import pandas as pd
from core.common.chargeurs_donnees import load_pej, load_point_ctrl, merge_pej_faits_locations
from core.engine.agregations_profil import _build_global_proc_detail

root = Path('.')
point = load_point_ctrl(root)
pej = load_pej(root)

print("Total PEJ rows raw:", len(pej))
# Simulate orchestrateur merging
pej_dept = merge_pej_faits_locations(pej, root, dept_code="21")

pej_detail_old = _build_global_proc_detail(
    pej_dept, "PEJ", ["DC_ID"], ["DATE_REF"], ["COMMUNE", "nom_commune", "INF-LIEU", "INF-INSEE"], ["THEME"], ["DOMAINE"]
)

pej_detail_new = _build_global_proc_detail(
    pej_dept, "PEJ", ["DC_ID"], ["DATE_REF"], ["COMMUNE", "NOM_COM", "nom_commune", "INF-LIEU", "INF-INSEE"], ["THEME"], ["DOMAINE"]
)

mask_old = pej_detail_old["commune"].isin(["n.d.", "INC", "ND"])
mask_new = pej_detail_new["commune"].isin(["n.d.", "INC", "ND"])

print("Without NOM_COM, missing communes:", mask_old.sum())
print("With NOM_COM, missing communes:", mask_new.sum())
