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
from core.common.chargeurs_donnees import load_pve, load_pej
from core.engine.agregations_profil import _build_global_proc_detail

root = Path('.')
pve = load_pve(root)
pej = load_pej(root)

print(pve[['INF-DATE-INTG', 'INF-DATE-MIF']].head())

pve_detail = _build_global_proc_detail(
    pve, 'PVe', ['INF-ID'], 
    ['INF-DATE', 'INF-DATE-INTG', 'INF-DATE-MIF', 'INF-DATE-I', 'INF_DATE', 'DATE_FAITS'], 
    ['COMMUNE_LIB', 'INF-LIEU', 'COMMUNE', 'NOM_COM', 'INF-INSEE', 'INSEE_DEP'], 
    ['INF-NATINF'], ['DOMAINE']
)

print("PVE DETAIL:")
print(pve_detail[['date', 'commune']].head())

pej_detail = _build_global_proc_detail(
    pej, 'PEJ', ['DC_ID'], 
    ['DATE_REF'], 
    ['COMMUNE', 'nom_commune', 'INF-LIEU', 'INF-INSEE'], 
    ['THEME'], ['DOMAINE']
)
print("PEJ DETAIL:")
print(pej_detail[['date', 'commune']].head())
