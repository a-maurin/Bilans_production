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
from core.engine.orchestrateur_profils import load_profile_config
from core.common.chargeurs_donnees import load_pej

project_root = Path('.')
try:
    p = load_profile_config(project_root, 'ppp')
    df_pej = load_pej(project_root, echelle='national', code='', date_deb='2025-01-01', date_fin='2025-12-31')
    
    natinf_pej = p.get('natinf_pej', [])
    print("NATINF in profile:", natinf_pej[:5], "...")
    
    import re
    from core.common.utilitaires_metier import series_str_contains
    
    pattern = "|".join(rf"(?:^|_){re.escape(c)}(?:_|$)" for c in natinf_pej)
    natinf_col = "NATINF_PEJ" if "NATINF_PEJ" in df_pej.columns else "NATINF"
    print("NATINF column:", natinf_col)
    
    if natinf_col in df_pej.columns:
        res = df_pej[series_str_contains(df_pej[natinf_col], pattern, regex=True)]
        print("Filtered PEJ length:", len(res))
        if len(res) == 0:
            print("Why 0?")
            print("Sample NATINF_PEJ values:", df_pej[natinf_col].dropna().head().tolist())
    else:
        print("No natinf col found.")
except Exception as e:
    print("Error:", e)
