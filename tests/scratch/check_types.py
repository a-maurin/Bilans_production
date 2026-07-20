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

import os
from pathlib import Path
import pandas as pd
from core.common.chargeurs_donnees import load_point_ctrl, load_pej

project_root = Path(os.getcwd())
df_pts = load_point_ctrl(project_root, echelle="national", code="FR")
print("Unique type_actio in point_ctrl:")
print(df_pts["type_actio"].dropna().unique())

df_pej = load_pej(project_root, echelle="national", code="FR")
print("\nUnique type_action in PEJ:")
if "TYPE_ACTION" in df_pej.columns:
    print(df_pej["TYPE_ACTION"].dropna().unique())
elif "type_action" in df_pej.columns:
    print(df_pej["type_action"].dropna().unique())
