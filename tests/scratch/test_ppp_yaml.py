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

try:
    project_root = Path('.')
    p = load_profile_config(project_root, 'ppp')
    with open('test_ppp.txt', 'w', encoding='utf-8') as f:
        f.write(f"sources: {p.get('sources')}\n")
        f.write(f"natinf_pej len: {len(p.get('natinf_pej', []))}\n")
        f.write(f"natinf_pej: {p.get('natinf_pej')}\n")
except Exception as e:
    with open('test_ppp.txt', 'w', encoding='utf-8') as f:
        f.write(f"Error: {e}\n")
