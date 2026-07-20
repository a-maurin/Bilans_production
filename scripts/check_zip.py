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

import zipfile
import sys
from pathlib import Path

root = Path(__file__).resolve().parents[1]
zip_path = root / "pack_configuration_referentiels.zip"

if zip_path.exists():
    with zipfile.ZipFile(zip_path, 'r') as z:
        files = z.namelist()
        ppp = [f for f in files if "natinf_ppp" in f]
        print(f"Trouvé dans le zip : {ppp}")
else:
    print("Zip introuvable")
