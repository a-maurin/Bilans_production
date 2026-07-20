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
from pathlib import Path

def restaurer():
    root = Path(__file__).resolve().parents[1]
    zip_path = root / "pack_configuration_referentiels.zip"
    
    if not zip_path.exists():
        print("Erreur : pack_configuration_referentiels.zip introuvable à la racine.")
        return

    print("Décompression de pack_configuration_referentiels.zip en cours...")
    with zipfile.ZipFile(zip_path, 'r') as z:
        # Extraire tout à la racine (les chemins dans le zip devraient commencer par ref/ ou data/)
        z.extractall(root)
    
    print("Décompression terminée ! Le dossier ref/ devrait être complet.")

if __name__ == "__main__":
    restaurer()
