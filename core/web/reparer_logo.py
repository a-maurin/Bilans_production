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

def generer_logo_blanc():
    root = Path(__file__).resolve().parents[2]
    src_path = root / "ref" / "programme" / "logos" / "bandeau_ofbilan.svg"
    dst_path = root / "ref" / "programme" / "logos" / "bandeau_ofbilan_blanc.svg"
    local_dst_path = Path(__file__).resolve().parent / "logo.svg"
    
    if not src_path.exists():
        print(f"Erreur : Le fichier source est introuvable : {src_path}")
        return False
        
    try:
        content = src_path.read_text(encoding="utf-8")
        # Remplacer la couleur du texte (#2c406e) par du blanc (#ffffff)
        modified = content.replace('#2c406e', '#ffffff')
        
        # Écriture dans le dossier logos d'origine
        dst_path.write_text(modified, encoding="utf-8")
        
        # Écriture dans le dossier web local
        local_dst_path.write_text(modified, encoding="utf-8")
        
        return True
    except Exception as e:
        print(f"Erreur lors de la génération du logo : {e}")
        return False

if __name__ == "__main__":
    generer_logo_blanc()
