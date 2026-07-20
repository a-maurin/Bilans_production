#!/usr/bin/env python3
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

"""
Génère l'archive pack_configuration_referentiels.zip contenant les référentiels 
et les données brutes (ref/, data/sources/, data/sources_archive/).
Vérifie préalablement l'intégrité de la structure.
"""
from __future__ import annotations

import os
import sys
import zipfile
import subprocess
from pathlib import Path


def create_pack() -> int:
    project_root = Path(__file__).resolve().parents[1]
    
    # 1. Vérification
    print("Vérification de l'intégrité des référentiels...")
    verify_script = project_root / "scripts" / "verify_ref_layout.py"
    r = subprocess.run([sys.executable, str(verify_script), str(project_root)], cwd=project_root)
    if r.returncode != 0:
        print("Erreur : La vérification des référentiels a échoué. Corrigez les erreurs avant d'empaqueter.", file=sys.stderr)
        return r.returncode

    # 2. Empaquetage
    dist_dir = project_root / "distribution"
    dist_dir.mkdir(exist_ok=True)
    zip_path = dist_dir / "pack_configuration_referentiels.zip"
    
    print(f"\nCréation de l'archive : {zip_path}")
    
    targets = ["ref"]
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Assurer la structure pour data/sources
        zipf.writestr("data/sources/.gitkeep", "")
        
        for tgt in targets:
            src_dir = project_root / tgt
            if not src_dir.exists():
                print(f"Note : Le dossier {tgt} n'existe pas, ignoré.")
                continue
                
            for root, dirs, files in os.walk(src_dir):
                for file in files:
                    file_path = Path(root) / file
                    rel_path = file_path.relative_to(project_root)
                    print(f"Ajout : {rel_path}")
                    zipf.write(file_path, rel_path)

    print("\nArchive ZIP créée avec succès dans le dossier 'distribution'.")
    
    # 3. Copie du script d'installation
    import shutil
    installer_script = project_root / "scripts" / "installer_pack.bat"
    if installer_script.exists():
        shutil.copy2(installer_script, dist_dir / "installer_pack.bat")
        print(f"Script d'installation copié : {dist_dir / 'installer_pack.bat'}")
    
    return 0


if __name__ == "__main__":
    sys.exit(create_pack())
