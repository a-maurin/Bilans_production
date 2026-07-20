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
import zipfile

def create_pack():
    # Déterminer la racine du projet (parent du dossier contenant ce script)
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    zip_path = os.path.join(project_root, "distribution", "pack_configuration_referentiels.zip")
    os.makedirs(os.path.join(project_root, "distribution"), exist_ok=True)
    
    # Chemins relatifs à la racine du projet
    targets = [
        ("ref/programme/tables_reference", "ref/programme/tables_reference"),
        ("ref/programme/sig", "ref/programme/sig"),
        ("ref/programme/modele_ofb", "ref/programme/modele_ofb"),
        ("ref/programme/logos", "ref/programme/logos"),
        ("data/sources", "data/sources"),
    ]
    
    print(f"Création de l'archive : {zip_path}")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Assurer la structure pour data/sources
        zipf.writestr("data/sources/.gitkeep", "")
        
        for src_rel, arc_dir in targets:
            src_dir = os.path.join(project_root, src_rel)
            if not os.path.exists(src_dir):
                print(f"Note : Le dossier {src_dir} n'existe pas, il sera ignoré.")
                continue
                
            for root, dirs, files in os.walk(src_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    # Calcul du chemin relatif dans l'archive
                    rel_path = os.path.relpath(file_path, src_dir)
                    arc_path = os.path.join(arc_dir, rel_path)
                    print(f"Ajout : {file_path} -> {arc_path}")
                    zipf.write(file_path, arc_path)
                    
    print("Archive ZIP créée avec succès dans le dossier 'distribution'.")

if __name__ == "__main__":
    create_pack()
