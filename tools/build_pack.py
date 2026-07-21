# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, mais SANS AUCUNE GARANTIE ;
# sans même la garantie implicite de QUALITÉ MARCHANDE ou D'ADÉQUATION À UN USAGE PARTICULIER.
# Voir la Licence Publique Générale GNU pour plus de détails.
#
# CONDITIONS SUPPLÉMENTAIRES D'ATTRIBUTION (SECTION 7(b) DE LA GPL v3) :
# Conformément à la section 7(b) de la GNU GPL v3, vous devez expressément conserver
# intactes et lisibles toutes les mentions d'auteur, notices de copyright et la présente
# clause dans chaque fichier source ou interface utilisateur redistribué. Toute version modifiée
# doit clairement indiquer qu'elle a été altérée et ne doit en aucun cas supprimer le nom
# de l'auteur original (Aguirre MAURIN).

#
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