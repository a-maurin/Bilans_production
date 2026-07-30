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

"""
========================================================================================
SCRIPT DE NETTOYAGE : PURGE DES TEMPORAIRES ET LOGS (`nettoyer_fichiers_obsoletes.py`)
========================================================================================
Ce script utilitaire supprime les fichiers temporaires, journaux de logs de debug, et réorganise
les scripts batch Windows (`.bat`) dans le dossier `scripts/windows/`.

Actions :
  1. Suppression des logs temporaires (`geojson_error.log`, `serveur_error.log`).
  2. Nettoyage du dossier temporaire de travail `tests/sandbox/`.
  3. Déplacement ordonné des fichiers `.bat` de lancement.
========================================================================================
"""
from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parents[1]

# Liste des fichiers/dossiers à supprimer définitivement
TO_REMOVE = [
    # Doublons, scripts scratch et logs
    ROOT / "read_pdf_scratch.py",
    ROOT / "test_read_pdf.py",
    ROOT / "tests" / "simulate_api.py",
    ROOT / "tests" / "test_simulate_api.py",
    ROOT / "core" / "cartographie" / "param" / "profils_cartes.yaml",
    ROOT / "geojson_error.log",
    ROOT / "geojson_success.log",
    ROOT / "temp" / "_compare_pdc_out.json",
    ROOT / "temp" / "geojson_error.log",
    ROOT / "temp" / "geojson_success.log",
    ROOT / "ofbilan.egg-info",
    # Logs lourds et résidus de tests/scratch/
    ROOT / "tests" / "scratch" / "api_data_debug.log",
    ROOT / "tests" / "scratch" / "serveur_error.log",
    ROOT / "tests" / "scratch" / "debug_load_pve.txt",
    ROOT / "tests" / "scratch" / "debug_pve.txt",
    # Contenu du dossier scratch d'expérimentation tests/sandbox/
    ROOT / "tests" / "sandbox",
]

# Déplacement des batchs vers scripts/windows/
BATCH_MOVES = [
    (ROOT / "core" / "cartographie" / "lancer_osgeo4w.bat", ROOT / "scripts" / "windows" / "lancer_osgeo4w.bat"),
    (ROOT / "core" / "cartographie" / "lancer_production_cartographique.bat", ROOT / "scripts" / "windows" / "lancer_production_cartographique.bat"),
]

def main():
    cleaned = 0
    # 1. Suppression des cibles
    for target in TO_REMOVE:
        if target.is_file():
            target.unlink()
            print(f"Fichier supprimé : {target.relative_to(ROOT)}")
            cleaned += 1
        elif target.is_dir():
            shutil.rmtree(target)
            print(f"Dossier supprimé : {target.relative_to(ROOT)}")
            cleaned += 1

    # 2. Déplacement des fichiers batch
    win_dir = ROOT / "scripts" / "windows"
    win_dir.mkdir(parents=True, exist_ok=True)

    for src, dst in BATCH_MOVES:
        if src.exists():
            shutil.move(str(src), str(dst))
            print(f"Fichier déplacé : {src.relative_to(ROOT)} -> {dst.relative_to(ROOT)}")
            cleaned += 1

    print(f"\nRevue et nettoyage terminés : {cleaned} action(s) effectuée(s).")

if __name__ == "__main__":
    main()
