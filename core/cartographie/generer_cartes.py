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
"""
Wrapper Python pour la génération de cartes.

Pour l'instant, ce script délègue au lanceur existant basé sur QGIS
(`src/ofbilan/cartographie/lancer_production_cartographique.bat`),
ce qui permet d'avoir un point d'entrée CLI stable :

    python src/ofbilan/cartographie/generer_cartes.py --profil agrainage --date-deb 2025-01-01 --date-fin 2025-12-31 --dept-code 21
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Génération de cartes (wrapper QGIS).")
    parser.add_argument(
        "--profil",
        type=str,
        default="tous",
        help="Profil(s) carte : agrainage, global, global_usagers, ... ou liste séparée par des virgules.",
    )
    parser.add_argument("--date-deb", type=str, required=True, help="Date début (YYYY-MM-DD).")
    parser.add_argument("--date-fin", type=str, required=True, help="Date fin (YYYY-MM-DD).")
    parser.add_argument("--dept-code", type=str, default="21", help="Code département (ex. 21).")
    args = parser.parse_args()

    launcher = Path(__file__).resolve().parent / "lancer_production_cartographique.bat"
    if not launcher.exists():
        print(f"Erreur : lanceur QGIS introuvable : {launcher}", file=sys.stderr)
        return 1

    cmd = [
        str(launcher),
        args.profil,
        "--date-deb",
        args.date_deb,
        "--date-fin",
        args.date_fin,
        "--dept-code",
        args.dept_code,
    ]

    try:
        return subprocess.call(cmd)
    except OSError as e:
        print(f"Erreur lors de l'appel au lanceur QGIS : {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
