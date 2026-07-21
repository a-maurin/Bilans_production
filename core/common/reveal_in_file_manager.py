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
"""Ouverture d'un chemin (fichier/dossier) via l'application système par défaut."""
from __future__ import annotations

import logging
import os
import subprocess
import sys
from pathlib import Path

logger = logging.getLogger("ofbilan.reveal")


def reveal_path_in_file_manager(path: Path) -> None:
    """
    Ouvre ``path`` avec l'application système par défaut.

    Ne fait rien sous CI, si ``BILANS_OPEN_OUTPUT_DIR`` vaut 0/false/no/off,
    ou en cas d'échec (log warning, pas d'exception propagée).
    """
    if os.environ.get("CI"):
        return
    flag = os.environ.get("BILANS_OPEN_OUTPUT_DIR", "").strip().lower()
    if flag in ("0", "false", "no", "off"):
        return
    try:
        resolved = path.resolve()
    except OSError as exc:
        logger.warning("Impossible de résoudre le chemin %s : %s", path, exc)
        return
    if not resolved.exists():
        logger.warning("Ouverture ignorée (chemin inexistant) : %s", resolved)
        return
    try:
        if sys.platform == "win32":
            os.startfile(resolved)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(resolved)], check=False)
        else:
            subprocess.run(["xdg-open", str(resolved)], check=False)
    except Exception as exc:
        logger.warning("Impossible d'ouvrir le chemin %s : %s", resolved, exc)