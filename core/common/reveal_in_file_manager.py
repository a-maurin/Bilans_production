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
