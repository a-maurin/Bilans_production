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

"""Résolution PNG catalogue global."""
from pathlib import Path

from core.common.carte_helper import resolve_map_png_path


def test_resolve_map_png_from_catalog(tmp_path: Path, monkeypatch) -> None:
    cartes = tmp_path / "cartes"
    cartes.mkdir()
    (cartes / "carte_global_domaines.png").write_bytes(b"x")
    monkeypatch.setattr("core.common.carte_helper.get_cartes_dir", lambda: cartes)

    profile = {
        "cartographie": {
            "catalog": [
                {"id": "global_domaines", "label": "Domaines", "fichier": "carte_global_domaines.png"},
            ]
        }
    }
    path = resolve_map_png_path("global_domaines", bilan_profiles={"global": profile})
    assert path is not None
    assert path.name == "carte_global_domaines.png"
