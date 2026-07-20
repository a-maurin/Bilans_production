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

"""Résolution de plusieurs cartes PNG pour la section 5 du PDF."""

from __future__ import annotations

from pathlib import Path

from core.common.carte_helper import (
    expected_map_filenames,
    resolve_map_layout,
    resolve_profile_map_paths,
)


def test_resolve_deux_cartes_par_convention(tmp_path: Path, monkeypatch) -> None:
    cartes = tmp_path / "cartes"
    cartes.mkdir()
    (cartes / "carte_demo.png").write_bytes(b"x")
    (cartes / "carte_demo_2.png").write_bytes(b"y")

    monkeypatch.setattr("core.common.carte_helper.get_cartes_dir", lambda: cartes)

    paths = resolve_profile_map_paths("demo")
    assert len(paths) == 2
    assert paths[0].name == "carte_demo.png"
    assert paths[1].name == "carte_demo_2.png"


def test_resolve_une_seule_carte_si_pas_de_seconde(tmp_path: Path, monkeypatch) -> None:
    cartes = tmp_path / "cartes"
    cartes.mkdir()
    (cartes / "carte_demo.png").write_bytes(b"x")

    monkeypatch.setattr("core.common.carte_helper.get_cartes_dir", lambda: cartes)

    paths = resolve_profile_map_paths("demo")
    assert len(paths) == 1


def test_profils_yaml_cartographie_prioritaire(tmp_path: Path, monkeypatch) -> None:
    cartes = tmp_path / "cartes"
    cartes.mkdir()
    (cartes / "carte_x_contexte.png").write_bytes(b"a")
    (cartes / "carte_x_synthese.png").write_bytes(b"b")

    monkeypatch.setattr("core.common.carte_helper.get_cartes_dir", lambda: cartes)

    profile = {
        "cartographie": {
            "fichiers": ["carte_{map_id}_contexte.png", "carte_{map_id}_synthese.png"],
            "disposition": "verticale",
        }
    }
    paths = resolve_profile_map_paths("x", profile=profile)
    assert [p.name for p in paths] == ["carte_x_contexte.png", "carte_x_synthese.png"]
    assert resolve_map_layout(profile=profile) == "vertical"


def test_expected_map_filenames() -> None:
    names = expected_map_filenames("AGR")
    assert names == ["carte_AGR.png", "carte_AGR_2.png"]


def test_disposition_par_defaut_verticale() -> None:
    assert resolve_map_layout() == "vertical"
