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
"""Validation marqueur département dans ensure_maps / generate_maps."""
from __future__ import annotations

import logging
from pathlib import Path

from core.cartographie.pochoir_helper import write_map_dept_marker
from core.common import carte_helper


def _global_profile() -> dict:
    return {
        "cartographie": {
            "catalog": [
                {"id": "global", "label": "Contrôles", "fichier": "carte_global.png"},
            ],
        }
    }


def test_ensure_maps_rejects_stale_dept21_for_dept89(monkeypatch, caplog, tmp_path) -> None:
    cartes = tmp_path / "cartes"
    cartes.mkdir()
    png = cartes / "carte_global.png"
    png.write_bytes(b"x")
    write_map_dept_marker(png, "21")

    monkeypatch.setattr(carte_helper, "get_cartes_dir", lambda: cartes)
    monkeypatch.setattr("core.chemins_projet.get_cartes_dir", lambda: cartes)
    monkeypatch.setattr(carte_helper, "qgis_available", lambda: False)
    monkeypatch.setattr(carte_helper, "generate_maps", lambda *a, **k: [])
    monkeypatch.setattr(
        "core.cartographie.pochoir_helper.warn_if_unknown_carto_dept",
        lambda _c: True,
    )

    with caplog.at_level(logging.INFO):
        out = carte_helper.ensure_maps_for_profiles(
            ["global"],
            echelle="departement",
            code="89",
            bilan_profiles={"global": _global_profile()},
        )
    assert out == []
    assert any("ignorée" in r.message and "89" in r.message for r in caplog.records)


def test_generate_maps_only_returns_valid_dept(monkeypatch, tmp_path) -> None:
    cartes = tmp_path / "cartes"
    cartes.mkdir()
    png = cartes / "carte_global.png"
    png.write_bytes(b"x")
    write_map_dept_marker(png, "89")

    monkeypatch.setattr(carte_helper, "get_cartes_dir", lambda: cartes)
    monkeypatch.setattr(carte_helper, "qgis_available", lambda: True)
    monkeypatch.setattr(carte_helper, "get_qgis_app", lambda: object())
    monkeypatch.setattr(
        "core.cartographie.production_cartographique.run_export",
        lambda *a, **k: None,
    )

    paths = carte_helper.generate_maps(
        ["global"],
        echelle="departement",
        code="89",
        bilan_profiles={"global": _global_profile()},
    )
    assert len(paths) == 1
    assert paths[0].name == "carte_global.png"


def test_generate_maps_logs_exception_traceback(monkeypatch, caplog, tmp_path) -> None:
    cartes = tmp_path / "cartes"
    cartes.mkdir()

    def _boom(*_a, **_k):
        raise RuntimeError("projet QGIS introuvable")

    monkeypatch.setattr(carte_helper, "get_cartes_dir", lambda: cartes)
    monkeypatch.setattr(carte_helper, "qgis_available", lambda: True)
    monkeypatch.setattr(carte_helper, "get_qgis_app", lambda: object())
    monkeypatch.setattr(
        "core.cartographie.production_cartographique.run_export",
        _boom,
    )
    monkeypatch.setattr(
        "core.cartographie.pochoir_helper.warn_if_unknown_carto_dept",
        lambda _c: True,
    )

    with caplog.at_level(logging.ERROR):
        out = carte_helper.generate_maps(
            ["global"],
            echelle="departement",
            code="89",
            bilan_profiles={"global": _global_profile()},
        )
    assert out == []
    assert any(r.exc_info for r in caplog.records if "in-process" in r.message)