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
"""Tests — avertissements cartes lorsque QGIS est absent."""
from __future__ import annotations

import logging

from core.common import carte_helper


def test_generate_maps_warns_when_qgis_unavailable(monkeypatch, caplog) -> None:
    monkeypatch.setattr(carte_helper, "qgis_available", lambda: False)
    monkeypatch.setattr(
        "core.cartographie.qgis_runtime.run_cartography_export_subprocess",
        lambda *a, **k: False,
    )
    with caplog.at_level(logging.WARNING):
        result = carte_helper.generate_maps(
            ["global"],
            echelle="departement",
            code="25",
        )
    assert result == []
    assert any("échouée" in r.message or "sous-processus" in r.message for r in caplog.records)


def test_ensure_maps_warns_unresolved_without_qgis(monkeypatch, caplog, tmp_path) -> None:
    cartes = tmp_path / "cartes"
    cartes.mkdir()
    monkeypatch.setattr(carte_helper, "get_cartes_dir", lambda: cartes)
    monkeypatch.setattr(carte_helper, "qgis_available", lambda: False)
    monkeypatch.setattr(carte_helper, "find_map", lambda _pid: None)

    with caplog.at_level(logging.WARNING):
        out = carte_helper.ensure_maps_for_profiles(
            ["global"],
            echelle="departement",
            code="25",
        )
    assert out == []
    assert any("Cartes absentes" in r.message or "non valides" in r.message for r in caplog.records)