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
"""Tests — mode layout-driven (découverte couches depuis layout / arborescence QGIS)."""
from pathlib import Path

import pytest

from core.cartographie.config_cartes_model import LayerSymbologyConfig, ProfileConfig
from core.cartographie.layout_layers import (
    build_layer_configs_from_names,
    discover_layers_from_qgs_file,
    filter_operational_layer_names,
    infer_filter_type_for_layer,
    is_basemap_layer,
    is_operational_layer,
    parse_qgs_layer_tree_group,
    parse_qgs_layout_layerset,
)

QGZ = Path("ref/programme/sig/bilans_carte.qgz")


@pytest.fixture(scope="module")
def qgs_text() -> str:
    if not QGZ.exists():
        pytest.skip("Projet export QGZ absent")
    import zipfile

    with zipfile.ZipFile(QGZ) as zf:
        return zf.read("bilans_carte.qgs").decode("utf-8", "replace")


def test_is_basemap():
    assert is_basemap_layer("ESRI Topo") is True
    assert is_operational_layer("point_ctrl_20260505_wgs84") is True
    assert is_operational_layer("ESRI Topo") is False
    assert is_operational_layer("point_ctrl_20260505_wgs84 copie") is False


def test_filter_operational_excludes_basemaps():
    names = ["ESRI Topo", "point_ctrl_20260505_wgs84", "emprise_dep"]
    assert filter_operational_layer_names(names) == ["point_ctrl_20260505_wgs84", "emprise_dep"]


def test_infer_filter_type_agrainage():
    assert infer_filter_type_for_layer("point_ctrl_20260505_wgs84", "agrainage") == "point_ctrl_agrainage"
    assert infer_filter_type_for_layer("point_ctrl_20260505_wgs84", "global") == "point_ctrl_global"
    assert infer_filter_type_for_layer("localisation_infrac_FAITS_20260505", "global") == "pj"


def test_parse_layer_tree_group_controles(qgs_text):
    layers = parse_qgs_layer_tree_group(qgs_text, "Contrôles")
    assert any("point_ctrl" in n for n in layers)
    assert all(not is_basemap_layer(n) for n in layers)


def test_layout_layerset_empty_for_current_project(qgs_text):
    layers = parse_qgs_layout_layerset(
        qgs_text,
        "Bilan 2025 / 2026 - Agrainage illicite - Côte d'Or",
    )
    assert layers == []


def test_discover_from_qgs_with_group(qgs_text):
    layers = discover_layers_from_qgs_file(
        qgs_text,
        "Bilan – Chasse – SD21",
        layout_layer_group="Contrôles",
    )
    assert any("point_ctrl" in n for n in layers)


def test_build_layer_configs_from_discovered_names():
    prof = ProfileConfig(
        id="agrainage",
        title="test",
        layout_name="layout",
        output_filename="carte.png",
        layers={
            "point_controles": LayerSymbologyConfig(
                layer_name="point_ctrl_old",
                layer_role="point_controles",
                filter_type="point_ctrl_agrainage",
            ),
        },
    )
    configs = build_layer_configs_from_names(
        ["point_ctrl_20260505_wgs84", "emprise_dep", "ESRI Topo"],
        prof,
        prof.layers,
    )
    assert "point_ctrl_20260505_wgs84" in configs
    assert configs["point_ctrl_20260505_wgs84"].filter_type == "point_ctrl_agrainage"
    assert "ESRI Topo" not in configs


def test_profile_config_layers_from_layout_default():
    prof = ProfileConfig(
        id="x",
        title="t",
        layout_name="l",
        output_filename="c.png",
        layers_from_layout=True,
    )
    assert prof.layers_from_layout is True