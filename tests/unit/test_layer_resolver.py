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
"""Tests unitaires — résolution dynamique des couches QGIS."""
from core.cartographie.layer_resolver import (
    candidates_for_role,
    infer_layer_role,
    resolve_layer_name,
    should_apply_yaml_symbology,
)

PROJECT_LAYERS = [
    "ESRI Topo",
    "point_ctrl_20251231_wgs84",
    "point_ctrl_20260505_wgs84",
    "point_ctrl_20260505_wgs84 copie",
    "localisation_infrac_FAITS_20260306",
    "localisation_infrac_FAITS_20260505",
    "pochoir_sd21",
    "pochoir_sd21 copie",
    "Zone d'interdiction d'agrainage 2026",
    "Zone infectée",
    "Zone Infectee 2026",
]


def test_resolve_point_controles_latest_snapshot():
    name, source = resolve_layer_name(
        configured_name="point_ctrl_20260205_wgs84",
        layer_role="point_controles",
        available_names=PROJECT_LAYERS,
    )
    assert name == "point_ctrl_20260505_wgs84"
    assert source == "role"


def test_resolve_pej_latest_snapshot():
    name, source = resolve_layer_name(
        configured_name="localisation_infrac_FAITS_20260205",
        layer_role="pej",
        available_names=PROJECT_LAYERS,
    )
    assert name == "localisation_infrac_FAITS_20260505"
    assert source == "role"


def test_resolve_exact_name_preferred():
    name, source = resolve_layer_name(
        configured_name="pochoir_sd21",
        layer_role="pochoir",
        available_names=PROJECT_LAYERS,
        dept_code="21",
    )
    assert name == "pochoir_sd21"
    assert source == "dept"


def test_infer_role_from_layer_key():
    assert infer_layer_role("point_ctrl_20260205_wgs84", "") == "point_controles"
    assert infer_layer_role("localisation_infrac_FAITS_20260205", "") == "pej"


def test_candidates_exclude_copie():
    candidates = candidates_for_role("point_controles", PROJECT_LAYERS)
    assert "point_ctrl_20260505_wgs84 copie" not in candidates
    assert candidates[-1] == "point_ctrl_20260505_wgs84"


def test_symbology_source_default_qgis():
    assert should_apply_yaml_symbology(None, "qgis", "qgis") is False


def test_symbology_source_yaml_override():
    assert should_apply_yaml_symbology(None, "yaml", "qgis") is True
    assert should_apply_yaml_symbology("yaml", "qgis", "qgis") is True
    assert should_apply_yaml_symbology("qgis", "yaml", "qgis") is False


def test_missing_layer():
    name, source = resolve_layer_name(
        configured_name="couche_inexistante",
        layer_role=None,
        available_names=PROJECT_LAYERS,
    )
    assert name is None
    assert source == "missing"