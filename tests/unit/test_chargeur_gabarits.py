# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, me SANS AUCUNE GARANTIE ;
# sans même la garantie implicite de QUALITÉ MARCHANDE ou D'ADÉQUATION À UN USAGE PARTICULIER.
# Voir la Licence Publique Générale GNU pour plus de détails.

from pathlib import Path
from core.common.chargeur_gabarits import (
    list_gabarits,
    load_gabarit,
    is_gabarit_compatible,
    resolve_gabarit_for_service,
)


def test_list_gabarits_finds_srp_r27():
    gabarits = list_gabarits()
    ids = [g["gabarit_id"] for g in gabarits]
    assert "srp_r27" in ids

    srp = next(g for g in gabarits if g["gabarit_id"] == "srp_r27")
    assert "Service Régional Police" in srp["label"]
    assert srp["organisation"]["code_region"] == "r27"
    assert srp["organisation"]["service"] == "srp"


def test_load_gabarit_valid_and_missing():
    g_srp = load_gabarit("srp_r27")
    assert g_srp is not None
    assert g_srp["gabarit_id"] == "srp_r27"
    assert g_srp["layout"] == "brochure_custom"

    # Inexistant -> None avec log warning gracieux
    g_unknown = load_gabarit("gabarit_inexistant_xyz")
    assert g_unknown is None


def test_is_gabarit_compatible():
    g = {
        "gabarit_id": "test_g",
        "cible": "bilan",
        "profils_compatibles": ["chasse", "global"],
    }
    assert is_gabarit_compatible(g, profile_id="chasse", cible="bilan") is True
    assert is_gabarit_compatible(g, profile_id="agrainage", cible="bilan") is False
    assert is_gabarit_compatible(g, profile_id="chasse", cible="brochure") is False


def test_resolve_gabarit_for_service():
    resolved = resolve_gabarit_for_service(code_region="r27", code_service="srp")
    assert resolved == "srp_r27"

    resolved_none = resolve_gabarit_for_service(code_region="r99", code_service="inconnu")
    assert resolved_none is None


def test_list_gabarits_finds_brochure_defaut():
    gabarits = list_gabarits()
    ids = [g["gabarit_id"] for g in gabarits]
    assert "brochure_defaut" in ids

    g_defaut = load_gabarit("brochure_defaut")
    assert g_defaut is not None
    assert g_defaut["gabarit_id"] == "brochure_defaut"
    assert g_defaut["layout"] == "brochure"

