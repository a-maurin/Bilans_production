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
"""Règles de ventilation temporelle en mode ``auto``."""

from __future__ import annotations

import pandas as pd

from core.engine import orchestrateur_profils as op
from core.engine.generation_pdf_profil import resolve_ventilation_mode_global
from core.engine.ventilation_temporelle import (
    VENTILATION_JOURS_DEUX_ANS,
    VENTILATION_JOURS_SIX_MOIS,
    VENTILATION_JOURS_UN_AN,
    resolve_ventilation_auto,
)


def _profile_auto(seuil: int = 366) -> dict:
    return {"periode_analyse": {"ventilation": {"type": "auto", "seuil_jours": seuil}}}


def test_hebdomadaire_si_duree_inferieure_ou_egale_a_six_mois() -> None:
    d1 = pd.Timestamp("2025-01-01")
    d2 = pd.Timestamp("2025-07-01")
    assert (d2 - d1).days <= VENTILATION_JOURS_SIX_MOIS
    mode, *_ = op._resolve_ventilation_mode_from_profile(_profile_auto(), date_deb_ts=d1, date_fin_ts=d2)
    assert mode == "hebdomadaire"
    assert resolve_ventilation_mode_global(d1, d2) == "hebdomadaire"


def test_mensuelle_si_duree_superieure_six_mois_et_inferieure_ou_egale_un_an() -> None:
    d1 = pd.Timestamp("2025-01-01")
    d2 = pd.Timestamp("2025-12-31")
    assert VENTILATION_JOURS_SIX_MOIS < (d2 - d1).days <= VENTILATION_JOURS_UN_AN
    mode, *_ = op._resolve_ventilation_mode_from_profile(_profile_auto(), date_deb_ts=d1, date_fin_ts=d2)
    assert mode == "mensuelle"
    assert resolve_ventilation_mode_global(d1, d2) == "mensuelle"


def test_mensuelle_a_exactement_un_an() -> None:
    """Année civile 2024 → 2025 (366 j, année bissextile) : mensuelle."""
    d1 = pd.Timestamp("2024-01-01")
    d2 = pd.Timestamp("2025-01-01")
    assert (d2 - d1).days == VENTILATION_JOURS_UN_AN
    assert resolve_ventilation_auto((d2 - d1).days) == "mensuelle"


def test_hebdomadaire_a_six_mois() -> None:
    d1 = pd.Timestamp("2025-01-01")
    d2 = pd.Timestamp("2025-07-01")
    assert (d2 - d1).days <= VENTILATION_JOURS_SIX_MOIS
    assert resolve_ventilation_auto((d2 - d1).days) == "hebdomadaire"


def test_mensuelle_entre_un_an_et_deux_ans() -> None:
    p = _profile_auto()
    d1 = pd.Timestamp("2024-01-01")
    d2 = pd.Timestamp("2025-11-30")
    mode, *_ = op._resolve_ventilation_mode_from_profile(p, date_deb_ts=d1, date_fin_ts=d2)
    assert mode == "mensuelle"
    assert resolve_ventilation_mode_global(d1, d2) == "mensuelle"


def test_trimestrielle_a_exactement_deux_ans() -> None:
    d1 = pd.Timestamp("2018-01-01")
    d2 = pd.Timestamp("2020-01-01")
    assert (d2 - d1).days == VENTILATION_JOURS_DEUX_ANS
    assert resolve_ventilation_auto((d2 - d1).days, seuil_jours=366) == "trimestrielle"


def test_auto_trimestrielle_si_duree_entre_deux_ans_et_seuil() -> None:
    """Avec un seuil > 2 ans, une période ≥ 2 ans mais ≤ seuil reste trimestrielle."""
    p = {"periode_analyse": {"ventilation": {"type": "auto", "seuil_jours": 800}}}
    d1 = pd.Timestamp("2020-01-01")
    d2 = pd.Timestamp("2022-01-10")
    mode, *_ = op._resolve_ventilation_mode_from_profile(p, date_deb_ts=d1, date_fin_ts=d2)
    assert mode == "trimestrielle"


def test_auto_annuelle_si_duree_au_dela_du_seuil() -> None:
    p = {"periode_analyse": {"ventilation": {"type": "auto", "seuil_jours": 800}}}
    d1 = pd.Timestamp("2010-01-01")
    d2 = pd.Timestamp("2025-01-01")
    mode, *_ = op._resolve_ventilation_mode_from_profile(p, date_deb_ts=d1, date_fin_ts=d2)
    assert mode == "annuelle"


def test_ventilation_forcee_mensuelle() -> None:
    p = {"periode_analyse": {"ventilation": {"type": "mensuelle"}}}
    d1 = pd.Timestamp("2025-01-01")
    d2 = pd.Timestamp("2025-12-31")
    mode, vent_type, *_ = op._resolve_ventilation_mode_from_profile(p, date_deb_ts=d1, date_fin_ts=d2)
    assert mode == "mensuelle"
    assert str(vent_type).lower() == "mensuelle"