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
"""Régression : PA dérivées des contrôles à manquement (bilan global)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from core.engine.agregations_profil import analyse_pej_pa_global


def test_pa_par_domaine_depuis_controles_manquement(tmp_path: Path) -> None:
    point = pd.DataFrame(
        {
            "dc_id": ["dc1", "dc2", "dc3"],
            "date_ctrl": pd.to_datetime(["2025-01-01", "2025-02-01", "2025-03-01"]),
            "resultat": ["Manquement", "Manquement et infraction", "Conforme"],
            "code_pa": ["PA1", "PA2", None],
            "domaine": ["Eau", "Faune", "Flore"],
            "theme": ["Th1", "Th2", "Th3"],
        }
    )
    pej = pd.DataFrame(
        {
            "ENTITE_ORIGINE_PROCEDURE": ["SD21", "SD21"],
            "DC_ID": ["p1", "p2"],
            "DOMAINE": ["A", "B"],
        }
    )
    pa_ods = pd.DataFrame({"DC_ID": ["dc1"]})
    analyse_pej_pa_global(
        tmp_path,
        point,
        pa_ods,
        pej,
        tmp_path,
        echelle="departement",
        code="21",
    )
    pa_dom = pd.read_csv(tmp_path / "pa_global_par_domaine.csv", sep=";")
    assert len(pa_dom) == 2
    assert int(pa_dom["nb_pa"].sum()) == 2
    resume = pd.read_csv(tmp_path / "pa_global_resume.csv", sep=";")
    assert int(resume.iloc[0]["nb_pa_global"]) == 2