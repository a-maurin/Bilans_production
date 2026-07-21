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
"""PA dérivées des contrôles à manquement et agrégations procédures."""

from __future__ import annotations

import numpy as np
import pandas as pd

from core.common.utilitaires_metier import (
    agg_procedures_dossiers_par_domaine,
    agg_procedures_par_type_usager_domaine,
    build_tab_resultats,
    count_procedures_liees_controle_sur_points,
    count_pa_induites_par_controles,
    is_filled_procedure_code,
    points_as_pa_lignes,
    resultat_induit_pa,
)


def test_is_filled_procedure_code_ignore_nan() -> None:
    assert not is_filled_procedure_code(np.nan)
    assert not is_filled_procedure_code(None)
    assert not is_filled_procedure_code("")
    assert is_filled_procedure_code("SD21-2025-PA-0011")



def test_build_tab_resultats_inclut_en_attente() -> None:
    point = pd.DataFrame(
        {
            "resultat": ["Conforme", "Infraction", np.nan],
        }
    )
    tab = build_tab_resultats(point)
    assert int(tab["nb"].sum()) == 3
    assert int(tab.loc[tab["resultat"] == "En attente", "nb"].iloc[0]) == 1
    assert abs(float(tab["taux"].sum()) - 1.0) < 1e-9


def test_agg_procedures_pa_depuis_code_pa() -> None:
    df = pd.DataFrame(
        {
            "type_usager": ["Agriculteur 1", "Agriculteur 1", "Agriculteur 1"],
            "domaine": ["Dom A", "Dom A", "Dom B"],
            "code_pej": [np.nan, np.nan, "PEJ-001"],
            "code_pa": [np.nan, "PA-001", np.nan],
        }
    )
    out = agg_procedures_par_type_usager_domaine(df)
    assert int(out["nb_pej"].sum()) == 1
    assert int(out["nb_pa"].sum()) == 1
    dom_a = out[out["domaine"] == "Dom A"].iloc[0]
    dom_b = out[out["domaine"] == "Dom B"].iloc[0]
    assert int(dom_a["nb_pej"]) == 0
    assert int(dom_a["nb_pa"]) == 1
    assert int(dom_b["nb_pej"]) == 1
    assert int(dom_b["nb_pa"]) == 0


def test_count_pa_induites_et_lignes_synthetiques() -> None:
    point = pd.DataFrame(
        {
            "date_ctrl": pd.to_datetime(["2025-01-01", "2025-02-01", "2025-03-01"]),
            "domaine": ["Dom A", "Dom A", "Dom B"],
            "theme": ["Th A", "Th B", "Th C"],
            "code_pa": [np.nan, "PA-1", np.nan],
            "code_pej": ["PEJ-1", np.nan, np.nan],
            "dc_id": ["dc1", "dc2", "dc3"],
        }
    )
    pej = pd.DataFrame({"DOMAINE": ["Dom A", "Dom B"], "type_usager": ["A", "A"]})
    assert count_pa_induites_par_controles(point) == 1
    nb_pej, nb_pa = count_procedures_liees_controle_sur_points(point)
    assert nb_pej == 1
    assert nb_pa == 1
    pa_lignes = points_as_pa_lignes(point)
    assert len(pa_lignes) == 1
    assert pa_lignes.iloc[0]["DOMAINE"] == "Dom A"
    dossiers = agg_procedures_dossiers_par_domaine(pej, pa_lignes)
    assert int(dossiers["nb_pej"].sum()) == 2
    assert int(dossiers["nb_pa"].sum()) == 1
