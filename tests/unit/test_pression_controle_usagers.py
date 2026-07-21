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
"""Section 4 pression de contrôle : effectifs multi-usagers par catégorie."""
from __future__ import annotations

import pandas as pd

from core.common.utilitaires_metier import (
    agg_controles_par_type_usager_domaine,
    agg_controles_par_type_usager_theme,
    agg_effectifs_usagers,
    agg_nb_localisations_par_type_usager,
    count_multi_usager_controles,
)


def test_agg_effectifs_repartit_multi_usagers():
    df = pd.DataFrame(
        {
            "type_usager": [
                "Agriculteur 3, Particulier 1",
                "Particulier 2",
            ],
        }
    )
    out = agg_effectifs_usagers(df)
    assert int(out["nb"].sum()) == 6


def test_agg_effectifs_vide():
    out = agg_effectifs_usagers(pd.DataFrame())
    assert list(out.columns) == ["type_usager", "nb", "nb_operations"]
    assert out.empty


def test_agg_effectifs_ne_compte_qu_une_fois_par_fc_id():
    df = pd.DataFrame(
        {
            "fc_id": ["FC-1", "FC-1"],
            "type_usager": ["Collectivité 12", "Collectivité 12"],
        }
    )

    out = agg_effectifs_usagers(df)

    assert int(out["nb"].sum()) == 12
    assert int(out.loc[out["type_usager"] == "Collectivité", "nb"].sum()) == 12


def test_agg_nb_localisations_consolide_par_fc_id():
    df = pd.DataFrame(
        {
            "fc_id": ["FC-1", "FC-1"],
            "type_usager": ["Collectivité 12", "Collectivité 12"],
        }
    )

    out = agg_nb_localisations_par_type_usager(df)

    assert int(out["nb"].sum()) == 1
    assert int(out.loc[out["type_usager"] == "Collectivité", "nb"].sum()) == 1


def test_agg_controles_par_type_usager_dimension_consolide_par_fc_id():
    df = pd.DataFrame(
        {
            "fc_id": ["FC-1", "FC-1"],
            "domaine": ["Eau", "Eau"],
            "theme": ["Thème A", "Thème A"],
            "type_usager": ["Collectivité 12", "Collectivité 12"],
        }
    )

    out_dom = agg_controles_par_type_usager_domaine(df)
    out_theme = agg_controles_par_type_usager_theme(df)

    assert "nb_localisations" in out_dom.columns
    assert int(out_dom["nb_localisations"].sum()) == 1
    assert int(
        out_dom.loc[
            (out_dom["type_usager"] == "Collectivité") & (out_dom["domaine"] == "Eau"),
            "nb_localisations",
        ].sum()
    ) == 1
    assert int(out_theme["nb_localisations"].sum()) == 1
    assert int(
        out_theme.loc[
            (out_theme["type_usager"] == "Collectivité") & (out_theme["theme"] == "Thème A"),
            "nb_localisations",
        ].sum()
    ) == 1


def test_count_multi_usager_controles_consolide_par_fc_id():
    df = pd.DataFrame(
        {
            "fc_id": ["FC-1", "FC-1", "FC-2"],
            "type_usager": [
                "Collectivité 1, Agriculteur 1",
                "Collectivité 1, Agriculteur 1",
                "Collectivité 1",
            ],
        }
    )

    assert count_multi_usager_controles(df) == 1