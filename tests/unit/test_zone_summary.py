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
"""Non-régression : décompte zone « Département » = hors TUB et hors PNF."""

from __future__ import annotations

import pandas as pd

from core.common.utilitaires_metier import (
    ZONE_KEY_DEPARTEMENT,
    ZONE_LABEL_DEPARTEMENT_HORS,
    _zone_count,
    _zone_summary,
    zone_table_display_label,
)


def test_zone_summary_departement_counts_hors_tub_pnf() -> None:
    df = pd.DataFrame(
        {
            "insee": ["01001", "01002", "01003", "01004"],
            "resultat": ["Conforme", "Non conforme", "Conforme", "Conforme"],
        }
    )
    tub = {"01001", "01002"}
    pnf = {"01002", "01003"}
    summary = _zone_summary(df, "insee", tub, pnf)
    dep = summary.loc[summary["zone"] == ZONE_KEY_DEPARTEMENT].iloc[0]
    assert int(dep["nb_total"]) == 1
    assert int(dep["nb_non_conforme"]) == 0


def test_zone_count_departement_hors_tub_pnf() -> None:
    df = pd.DataFrame({"insee": ["01001", "01002", "01003"]})
    tub = {"01001"}
    pnf = {"01002"}
    counts = _zone_count(df, "insee", tub, pnf)
    dep_nb = int(counts.loc[counts["zone"] == ZONE_KEY_DEPARTEMENT, "nb"].iloc[0])
    assert dep_nb == 1


def test_zone_table_display_label() -> None:
    assert zone_table_display_label("Département") == ZONE_LABEL_DEPARTEMENT_HORS
    assert zone_table_display_label("Zone TUB") == "Zone TUB"