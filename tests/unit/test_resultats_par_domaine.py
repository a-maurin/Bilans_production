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
from __future__ import annotations

import pandas as pd

from core.engine.agregations_profil import _resultats_par_domaine_pour_pdf


def test_resultats_par_domaine_conforme_nc_attente() -> None:
    pt = pd.DataFrame(
        {
            "domaine": ["A", "A", "B", "B", "B", "C"],
            "resultat": [
                "Conforme",
                "Infraction",
                "Conforme",
                "Manquement",
                "Conforme",
                "Autre",
            ],
        }
    )
    out = _resultats_par_domaine_pour_pdf(pt)
    row_b = out.loc[out["domaine"] == "B"].iloc[0]
    assert int(row_b["Conforme"]) == 2
    assert int(row_b["Non-conforme"]) == 1
    assert int(row_b["En attente"]) == 0

    row_c = out.loc[out["domaine"] == "C"].iloc[0]
    assert int(row_c["Conforme"]) == 0
    assert int(row_c["Non-conforme"]) == 0
    assert int(row_c["En attente"]) == 1