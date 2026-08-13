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

"""
========================================================================================
MODULE : VENTILATION TEMPORELLE AUTOMATIQUE (`ventilation_temporelle.py`)
========================================================================================
Ce module détermine la maille temporelle d'analyse optimale (hebdomadaire, mensuelle,
trimestrielle ou annuelle) en fonction de la durée de la période sélectionnée.

Règles appliquées :
  - Périodes ≤ 6 mois (183 jours) : découpage hebdomadaire.
  - Périodes entre 6 mois et 2 ans : découpage mensuel.
  - Périodes de 2 ans : découpage trimestriel.
  - Périodes > 2 ans : découpage annuel.
========================================================================================
"""
from __future__ import annotations

# Bornes en jours (durée inclusive date_fin − date_deb).
VENTILATION_JOURS_SIX_MOIS = 183  # ~6 mois civil
VENTILATION_JOURS_UN_AN = 366
VENTILATION_JOURS_DEUX_ANS = 730


def resolve_ventilation_auto(duree_jours: int, *, seuil_jours: int = 366) -> str:
    """Ventilation en mode ``auto`` selon la durée de la période d'analyse (en jours).

    - < 183 j (< 6 mois) : hebdomadaire (libellé de période ``YYYY-Sww``)
    - 183 j à 366 j (6 mois à 1 an) : mensuelle (libellé ``YYYY-MM``)
    - 367 j à 730 j (1 an à 2 ans) : trimestrielle
    - > 730 j (> 2 ans) : annuelle
    """
    if duree_jours < VENTILATION_JOURS_SIX_MOIS:
        return "hebdomadaire"
    if duree_jours <= VENTILATION_JOURS_UN_AN:
        return "mensuelle"
    if duree_jours <= VENTILATION_JOURS_DEUX_ANS:
        return "trimestrielle"
    return "annuelle"