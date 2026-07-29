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
MODULE : POINT D'ENTRÉE DU MOTEUR D'ANALYSIS (`core/engine/__init__.py`)
========================================================================================
Ce fichier d'initialisation expose l'API publique du moteur de génération d'OFBilan.

Fonctions et classes exportées :
  - `list_profiles()` : liste des profils de bilan actifs.
  - `resolve_profile_ids()` : résolution des profils par identifiant ou index.
  - `run_profile()` / `run_profiles_batch()` : exécution simple ou par lot des bilans.
  - `SectionRegistry` : registre central des fonctions de rendu des sections.
========================================================================================
"""

from core.engine.catalogue_profils import list_profiles, resolve_profile_ids
from core.engine.registre_sections_pdf import SectionRegistry
from core.engine.execution_lots_profils import run_profile, run_profiles_batch

__all__ = [
    "list_profiles",
    "resolve_profile_ids",
    "run_profile",
    "run_profiles_batch",
    "SectionRegistry",
]