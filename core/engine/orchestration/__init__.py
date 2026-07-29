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
MODULE : PACKAGE D'ORCHESTRATION DES PROFILS (`core/engine/orchestration/__init__.py`)
========================================================================================
Ce fichier d'initialisation sous-système ré-exporte les fonctions clés du chargeur de profils.

Exports :
  - `load_profile_config` : fonction principale de lecture et normalisation des profils YAML.
========================================================================================
"""
from core.engine.orchestration.loader import (
    load_profile_config,
    _deep_merge_dicts,
    _normalize_profile,
)

__all__ = [
    "load_profile_config",
    "_deep_merge_dicts",
    "_normalize_profile",
]
