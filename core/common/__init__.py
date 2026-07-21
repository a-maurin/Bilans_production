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
"""Sous-package `bilans.common` exposant les utilitaires partagés."""

from core.common.bilan_config import *  # noqa: F401,F403
from core.common.chargeurs_donnees import *  # noqa: F401,F403
from core.common.utilitaires_metier import *  # noqa: F401,F403
from core.common.ofb_charte import *  # noqa: F401,F403
from core.common.pdf_report_builder import *  # noqa: F401,F403
from core.common.pdf_utils import *  # noqa: F401,F403
from core.common.rendus_graphiques import *  # noqa: F401,F403
from core.common.carte_helper import *  # noqa: F401,F403
from core.common.prompt_periode import *  # noqa: F401,F403
