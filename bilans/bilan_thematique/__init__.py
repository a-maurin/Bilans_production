"""
Sous-package `bilans.bilan_thematique`.

À ce stade, le moteur thématique reste dans `scripts.bilan_thematique`. Ce module
ré-exporte les fonctions publiques pour stabiliser les imports.
"""

from scripts.bilan_thematique.run_bilan_thematique import (  # noqa: F401
    run_thematic,
    _list_profiles,
    _resolve_profils,
)
from scripts.bilan_thematique.bilan_thematique_engine import *  # noqa: F401,F403

