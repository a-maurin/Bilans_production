"""
DEPRECATED — Ce script est remplacé par le moteur thématique unifié (bilan_thematique_engine.py).

Tous les profils par mots-clés passent désormais par le moteur unifié.
Utiliser : python scripts/run_bilan.py --mode thematique --profil <id>
"""
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))


def run_generic_bilan(
    profil_id: str,
    date_deb: str,
    date_fin: str,
    dept_code: str,
) -> int:
    """Redirige vers le moteur unifié."""
    from scripts.bilan_thematique.bilan_thematique_engine import run_engine
    return run_engine(profil_id, date_deb, date_fin, dept_code)
