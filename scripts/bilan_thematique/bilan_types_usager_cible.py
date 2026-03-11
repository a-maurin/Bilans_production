"""
DEPRECATED — Ce script est remplacé par le moteur thématique unifié.

Utiliser : python scripts/run_bilan.py --mode thematique --profil types_usager_cible
"""
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))


def run_bilan(date_deb: str, date_fin: str, dept_code: str) -> int:
    from scripts.bilan_thematique.bilan_thematique_engine import run_engine
    return run_engine("types_usager_cible", date_deb, date_fin, dept_code)
