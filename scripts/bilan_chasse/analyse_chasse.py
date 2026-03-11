"""
DEPRECATED — Ce script est remplacé par le moteur thématique unifié.

Utiliser à la place :
  python scripts/run_bilan.py --mode thematique --profil chasse
  python scripts/bilan_thematique/run_bilan_thematique.py --profil chasse

L'ancienne version est archivée dans old/scripts/bilan_chasse/.
"""
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))


def run_bilan(date_deb: str, date_fin: str, dept_code: str) -> int:
    """Redirige vers le moteur unifié."""
    from scripts.bilan_thematique.bilan_thematique_engine import run_engine
    return run_engine("chasse", date_deb, date_fin, dept_code)


def main() -> None:
    print("ATTENTION : ce script est déprécié. Utilisez :")
    print("  python scripts/run_bilan.py --mode thematique --profil chasse")
    from scripts.common.prompt_periode import ask_periode_dept
    date_deb, date_fin, dept_code = ask_periode_dept()
    sys.exit(run_bilan(date_deb, date_fin, dept_code))


if __name__ == "__main__":
    main()
