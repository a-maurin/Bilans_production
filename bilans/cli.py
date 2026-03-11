"""
Interface en ligne de commande principale pour le package `bilans`.

Cette CLI délègue pour l'instant à `scripts.run_bilan.main` afin de rester
compatible avec l'architecture existante, tout en offrant un point d'entrée
standard (`python -m bilans` ou script console).
"""

from __future__ import annotations

import sys

from scripts.run_bilan import main as _run_bilan_main


def main() -> int:
    """Point d'entrée CLI : délègue à `scripts.run_bilan.main`."""
    return _run_bilan_main()


if __name__ == "__main__":
    sys.exit(main())

