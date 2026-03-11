from __future__ import annotations

"""
Outil de comparaison simple entre l'ancien code (legacy) et le nouveau moteur.

Usage indicatif :

    python tools/compare_legacy_vs_new.py --old path\\vers\\csv_legacy.csv --new path\\vers\\csv_nouveau.csv

L'objectif est d'offrir un point de départ léger pour vérifier que les
indicateurs principaux (nombre de lignes, sommes éventuelles) restent proches.
Ce script n'est pas appelé par les scripts de production.
"""

import argparse
from pathlib import Path

import pandas as pd


def load_csv(path: Path) -> pd.DataFrame:
    """Charge un CSV en devinant le séparateur (point-virgule ou virgule)."""
    raw = path.read_text(encoding="utf-8", errors="ignore")
    first = raw.splitlines()[0] if raw else ""
    sep = ";" if ";" in first else ","
    return pd.read_csv(path, sep=sep)


def summarize(df: pd.DataFrame) -> dict:
    """Retourne quelques indicateurs simples sur un DataFrame."""
    return {
        "nb_lignes": int(len(df)),
        "colonnes": list(df.columns),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Comparer un CSV legacy et un CSV nouveau.")
    parser.add_argument("--old", type=Path, required=True, help="CSV issu de l'ancien script (legacy).")
    parser.add_argument("--new", type=Path, required=True, help="CSV issu du nouveau moteur.")
    args = parser.parse_args()

    if not args.old.exists():
        print(f"Fichier legacy introuvable : {args.old}")
        return 1
    if not args.new.exists():
        print(f"Fichier nouveau introuvable : {args.new}")
        return 1

    old_df = load_csv(args.old)
    new_df = load_csv(args.new)

    old_summary = summarize(old_df)
    new_summary = summarize(new_df)

    print("=== Legacy ===")
    print(f"Nb lignes : {old_summary['nb_lignes']}")
    print(f"Colonnes : {', '.join(old_summary['colonnes'])}")

    print("\n=== Nouveau ===")
    print(f"Nb lignes : {new_summary['nb_lignes']}")
    print(f"Colonnes : {', '.join(new_summary['colonnes'])}")

    diff = new_summary["nb_lignes"] - old_summary["nb_lignes"]
    print(f"\nΔ nb_lignes (nouveau - legacy) : {diff}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

