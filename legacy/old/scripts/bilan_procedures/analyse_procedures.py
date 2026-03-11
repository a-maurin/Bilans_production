"""
Analyses procédures PEJ : indicateurs sur DUREE_PEJ (moyenne, médiane, quantiles),
répartition par CLOTUR_PEJ et SUITE, éventuellement par THEME ou DOMAINE.

Sortie : out/bilan_procedures/.

Usage:
  python -m scripts.bilan_procedures.analyse_procedures
  python -m scripts.bilan_procedures.analyse_procedures --date-deb 2025-01-01 --date-fin 2026-02-05
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

import pandas as pd

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from scripts.paths import PROJECT_ROOT, get_out_dir
from scripts.common.loaders import load_pej


def run_analyse_procedures(
    root: Path,
    out_dir: Path,
    date_deb: Optional[pd.Timestamp],
    date_fin: Optional[pd.Timestamp],
) -> None:
    """Calcule indicateurs PEJ : durée (moy, médiane, quantiles), CLOTUR_PEJ, SUITE (global et par THEME)."""
    try:
        pej = load_pej(root, date_deb=date_deb, date_fin=date_fin)
    except FileNotFoundError as e:
        print(f"  [SKIP] PEJ : {e}")
        return

    duree_serie = pd.to_numeric(pej["DUREE_PEJ"], errors="coerce") if "DUREE_PEJ" in pej.columns else pd.Series(dtype=float)
    if duree_serie.notna().any():
        resume = pd.DataFrame([{
            "nb_pej": len(pej),
            "duree_moy_j": duree_serie.mean(),
            "duree_mediane_j": duree_serie.median(),
            "duree_p25_j": duree_serie.quantile(0.25),
            "duree_p75_j": duree_serie.quantile(0.75),
        }])
        resume.to_csv(out_dir / "pej_duree_resume.csv", sep=";", index=False, encoding="utf-8")
    else:
        pd.DataFrame([{"nb_pej": len(pej), "duree_moy_j": None, "duree_mediane_j": None}]).to_csv(
            out_dir / "pej_duree_resume.csv", sep=";", index=False, encoding="utf-8"
        )

    if "CLOTUR_PEJ" in pej.columns:
        clotur = pej["CLOTUR_PEJ"].fillna("(vide)").astype(str).value_counts().rename_axis("valeur").reset_index(name="nb")
        clotur.to_csv(out_dir / "pej_clotur_global.csv", sep=";", index=False, encoding="utf-8")

    if "SUITE" in pej.columns:
        suite = pej["SUITE"].fillna("(vide)").astype(str).value_counts().rename_axis("valeur").reset_index(name="nb")
        suite.to_csv(out_dir / "pej_suite_global.csv", sep=";", index=False, encoding="utf-8")

    if "THEME" in pej.columns and duree_serie.notna().any():
        pej_copy = pej.copy()
        pej_copy["_duree"] = duree_serie
        par_theme = pej_copy.groupby("THEME").agg(
            nb_pej=("DC_ID", "count"),
            duree_moy_j=("_duree", "mean"),
            duree_mediane_j=("_duree", "median"),
        ).reset_index()
        par_theme.to_csv(out_dir / "pej_duree_par_theme.csv", sep=";", index=False, encoding="utf-8")

    if "DOMAINE" in pej.columns and duree_serie.notna().any():
        pej_copy = pej.copy()
        pej_copy["_duree"] = duree_serie
        par_dom = pej_copy.groupby("DOMAINE").agg(
            nb_pej=("DC_ID", "count"),
            duree_moy_j=("_duree", "mean"),
            duree_mediane_j=("_duree", "median"),
        ).reset_index()
        par_dom.to_csv(out_dir / "pej_duree_par_domaine.csv", sep=";", index=False, encoding="utf-8")


def run_bilan(date_deb: str, date_fin: str, dept_code: str) -> int:
    """Entry point callable by the orchestrator."""
    root = PROJECT_ROOT
    out_dir = get_out_dir("bilan_procedures")
    out_dir.mkdir(parents=True, exist_ok=True)

    ts_deb = pd.to_datetime(date_deb) if date_deb else None
    ts_fin = pd.to_datetime(date_fin) if date_fin else None

    print("Analyses procédures PEJ (durée, clôture, suite)")
    print(f"Sortie : {out_dir}\n")

    run_analyse_procedures(root, out_dir, ts_deb, ts_fin)
    print(f"\nTerminé. Fichiers dans {out_dir}")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Indicateurs PEJ : durée (moyenne, médiane, quantiles), CLOTUR_PEJ, SUITE."
    )
    parser.add_argument("--date-deb", type=str, default=None, help="Date début YYYY-MM-DD.")
    parser.add_argument("--date-fin", type=str, default=None, help="Date fin YYYY-MM-DD.")
    parser.add_argument("--dept-code", type=str, default=None, help="Code département (ex. 21).")
    parser.add_argument("--out-dir", type=str, default=None, help="Dossier de sortie (défaut: out/bilan_procedures).")
    args = parser.parse_args()

    root = PROJECT_ROOT
    out_dir = Path(args.out_dir) if args.out_dir else get_out_dir("bilan_procedures")
    out_dir.mkdir(parents=True, exist_ok=True)

    date_deb = pd.to_datetime(args.date_deb) if args.date_deb else None
    date_fin = pd.to_datetime(args.date_fin) if args.date_fin else None

    print("Analyses procédures PEJ (durée, clôture, suite)")
    print(f"Sortie : {out_dir}\n")

    run_analyse_procedures(root, out_dir, date_deb, date_fin)
    print(f"\nTerminé. Fichiers dans {out_dir}")


if __name__ == "__main__":
    main()
