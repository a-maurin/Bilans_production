"""
Finesse temporelle : évolution des contrôles, PVe et PEJ par mois et par trimestre.

Charge les données (points de contrôle, PVe, PEJ), ajoute des colonnes mois/trimestre
à partir des colonnes date, agrège les effectifs et exporte des CSV (par mois, par trimestre).
Sortie : out/bilan_temporel/.

Usage:
  python -m scripts.bilan_temporel.analyse_temporelle
  python -m scripts.bilan_temporel.analyse_temporelle --dept 21 --date-deb 2025-01-01 --date-fin 2026-02-05
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
from scripts.common.loaders import load_pej, load_point_ctrl, load_pve


def _add_periods(df: pd.DataFrame, col_date: str) -> pd.DataFrame:
    """Ajoute colonnes annee, mois, trimestre à partir de col_date."""
    df = df.copy()
    if col_date not in df.columns:
        return df
    dt = pd.to_datetime(df[col_date], errors="coerce")
    df["_annee"] = dt.dt.year
    df["_mois"] = dt.dt.month
    df["_trimestre"] = dt.dt.quarter
    df["_mois_str"] = dt.dt.to_period("M").astype(str)
    df["_trimestre_str"] = df["_annee"].astype(str) + "-T" + df["_trimestre"].astype(str)
    return df


def run_controles_par_periode(
    root: Path,
    out_dir: Path,
    dept_code: Optional[str],
    date_deb: Optional[pd.Timestamp],
    date_fin: Optional[pd.Timestamp],
) -> None:
    """Effectifs des points de contrôle par mois et par trimestre."""
    try:
        point = load_point_ctrl(
            root,
            dept_code=dept_code,
            date_deb=date_deb,
            date_fin=date_fin,
        )
    except FileNotFoundError as e:
        print(f"  [SKIP] Points de contrôle : {e}")
        return
    point = _add_periods(point, "date_ctrl")
    if "_mois_str" not in point.columns:
        return

    par_mois = point.groupby("_mois_str", dropna=False).size().reset_index(name="nb_controles")
    par_mois = par_mois.rename(columns={"_mois_str": "periode"})
    par_mois.to_csv(out_dir / "controles_par_mois.csv", sep=";", index=False, encoding="utf-8")

    par_tri = point.groupby("_trimestre_str", dropna=False).size().reset_index(name="nb_controles")
    par_tri = par_tri.rename(columns={"_trimestre_str": "periode"})
    par_tri.to_csv(out_dir / "controles_par_trimestre.csv", sep=";", index=False, encoding="utf-8")


def run_pve_par_periode(
    root: Path,
    out_dir: Path,
    dept_code: Optional[str],
    date_deb: Optional[pd.Timestamp],
    date_fin: Optional[pd.Timestamp],
) -> None:
    """Effectifs PVe par mois et par trimestre."""
    try:
        pve = load_pve(
            root,
            dept_code=dept_code,
            date_deb=date_deb,
            date_fin=date_fin,
        )
    except FileNotFoundError as e:
        print(f"  [SKIP] PVe : {e}")
        return
    if "INF-DATE-INTG" not in pve.columns:
        return
    pve = _add_periods(pve, "INF-DATE-INTG")
    if "_mois_str" not in pve.columns:
        return

    par_mois = pve.groupby("_mois_str", dropna=False).size().reset_index(name="nb_pve")
    par_mois = par_mois.rename(columns={"_mois_str": "periode"})
    par_mois.to_csv(out_dir / "pve_par_mois.csv", sep=";", index=False, encoding="utf-8")

    par_tri = pve.groupby("_trimestre_str", dropna=False).size().reset_index(name="nb_pve")
    par_tri = par_tri.rename(columns={"_trimestre_str": "periode"})
    par_tri.to_csv(out_dir / "pve_par_trimestre.csv", sep=";", index=False, encoding="utf-8")


def run_pej_par_periode(
    root: Path,
    out_dir: Path,
    date_deb: Optional[pd.Timestamp],
    date_fin: Optional[pd.Timestamp],
) -> None:
    """Effectifs PEJ par mois et par trimestre."""
    try:
        pej = load_pej(root, date_deb=date_deb, date_fin=date_fin)
    except FileNotFoundError as e:
        print(f"  [SKIP] PEJ : {e}")
        return
    if "DATE_REF" not in pej.columns:
        return
    pej = _add_periods(pej, "DATE_REF")
    if "_mois_str" not in pej.columns:
        return

    par_mois = pej.groupby("_mois_str", dropna=False).size().reset_index(name="nb_pej")
    par_mois = par_mois.rename(columns={"_mois_str": "periode"})
    par_mois.to_csv(out_dir / "pej_par_mois.csv", sep=";", index=False, encoding="utf-8")

    par_tri = pej.groupby("_trimestre_str", dropna=False).size().reset_index(name="nb_pej")
    par_tri = par_tri.rename(columns={"_trimestre_str": "periode"})
    par_tri.to_csv(out_dir / "pej_par_trimestre.csv", sep=";", index=False, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Évolution des contrôles, PVe et PEJ par mois et par trimestre."
    )
    parser.add_argument("--dept", type=str, default=None, help="Filtrer contrôles et PVe par département (ex: 21).")
    parser.add_argument("--date-deb", type=str, default=None, help="Date début YYYY-MM-DD.")
    parser.add_argument("--date-fin", type=str, default=None, help="Date fin YYYY-MM-DD.")
    parser.add_argument("--out-dir", type=str, default=None, help="Dossier de sortie (défaut: out/bilan_temporel).")
    args = parser.parse_args()

    root = PROJECT_ROOT
    out_dir = Path(args.out_dir) if args.out_dir else get_out_dir("bilan_temporel")
    out_dir.mkdir(parents=True, exist_ok=True)

    date_deb = pd.to_datetime(args.date_deb) if args.date_deb else None
    date_fin = pd.to_datetime(args.date_fin) if args.date_fin else None

    print("Finesse temporelle — contrôles, PVe, PEJ par mois/trimestre")
    print(f"Sortie : {out_dir}\n")

    print("Contrôles par mois et par trimestre...")
    run_controles_par_periode(root, out_dir, args.dept, date_deb, date_fin)

    print("PVe par mois et par trimestre...")
    run_pve_par_periode(root, out_dir, args.dept, date_deb, date_fin)

    print("PEJ par mois et par trimestre...")
    run_pej_par_periode(root, out_dir, date_deb, date_fin)

    print(f"\nTerminé. Fichiers dans {out_dir}")


if __name__ == "__main__":
    main()
