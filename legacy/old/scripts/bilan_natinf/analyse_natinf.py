"""
Analyses PVe et PEJ par NATINF.

Produit des tableaux par code NATINF (effectifs PVe/PEJ, répartition par zone TUB/PNF,
durée moyenne PEJ, répartition par CLOTUR_PEJ et SUITE) avec libellés issus de
ref/liste_natinf.csv. Sorties CSV dans out/bilan_natinf/.

Usage:
  python -m scripts.bilan_natinf.analyse_natinf
  python -m scripts.bilan_natinf.analyse_natinf --dept 21 --date-deb 2025-01-01 --date-fin 2026-02-05
  python -m scripts.bilan_natinf.analyse_natinf --natinf 27742,25001,321
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import List, Optional

import pandas as pd

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from scripts.paths import PROJECT_ROOT, get_out_dir
from scripts.common.loaders import load_pej, load_pve, load_pnf, load_tub, load_tub_pnf_codes
from scripts.common.utils import _zone_count


def _load_natinf_ref(root: Path) -> pd.DataFrame:
    """Charge le référentiel NATINF (ref/liste_natinf.csv)."""
    for base in ("ref", "sources"):
        for name in ("liste_natinf.csv", "liste-natinf-avril2023.csv"):
            path = root / base / name
            if not path.exists():
                continue
            try:
                raw = path.read_text(encoding="utf-8", errors="ignore")
                sep = ";" if ";" in raw.split("\n")[0] else ","
                df = pd.read_csv(path, sep=sep, dtype=str, encoding="utf-8", on_bad_lines="skip")
                for c in ("Numéro NATINF", "NATINF", "natinf"):
                    if c in df.columns:
                        df = df.rename(columns={c: "numero_natinf"})
                        break
                if "numero_natinf" not in df.columns:
                    continue
                lib_col = None
                for c in df.columns:
                    if "nature" in c.lower() or "infraction" in c.lower():
                        lib_col = c
                        break
                if lib_col:
                    df = df.rename(columns={lib_col: "libelle_natinf"})
                if "libelle_natinf" not in df.columns:
                    df["libelle_natinf"] = ""
                return df[["numero_natinf", "libelle_natinf"]].drop_duplicates()
            except Exception:
                continue
    return pd.DataFrame(columns=["numero_natinf", "libelle_natinf"])


def _natinf_codes_from_series(series: pd.Series) -> set[str]:
    """Extrait les codes NATINF numériques d'une série (ex. '27742', '27742_25001')."""
    codes = set()
    for v in series.dropna().astype(str).unique():
        for part in re.split(r"[\s_,;]+", v):
            part = part.strip()
            if part.isdigit():
                codes.add(part)
    return codes


def _libelle_for_code(code: str, natinf_ref: pd.DataFrame) -> str:
    """Retourne le libellé NATINF pour un code (ref/liste_natinf.csv)."""
    if natinf_ref.empty or "libelle_natinf" not in natinf_ref.columns:
        return ""
    match = natinf_ref[natinf_ref["numero_natinf"].astype(str).str.strip() == code]
    if match.empty:
        return ""
    return str(match["libelle_natinf"].iloc[0])[:80]


def _row_contains_natinf(val: str, code: str) -> bool:
    """True si la valeur contient le code NATINF (ex. 27742 dans '27742' ou '27742_25001')."""
    if pd.isna(val):
        return False
    s = str(val).strip()
    return bool(re.search(rf"(^|[\s_,;]){re.escape(code)}([\s_,;]|$)", s))


def run_analyse_pve_par_natinf(
    root: Path,
    out_dir: Path,
    dept_code: Optional[str],
    date_deb: Optional[pd.Timestamp],
    date_fin: Optional[pd.Timestamp],
    natinf_list: Optional[List[str]],
    natinf_ref: pd.DataFrame,
) -> None:
    """Effectifs PVe par NATINF et par zone (Département, TUB, PNF)."""
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
    if "INF-NATINF" not in pve.columns:
        print("  [SKIP] PVe : colonne INF-NATINF absente")
        return

    tub_codes, pnf_codes = load_tub_pnf_codes(root)

    if natinf_list is None:
        natinf_list = sorted(_natinf_codes_from_series(pve["INF-NATINF"]))
    if not natinf_list:
        print("  [SKIP] PVe : aucun code NATINF à analyser")
        return

    # Tableau global : un enregistrement par NATINF
    rows_global = []
    zone_rows = []

    for code in natinf_list:
        mask = pve["INF-NATINF"].apply(lambda v: _row_contains_natinf(v, code))
        sub = pve[mask].copy()
        nb = len(sub)
        libelle = _libelle_for_code(code, natinf_ref)
        rows_global.append({"natinf": code, "libelle": libelle, "nb_pve": nb})

        if sub.empty:
            zone_rows.append({"natinf": code, "libelle": libelle, "zone": "Département", "nb": 0})
            zone_rows.append({"natinf": code, "libelle": libelle, "zone": "Zone TUB", "nb": 0})
            zone_rows.append({"natinf": code, "libelle": libelle, "zone": "PNF", "nb": 0})
            continue
        if "INF-INSEE" not in sub.columns:
            sub["INF-INSEE"] = ""
        sub["INF-INSEE"] = sub["INF-INSEE"].astype(str).str.zfill(5)
        zc = _zone_count(sub, "INF-INSEE", tub_codes, pnf_codes)
        zc["natinf"] = code
        zc["libelle"] = libelle
        zc = zc.rename(columns={"nb": "nb"})
        zone_rows.extend(zc.to_dict("records"))

    pd.DataFrame(rows_global).to_csv(out_dir / "pve_par_natinf.csv", sep=";", index=False, encoding="utf-8")
    pd.DataFrame(zone_rows).to_csv(out_dir / "pve_par_natinf_zone.csv", sep=";", index=False, encoding="utf-8")


def run_analyse_pej_par_natinf(
    root: Path,
    out_dir: Path,
    date_deb: Optional[pd.Timestamp],
    date_fin: Optional[pd.Timestamp],
    natinf_list: Optional[List[str]],
    natinf_ref: pd.DataFrame,
) -> None:
    """Effectifs PEJ par NATINF, durée moyenne, répartition par CLOTUR_PEJ et SUITE."""
    try:
        pej = load_pej(root, date_deb=date_deb, date_fin=date_fin)
    except FileNotFoundError as e:
        print(f"  [SKIP] PEJ : {e}")
        return
    if "NATINF_PEJ" not in pej.columns:
        print("  [SKIP] PEJ : colonne NATINF_PEJ absente")
        return

    if natinf_list is None:
        natinf_list = sorted(_natinf_codes_from_series(pej["NATINF_PEJ"]))
    if not natinf_list:
        print("  [SKIP] PEJ : aucun code NATINF à analyser")
        return

    rows_global = []
    rows_clotur = []
    rows_suite = []
    rows_theme = []

    for code in natinf_list:
        mask = pej["NATINF_PEJ"].apply(lambda v: _row_contains_natinf(v, code))
        sub = pej[mask].copy()
        nb = len(sub)
        libelle = _libelle_for_code(code, natinf_ref)

        duree_moy = None
        if "DUREE_PEJ" in sub.columns:
            duree_moy = pd.to_numeric(sub["DUREE_PEJ"], errors="coerce").mean()
        rows_global.append({
            "natinf": code,
            "libelle": libelle,
            "nb_pej": nb,
            "duree_moy_pej_j": duree_moy,
        })

        if "CLOTUR_PEJ" in sub.columns:
            vc = sub["CLOTUR_PEJ"].fillna("(vide)").astype(str).value_counts()
            for val, cnt in vc.items():
                rows_clotur.append({"natinf": code, "libelle": libelle, "clotur_pej": val, "nb": int(cnt)})
        if "SUITE" in sub.columns:
            vc = sub["SUITE"].fillna("(vide)").astype(str).value_counts()
            for val, cnt in vc.items():
                rows_suite.append({"natinf": code, "libelle": libelle, "suite": val, "nb": int(cnt)})
        if "THEME" in sub.columns or "DOMAINE" in sub.columns:
            cols_grp = [c for c in ("DOMAINE", "THEME") if c in sub.columns]
            if cols_grp:
                grp = sub.groupby(cols_grp).size().reset_index(name="nb")
                grp["natinf"] = code
                grp["libelle"] = libelle
                for _, r in grp.iterrows():
                    rows_theme.append({
                        "natinf": code,
                        "libelle": libelle,
                        **{c: r[c] for c in cols_grp},
                        "nb": int(r["nb"]),
                    })

    pd.DataFrame(rows_global).to_csv(out_dir / "pej_par_natinf.csv", sep=";", index=False, encoding="utf-8")
    if rows_clotur:
        pd.DataFrame(rows_clotur).to_csv(out_dir / "pej_par_natinf_clotur.csv", sep=";", index=False, encoding="utf-8")
    if rows_suite:
        pd.DataFrame(rows_suite).to_csv(out_dir / "pej_par_natinf_suite.csv", sep=";", index=False, encoding="utf-8")
    if rows_theme:
        pd.DataFrame(rows_theme).to_csv(out_dir / "pej_par_natinf_theme.csv", sep=";", index=False, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyses PVe et PEJ par NATINF (avec libellés ref/liste_natinf.csv).")
    parser.add_argument("--dept", type=str, default=None, help="Filtrer PVe par département (ex: 21).")
    parser.add_argument("--date-deb", type=str, default=None, help="Date début YYYY-MM-DD.")
    parser.add_argument("--date-fin", type=str, default=None, help="Date fin YYYY-MM-DD.")
    parser.add_argument(
        "--natinf",
        type=str,
        default=None,
        help="Liste de codes NATINF séparés par des virgules (ex: 27742,25001). Si absent, tous les NATINF présents dans les données.",
    )
    parser.add_argument("--out-dir", type=str, default=None, help="Dossier de sortie (défaut: out/bilan_natinf).")
    args = parser.parse_args()

    root = PROJECT_ROOT
    out_dir = Path(args.out_dir) if args.out_dir else get_out_dir("bilan_natinf")
    out_dir.mkdir(parents=True, exist_ok=True)

    date_deb = pd.to_datetime(args.date_deb) if args.date_deb else None
    date_fin = pd.to_datetime(args.date_fin) if args.date_fin else None
    natinf_list = [s.strip() for s in args.natinf.split(",")] if args.natinf else None

    print("Analyses PVe et PEJ par NATINF")
    print(f"Sortie : {out_dir}\n")

    natinf_ref = _load_natinf_ref(root)
    if not natinf_ref.empty:
        print("Référentiel NATINF chargé (libellés).")

    print("PVe par NATINF et par zone...")
    run_analyse_pve_par_natinf(root, out_dir, args.dept, date_deb, date_fin, natinf_list, natinf_ref)

    print("PEJ par NATINF (effectifs, durée, clôture, suite, thème)...")
    run_analyse_pej_par_natinf(root, out_dir, date_deb, date_fin, natinf_list, natinf_ref)

    print(f"\nTerminé. Fichiers dans {out_dir}")


if __name__ == "__main__":
    main()
