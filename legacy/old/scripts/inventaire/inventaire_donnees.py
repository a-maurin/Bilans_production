"""
Inventaire des données sources (phase 0).

Lit les données OSCEAN (points de contrôle, PEJ, PA, PVe, points infractions PJ),
produit pour chaque source les valeurs distinctes et effectifs des champs clés,
et exporte un rapport CSV + résumé texte. N'altère pas les bilans existants.

Usage:
  python -m scripts.inventaire.inventaire_donnees
  python -m scripts.inventaire.inventaire_donnees --dept 21 --out-dir out/inventaire
  python -m scripts.inventaire.inventaire_donnees --date-deb 2025-01-01 --date-fin 2026-02-05
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd

# Bootstrap chemin projet
_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from scripts.paths import PROJECT_ROOT
from scripts.common.loaders import (
    load_pa,
    load_pej,
    load_point_ctrl,
    load_pve,
    get_points_infrac_pj_path,
)
import geopandas as gpd


# Colonnes à inventorier par source
COLS_POINT_CTRL = [
    "theme",
    "type_actio",
    "domaine",
    "nom_dossie",
    "entit_ctrl",
    # Champ OSCEAN décrivant les usagers (peut contenir plusieurs catégories + effectifs)
    "type_usager",
    "fc_type",
    "resultat",
]
COLS_PEJ = ["THEME", "TYPE_ACTION", "DOMAINE", "NATINF_PEJ", "CLOTUR_PEJ", "SUITE"]
COLS_PA = ["THEME", "TYPE_ACTION", "ENTITE_ORIGINE_PROCEDURE"]
COLS_PVE = ["INF-NATINF", "INF-TYP-INF-STAT-LIB"]
COLS_PJ = ["natinf", "entite"]

USAGERS_COL_REGEX = re.compile(r"(usag|acteur|acte?ur|public|categorie|cat[ée]gorie)", re.IGNORECASE)


def _value_counts_df(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """Valeurs distinctes et effectifs pour une colonne (gère les absences)."""
    if col not in df.columns:
        return pd.DataFrame(columns=["valeur", "effectif"])
    s = df[col].fillna("(vide)").astype(str).str.strip()
    vc = s.value_counts(dropna=False).rename_axis("valeur").reset_index(name="effectif")
    return vc


def _inventaire_source(
    out_dir: Path,
    source_name: str,
    df: pd.DataFrame,
    cols: list[str],
    prefix: str = "inventaire",
) -> list[Path]:
    """Pour une source et une liste de colonnes, exporte un CSV par colonne."""
    written = []
    for col in cols:
        if col not in df.columns:
            continue
        vc = _value_counts_df(df, col)
        if vc.empty:
            continue
        vc["source"] = source_name
        vc["champ"] = col
        fname = f"{prefix}_{source_name}_{col}.csv".replace(" ", "_").replace("-", "_")
        path = out_dir / fname
        vc.to_csv(path, sep=";", index=False, encoding="utf-8")
        written.append(path)
    return written


def _load_natinf_libelles(root: Path) -> pd.DataFrame | None:
    """Charge ref/liste_natinf.csv (ou sources/liste_natinf.csv) pour libeller les NATINF dans les exports PEJ/PVe."""
    for base in ("ref", "sources"):
        for name in ("liste_natinf.csv", "liste-natinf-avril2023.csv"):
            path = root / base / name
            if not path.exists():
                continue
            try:
                raw = path.read_text(encoding="utf-8", errors="ignore")
                sep = ";" if ";" in raw.split("\n")[0] else ","
                df = pd.read_csv(path, sep=sep, dtype=str, encoding="utf-8", on_bad_lines="skip")
                for c in ("Numéro NATINF", "numero_natinf", "NATINF", "natinf"):
                    if c in df.columns:
                        df = df.rename(columns={c: "numero_natinf"})
                        break
                if "numero_natinf" not in df.columns:
                    continue
                # Normaliser la colonne libellé comme dans bilan_natinf pour export cohérent
                lib_col = None
                for c in df.columns:
                    if c == "numero_natinf":
                        continue
                    if "nature" in c.lower() or "infraction" in c.lower():
                        lib_col = c
                        break
                if lib_col:
                    df = df.rename(columns={lib_col: "libelle_natinf"})
                else:
                    df["libelle_natinf"] = ""
                return df[["numero_natinf", "libelle_natinf"]].drop_duplicates()
            except Exception:
                continue
    return None


def run_inventaire_point_ctrl(
    root: Path,
    out_dir: Path,
    dept_code: str | None = None,
    date_deb: str | None = None,
    date_fin: str | None = None,
) -> list[Path]:
    """Inventaire des points de contrôle."""
    try:
        point = load_point_ctrl(
            root,
            dept_code=dept_code,
            date_deb=pd.to_datetime(date_deb) if date_deb else None,
            date_fin=pd.to_datetime(date_fin) if date_fin else None,
        )
    except FileNotFoundError as e:
        print(f"  [SKIP] Points de contrôle : {e}")
        return []
    cols = [c for c in COLS_POINT_CTRL if c in point.columns]
    return _inventaire_source(out_dir, "point_ctrl", point, cols)


def run_inventaire_pej(
    root: Path,
    out_dir: Path,
    natinf_ref: pd.DataFrame | None,
) -> list[Path]:
    """Inventaire des PEJ (sans filtre date pour voir toutes les valeurs)."""
    try:
        pej = load_pej(root, date_deb=None, date_fin=None)
    except FileNotFoundError as e:
        print(f"  [SKIP] PEJ : {e}")
        return []
    written = []
    for col in COLS_PEJ:
        if col not in pej.columns:
            continue
        vc = _value_counts_df(pej, col)
        if vc.empty:
            continue
        vc["source"] = "pej"
        vc["champ"] = col
        if col == "NATINF_PEJ" and natinf_ref is not None and "libelle_natinf" in natinf_ref.columns:
            num_col = "numero_natinf"
            ref_short = natinf_ref[[num_col, "libelle_natinf"]].drop_duplicates()
            ref_short[num_col] = ref_short[num_col].astype(str).str.strip()
            vc["valeur_clean"] = vc["valeur"].str.replace("(vide)", "").str.extract(r"(\d+)", expand=False)
            vc = vc.merge(
                ref_short,
                left_on="valeur_clean",
                right_on=num_col,
                how="left",
            )
            vc = vc.drop(columns=[c for c in ("valeur_clean", num_col) if c in vc.columns], errors="ignore")
        fname = f"inventaire_pej_{col}.csv"
        path = out_dir / fname
        vc.to_csv(path, sep=";", index=False, encoding="utf-8")
        written.append(path)
    return written


def run_inventaire_pa(root: Path, out_dir: Path) -> list[Path]:
    """Inventaire des PA."""
    try:
        pa = load_pa(root, date_deb=None, date_fin=None)
    except FileNotFoundError as e:
        print(f"  [SKIP] PA : {e}")
        return []
    cols = [c for c in COLS_PA if c in pa.columns]
    return _inventaire_source(out_dir, "pa", pa, cols)


def run_inventaire_pve(root: Path, out_dir: Path, natinf_ref: pd.DataFrame | None) -> list[Path]:
    """Inventaire des PVe (sans filtre département pour voir tous NATINF)."""
    try:
        pve = load_pve(root, dept_code=None, date_deb=None, date_fin=None)
    except FileNotFoundError as e:
        print(f"  [SKIP] PVe : {e}")
        return []
    written = []
    for col in COLS_PVE:
        if col not in pve.columns:
            continue
        vc = _value_counts_df(pve, col)
        if vc.empty:
            continue
        vc["source"] = "pve"
        vc["champ"] = col
        if col == "INF-NATINF" and natinf_ref is not None and "libelle_natinf" in natinf_ref.columns:
            num_col = "numero_natinf"
            ref_short = natinf_ref[[num_col, "libelle_natinf"]].drop_duplicates()
            ref_short[num_col] = ref_short[num_col].astype(str).str.strip()
            vc["valeur_clean"] = vc["valeur"].replace("(vide)", "").str.extract(r"(\d+)", expand=False)
            vc = vc.merge(
                ref_short,
                left_on="valeur_clean",
                right_on=num_col,
                how="left",
            )
            vc = vc.drop(columns=[c for c in ("valeur_clean", num_col) if c in vc.columns], errors="ignore")
        fname = f"inventaire_pve_{col.replace('-', '_')}.csv"
        path = out_dir / fname
        vc.to_csv(path, sep=";", index=False, encoding="utf-8")
        written.append(path)
    return written


def run_inventaire_points_pj(root: Path, out_dir: Path) -> list[Path]:
    """Inventaire des points infractions PJ (lecture directe du GPKG sans filtre NATINF)."""
    path_gpkg = get_points_infrac_pj_path(root)
    if not path_gpkg.exists():
        print(f"  [SKIP] Points infractions PJ : fichier absent ({path_gpkg})")
        return []
    try:
        gdf = gpd.read_file(path_gpkg)
    except Exception as e:
        print(f"  [SKIP] Points infractions PJ : {e}")
        return []
    df = pd.DataFrame(gdf.drop(columns=["geometry"], errors="ignore"))
    written = []
    for col in COLS_PJ:
        if col not in df.columns:
            continue
        vc = _value_counts_df(df, col)
        if vc.empty:
            continue
        vc["source"] = "points_infrac_pj"
        vc["champ"] = col
        fname = f"inventaire_points_infrac_pj_{col}.csv"
        path = out_dir / fname
        vc.to_csv(path, sep=";", index=False, encoding="utf-8")
        written.append(path)
    return written


def write_resume(out_dir: Path, written: list[Path], opts: dict) -> Path:
    """Écrit un fichier résumé texte."""
    path = out_dir / "inventaire_resume.txt"
    with open(path, "w", encoding="utf-8") as f:
        f.write("Inventaire des données sources — Bilans_production\n")
        f.write("=" * 60 + "\n")
        f.write(f"Généré le : {datetime.now().isoformat(timespec='seconds')}\n")
        f.write(f"Options : dept={opts.get('dept')}, date_deb={opts.get('date_deb')}, date_fin={opts.get('date_fin')}\n\n")
        f.write("Fichiers CSV générés :\n")
        for p in sorted(written):
            f.write(f"  - {p.name}\n")
        f.write("\nUtilisez ces CSV pour valider les thèmes, NATINF et entités avant de coder les filtres des bilans.\n")
    return path


def write_usagers_sources_csv(out_dir: Path, written: list[Path]) -> Path:
    """
    Construit un CSV consolidé des valeurs liées aux usagers (toutes sources confondues),
    à partir des CSV générés par l'inventaire.

    Colonnes : source, champ, valeur_source, effectif
    """
    frames: list[pd.DataFrame] = []
    for p in written:
        name = p.name.lower()
        # Heuristique : prendre les inventaires dont le champ ou le fichier suggère un lien usagers/acteurs
        if ("type_usage" in name) or ("usag" in name) or ("acteur" in name):
            try:
                df = pd.read_csv(p, sep=";", dtype=str, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_csv(p, sep=";", dtype=str, encoding="latin-1")
            if df.empty:
                continue
            df = df.rename(columns={"valeur": "valeur_source"})
            keep = [c for c in ["source", "champ", "valeur_source", "effectif"] if c in df.columns]
            if not keep:
                continue
            frames.append(df[keep])
    out = out_dir / "usagers_sources.csv"
    if frames:
        pd.concat(frames, ignore_index=True).to_csv(out, sep=";", index=False, encoding="utf-8")
    else:
        pd.DataFrame(columns=["source", "champ", "valeur_source", "effectif"]).to_csv(
            out, sep=";", index=False, encoding="utf-8"
        )
    return out


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Inventaire des champs et valeurs des données sources (phase 0)."
    )
    parser.add_argument(
        "--dept",
        type=str,
        default=None,
        help="Limiter les points de contrôle à ce département (ex: 21). Si absent, tous les départements.",
    )
    parser.add_argument(
        "--date-deb",
        type=str,
        default=None,
        help="Date début (YYYY-MM-DD) pour filtrer les points de contrôle.",
    )
    parser.add_argument(
        "--date-fin",
        type=str,
        default=None,
        help="Date fin (YYYY-MM-DD) pour filtrer les points de contrôle.",
    )
    parser.add_argument(
        "--out-dir",
        type=str,
        default=None,
        help="Dossier de sortie (défaut: out/inventaire).",
    )
    args = parser.parse_args()

    root = PROJECT_ROOT
    out_dir = Path(args.out_dir) if args.out_dir else root / "out" / "inventaire"
    out_dir.mkdir(parents=True, exist_ok=True)

    opts = {"dept": args.dept, "date_deb": args.date_deb, "date_fin": args.date_fin}
    print("Inventaire des données sources (phase 0)")
    print(f"Sortie : {out_dir}\n")

    natinf_ref = _load_natinf_libelles(root)
    if natinf_ref is not None:
        print("Référentiel NATINF chargé pour libellés.")

    written: list[Path] = []

    print("Points de contrôle...")
    written.extend(
        run_inventaire_point_ctrl(root, out_dir, args.dept, args.date_deb, args.date_fin)
    )

    print("PEJ...")
    written.extend(run_inventaire_pej(root, out_dir, natinf_ref))

    print("PA...")
    written.extend(run_inventaire_pa(root, out_dir))

    print("PVe...")
    written.extend(run_inventaire_pve(root, out_dir, natinf_ref))

    print("Points infractions PJ...")
    written.extend(run_inventaire_points_pj(root, out_dir))

    write_resume(out_dir, written, opts)
    write_usagers_sources_csv(out_dir, written)
    print(f"\nTerminé. {len(written)} fichier(s) CSV + résumé dans {out_dir}")


if __name__ == "__main__":
    main()
