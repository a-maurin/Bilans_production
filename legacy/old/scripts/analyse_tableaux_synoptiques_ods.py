"""
Analyse du fichier ODS « tableaux synoptiques police SD21 » dans ref/.

Lit exemple_tableaux_synoptiques_police_sd21.ods, liste les feuilles,
affiche la structure (colonnes, types, dimensions) et des statistiques
par feuille. Peut exporter un résumé en CSV ou afficher en console.

Usage:
  python -m scripts.analyse_tableaux_synoptiques_ods
  python -m scripts.analyse_tableaux_synoptiques_ods -v -o out/resume_synoptiques.csv
"""
from pathlib import Path
import argparse
import sys

import pandas as pd

# Bootstrap chemin projet (comme scripts.inventaire)
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
from scripts.paths import get_ref_dir

FICHIER_ODS_DEFAULT = "exemple_tableaux_synoptiques_police_sd21.ods"


def get_ods_path(fichier: str | None) -> Path:
    """Retourne le chemin absolu du fichier ODS (ref/ par défaut)."""
    if fichier is None:
        return get_ref_dir() / FICHIER_ODS_DEFAULT
    p = Path(fichier)
    if not p.is_absolute():
        p = get_ref_dir() / p
    return p


def charger_ods(chemin: Path) -> dict[str, pd.DataFrame]:
    """Charge toutes les feuilles du classeur ODS."""
    if not chemin.exists():
        raise FileNotFoundError(f"Fichier introuvable : {chemin}")
    # sheet_name=None => toutes les feuilles
    dfs = pd.read_excel(chemin, sheet_name=None, dtype=str, engine="odf")
    return dfs


def analyser_feuille(nom: str, df: pd.DataFrame) -> dict:
    """Retourne un résumé structuré pour une feuille."""
    return {
        "feuille": nom,
        "lignes": len(df),
        "colonnes": len(df.columns),
        "noms_colonnes": list(df.columns),
        "dtypes": df.dtypes.astype(str).to_dict(),
        "lignes_vides": int(df.isna().all(axis=1).sum()),
        "colonnes_entièrement_vides": int((df.isna().all(axis=0)).sum()),
    }


def afficher_analyse(chemin: Path, dfs: dict[str, pd.DataFrame], verbose: bool = False) -> None:
    """Affiche l'analyse en console."""
    print(f"Fichier : {chemin}")
    print(f"Feuilles : {list(dfs.keys())}")
    print("-" * 60)
    for nom, df in dfs.items():
        info = analyser_feuille(nom, df)
        print(f"\n[Feuille « {nom} »]")
        print(f"  Dimensions : {info['lignes']} lignes × {info['colonnes']} colonnes")
        print(f"  Lignes entièrement vides : {info['lignes_vides']}")
        print(f"  Colonnes : {info['noms_colonnes']}")
        if verbose:
            print("  Types (échantillon) :")
            for c, t in list(info["dtypes"].items())[:15]:
                print(f"    - {c}: {t}")
            if len(info["dtypes"]) > 15:
                print(f"    ... et {len(info['dtypes']) - 15} autres")
        print("\n  Aperçu (5 premières lignes) :")
        print(df.head().to_string(index=False))
        # Colonnes avec beaucoup de valeurs distinctes / vides
        if verbose and len(df) > 0:
            non_vides = (df.notna() & (df != "")).sum()
            print("\n  Taux de remplissage par colonne :")
            for col in df.columns[:20]:
                pct = 100.0 * non_vides[col] / len(df)
                print(f"    {col}: {pct:.1f}%")
            if len(df.columns) > 20:
                print(f"    ... et {len(df.columns) - 20} autres colonnes")
        print()


def exporter_resume(chemin_out: Path, dfs: dict[str, pd.DataFrame]) -> None:
    """Exporte un résumé (une ligne par feuille) en CSV."""
    lignes = []
    for nom, df in dfs.items():
        info = analyser_feuille(nom, df)
        lignes.append({
            "feuille": nom,
            "lignes": info["lignes"],
            "colonnes": info["colonnes"],
            "lignes_vides": info["lignes_vides"],
        })
    pd.DataFrame(lignes).to_csv(chemin_out, index=False, sep=";", encoding="utf-8-sig")
    print(f"Résumé exporté : {chemin_out}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Analyse du fichier ODS tableaux synoptiques police SD21 (ref/)."
    )
    parser.add_argument(
        "fichier",
        nargs="?",
        default=None,
        help=f"Fichier ODS à analyser (défaut : ref/{FICHIER_ODS_DEFAULT})",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Afficher types et taux de remplissage par colonne",
    )
    parser.add_argument(
        "-o", "--out",
        metavar="CSV",
        help="Exporter un résumé (une ligne par feuille) vers ce fichier CSV",
    )
    args = parser.parse_args()
    chemin = get_ods_path(args.fichier)
    try:
        dfs = charger_ods(chemin)
    except FileNotFoundError as e:
        print(e, file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier : {e}", file=sys.stderr)
        return 1
    afficher_analyse(chemin, dfs, verbose=args.verbose)
    if args.out:
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        exporter_resume(out_path, dfs)
    return 0


if __name__ == "__main__":
    sys.exit(main())
