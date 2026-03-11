"""Utilitaires partagés pour les bilans (filtrage, résumés, détection colonnes)."""
import re
from pathlib import Path
from typing import List

import geopandas as gpd
import pandas as pd

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_TYPES_USAGERS_PATH = _PROJECT_ROOT / "ref" / "types_usagers.csv"
_TYPES_USAGERS_MAPPING_CACHE: dict[tuple[str, str, str], str] | None = None


def _norm_key(s: str) -> str:
    return (s or "").strip().lower()


def _load_types_usagers_mapping() -> dict[tuple[str, str, str], str]:
    """Charge ref/types_usagers.csv et renvoie un mapping (source_table, source_champ, valeur_source_norm) -> type_usager."""
    global _TYPES_USAGERS_MAPPING_CACHE
    if _TYPES_USAGERS_MAPPING_CACHE is not None:
        return _TYPES_USAGERS_MAPPING_CACHE
    if not _TYPES_USAGERS_PATH.exists():
        _TYPES_USAGERS_MAPPING_CACHE = {}
        return _TYPES_USAGERS_MAPPING_CACHE
    df = pd.read_csv(_TYPES_USAGERS_PATH, sep=";", dtype=str, encoding="utf-8")
    df = df.fillna("")
    mapping: dict[tuple[str, str, str], str] = {}
    for _, r in df.iterrows():
        st = _norm_key(r.get("source_table", ""))
        sc = _norm_key(r.get("source_champ", ""))
        vs = _norm_key(r.get("valeur_source", ""))
        tu = (r.get("type_usager", "") or "").strip()
        if not st or not sc or not vs or not tu:
            continue
        mapping[(st, sc, vs)] = tu
    _TYPES_USAGERS_MAPPING_CACHE = mapping
    return mapping


def _parse_type_usager_tokens(valeur_source: str) -> list[tuple[str, int]]:
    """
    Parse une valeur OSCEAN de type_usager.

    Format observé :
    - \"Collectivité\" (sans effectif explicite)
    - \"Particulier (...) 6\"
    - \"Agriculteur ... 1, Collectivité 1, Particulier ... 1\"

    Renvoie une liste de (valeur_source_sans_effectif, effectif_int).
    """
    if pd.isna(valeur_source):
        return []
    s = str(valeur_source).strip()
    if not s or s == "(vide)":
        return []
    parts = [p.strip() for p in s.split(",") if p.strip()]
    out: list[tuple[str, int]] = []
    for p in parts:
        m = re.match(r"^(.*?)(?:\s+(\d+))?$", p)
        if not m:
            continue
        label = (m.group(1) or "").strip()
        n = int(m.group(2)) if m.group(2) and m.group(2).isdigit() else 1
        if label:
            out.append((label, n))
    return out


def map_type_usager(source_table: str, source_champ: str, valeur_source: str) -> str:
    """Mappe une valeur source vers un type d’usager (6 catégories cibles). Fallback : 'Autre'."""
    mapping = _load_types_usagers_mapping()
    key = (_norm_key(source_table), _norm_key(source_champ), _norm_key(valeur_source))
    return mapping.get(key, "Autre")


def serie_type_usager(df: pd.DataFrame, source_table: str, source_champ: str) -> pd.Series:
    """
    Déduit un type d’usager \"dominant\" par ligne à partir d’un champ source (ex. point_ctrl.type_usager).

    Règle :
    - si la ligne contient une seule catégorie → cette catégorie mappée ;
    - si plusieurs catégories → celle avec l'effectif max ; en cas d'égalité → 'Autre' ;
    - si vide → 'Autre'.
    """
    if source_champ not in df.columns:
        return pd.Series(["Autre"] * len(df), index=df.index, dtype="object")

    def _dominant(val: str) -> str:
        toks = _parse_type_usager_tokens(val)
        if not toks:
            return "Autre"
        # mapper chaque libellé vers une des 6 catégories
        mapped = [(map_type_usager(source_table, source_champ, lab), n) for lab, n in toks]
        # regrouper par catégorie (si doublons)
        agg: dict[str, int] = {}
        for cat, n in mapped:
            agg[cat] = agg.get(cat, 0) + int(n or 0)
        if len(agg) == 1:
            return next(iter(agg.keys()))
        # dominant
        max_n = max(agg.values())
        top = [k for k, v in agg.items() if v == max_n]
        return top[0] if len(top) == 1 else "Autre"

    return df[source_champ].apply(_dominant)


def filtre_periode(
    df: pd.DataFrame, col_date: str, date_deb: pd.Timestamp, date_fin: pd.Timestamp
) -> pd.DataFrame:
    """Filtre le DataFrame sur la plage de dates."""
    return df[(df[col_date] >= date_deb) & (df[col_date] <= date_fin)].copy()


def resume_resultat(s: pd.Series) -> str:
    """Consolide le résultat d'un dossier à partir des résultats de ses points."""
    vals = s.dropna()
    if vals.empty:
        return "Inconnu"
    if "Infraction" in vals.values:
        return "Infraction"
    if "Manquement" in vals.values:
        return "Manquement"
    mode = vals.mode()
    return mode.iloc[0] if not mode.empty else "Conforme"


def est_chasse_thematique(theme: str, type_action: str) -> bool:
    """Vérifie si le thème ou l'action concerne la chasse."""
    t = (theme or "").lower()
    a = (type_action or "").lower()
    return ("chasse" in t) or ("chasse" in a) or ("police de la chasse" in t)


def est_chasse_point(row: pd.Series) -> bool:
    """Détermine si un point de contrôle concerne la chasse."""
    return est_chasse_thematique(row.get("theme"), row.get("type_actio"))


def contient_natinf(s: str, natinf_list: List[str]) -> bool:
    """Vérifie si la chaîne contient l'un des codes NATINF (format X_Y ou isolé)."""
    s = str(s) if pd.notna(s) else ""
    for code in natinf_list:
        pattern = rf"(^|_){code}(_|$)"
        if re.search(pattern, s):
            return True
    return False


def _zone_summary(
    df: pd.DataFrame,
    col_insee: str,
    tub_codes: set,
    pnf_codes: set,
) -> pd.DataFrame:
    """Calcule nb_total, nb_conforme, nb_infraction pour dept / TUB / PNF."""
    insee = df[col_insee].astype(str).str.zfill(5)
    rows = []

    total = len(df)
    nb_inf_dept = (
        (df["resultat"].str.lower() == "infraction").sum()
        if "resultat" in df.columns
        else total
    )
    nb_conf_dept = total - nb_inf_dept
    rows.append(
        {
            "zone": "Département",
            "nb_total": total,
            "nb_conforme": nb_conf_dept,
            "nb_infraction": nb_inf_dept,
        }
    )

    mask_tub = insee.isin(tub_codes)
    sub_tub = df[mask_tub]
    nb_inf_tub = (
        (sub_tub["resultat"].str.lower() == "infraction").sum()
        if "resultat" in sub_tub.columns
        else len(sub_tub)
    )
    rows.append(
        {
            "zone": "Zone TUB",
            "nb_total": len(sub_tub),
            "nb_conforme": len(sub_tub) - nb_inf_tub,
            "nb_infraction": nb_inf_tub,
        }
    )

    mask_pnf = insee.isin(pnf_codes)
    sub_pnf = df[mask_pnf]
    nb_inf_pnf = (
        (sub_pnf["resultat"].str.lower() == "infraction").sum()
        if "resultat" in sub_pnf.columns
        else len(sub_pnf)
    )
    rows.append(
        {
            "zone": "PNF",
            "nb_total": len(sub_pnf),
            "nb_conforme": len(sub_pnf) - nb_inf_pnf,
            "nb_infraction": nb_inf_pnf,
        }
    )

    summary = pd.DataFrame(rows)
    summary["taux_infraction"] = (
        summary["nb_infraction"] / summary["nb_total"].replace(0, pd.NA)
    )
    return summary


def _zone_count(
    df: pd.DataFrame,
    col_insee: str,
    tub_codes: set,
    pnf_codes: set,
) -> pd.DataFrame:
    """Compte simple par zone (pour PVe / PEJ sans colonne 'resultat')."""
    insee = df[col_insee].astype(str).str.zfill(5)
    rows = [
        {"zone": "Département", "nb": len(df)},
        {"zone": "Zone TUB", "nb": int(insee.isin(tub_codes).sum())},
        {"zone": "PNF", "nb": int(insee.isin(pnf_codes).sum())},
    ]
    return pd.DataFrame(rows)


def _load_csv_opt(out_dir: Path, name: str) -> pd.DataFrame | None:
    """Charge un CSV optionnel ; retourne None si absent ou illisible."""
    p = out_dir / name
    if not p.exists():
        return None
    try:
        return pd.read_csv(p, sep=";", encoding="utf-8")
    except UnicodeDecodeError:
        return pd.read_csv(p, sep=";", encoding="latin-1")


def _detect_insee_column(communes: gpd.GeoDataFrame) -> str:
    """Détecte la colonne contenant le code INSEE dans une couche communes."""
    candidats = ["INSEE", "INSEE_COM", "CODE_INSEE", "INSEE_COMM", "INSEECO"]
    for col in candidats:
        if col in communes.columns:
            return col
    raise ValueError(
        "Impossible de trouver une colonne INSEE dans la couche communes."
    )
