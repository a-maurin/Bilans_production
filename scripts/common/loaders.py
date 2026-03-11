"""Chargement des données OSCEAN (points de contrôle, PEJ, PA, PVe, PNF, TUB, infractions PJ)."""
import sys
from pathlib import Path
from typing import List, Optional, Tuple, Union

import geopandas as gpd
import pandas as pd
import re

from scripts.common.utils import filtre_periode


def _gpkg_engine() -> str:
    """Detect the best available GPKG engine once (pyogrio > fiona)."""
    try:
        import pyogrio  # noqa: F401
        return "pyogrio"
    except ImportError:
        return "fiona"


_GPKG_ENGINE: str = _gpkg_engine()


def _find_latest_dated_file(directory: Path, prefix: str, exts: Tuple[str, ...]) -> Path:
    """
    Retourne, parmi les fichiers de `directory` commençant par `prefix` et se
    terminant par l'une des extensions de `exts`, celui dont la date suffixe
    (format YYYYMMDD) est la plus récente.
    """
    latest_path: Path | None = None
    latest_date: str | None = None

    for ext in exts:
        pattern = f"{prefix}*{ext}"
        for p in directory.glob(pattern):
            m = re.match(re.escape(prefix) + r"(\d{8})", p.name)
            if not m:
                continue
            date_str = m.group(1)
            if latest_date is None or date_str > latest_date:
                latest_date = date_str
                latest_path = p

    if latest_path is None:
        raise FileNotFoundError(
            f"Aucun fichier trouvé dans {directory} pour le préfixe '{prefix}' "
            f"et les extensions {exts}."
        )
    return latest_path


def load_point_ctrl(
    root: Path,
    dept_code: Optional[str] = None,
    date_deb: Optional[Union[str, pd.Timestamp]] = None,
    date_fin: Optional[Union[str, pd.Timestamp]] = None,
) -> pd.DataFrame:
    """
    Charge les points de contrôle OSCEAN à partir des fichiers
    point_ctrl_YYYYMMDD_wgs84.gpkg présents dans sources/sig.

    Pour chaque année, seul le GPKG le plus récent (date suffixe la plus élevée)
    est utilisé, afin d'éviter les doublons liés aux extractions mensuelles.

    Si dept_code, date_deb et date_fin sont fournis, seules les années
    recouvrant la période sont chargées et les lignes sont filtrées par
    département et période (accélère l'analyse sur sources nationales).
    """
    sources_sig = root / "sources" / "sig"
    if not sources_sig.exists():
        raise FileNotFoundError(
            f"Le dossier des données SIG sources n'existe pas : {sources_sig}"
        )

    # Nouvelle organisation : les fichiers de points de contrôle sont rangés par année
    # dans des sous-dossiers de type points_de_ctrl_OSCEAN_YYYY sous sources/sig.
    # Les couches peuvent être au format GPKG ou SHP et le nommage peut varier
    # (point/pts/pt, controle/ctrl, etc.). On doit donc rechercher de manière
    # plus souple que le strict pattern point_ctrl_YYYYMMDD_wgs84.*
    candidates: List[Path] = []
    year_dirs = [
        d
        for d in sources_sig.glob("points_de_ctrl_OSCEAN_*")
        if d.is_dir()
    ]
    for d in year_dirs:
        for ext in (".gpkg", ".shp"):
            # Noms usuels : point_ctrl_..., pts_ctrl_..., pt_controle_..., etc.
            for pat in ("*point*ctrl*",
                        "*pts*ctrl*",
                        "*pt*ctrl*",
                        "*ctrl*"):
                candidates.extend(d.glob(f"{pat}{ext}"))

    # Rétrocompatibilité : si aucun sous-dossier n'est trouvé (ancienne arborescence
    # ou environnement non migré), on cherche encore à la racine de sources/sig.
    if not candidates:
        candidates = []
        for ext in (".gpkg", ".shp"):
            for pat in ("*point*ctrl*",
                        "*pts*ctrl*",
                        "*pt*ctrl*",
                        "*ctrl*"):
                candidates.extend(sources_sig.glob(f"{pat}{ext}"))

    if not candidates:
        raise FileNotFoundError(
            f"Aucun fichier de points de contrôle (formats GPKG/SHP) trouvé dans {sources_sig} "
            f"(attendus dans des sous-dossiers points_de_ctrl_OSCEAN_YYYY ou à la racine, "
            f"avec un nom contenant 'ctrl')."
        )

    # Vérification des chemins et fichiers
    for candidate in candidates:
        if not candidate.is_file():
            raise FileNotFoundError(f"Le fichier {candidate} n'est pas un fichier valide.")

    # Sélectionner, pour chaque année, le fichier au suffixe de date le plus récent.
    per_year: dict[str, tuple[str, Path]] = {}
    for p in candidates:
        # Extraire une date au format YYYYMMDD à partir du nom de fichier, même si
        # elle est écrite sous la forme YYYY_MM_DD ou similaire. On concatène
        # tous les chiffres présents et on prend les 8 premiers si possible.
        digits = "".join(ch for ch in p.stem if ch.isdigit())
        if len(digits) >= 8:
            date_str = digits[:8]
            year = date_str[:4]
        elif len(digits) >= 4:
            # Fallback : le nom de fichier ne contient pas de date YYYYMMDD
            # complète (ex. point_ctrl_2024_wgs84 → "202484", 6 chiffres).
            # On tente d'extraire l'année depuis le nom du dossier parent
            # (points_de_ctrl_OSCEAN_YYYY) puis, à défaut, depuis les 4
            # premiers chiffres du nom de fichier.
            parent_digits = re.findall(r"\d{4}", p.parent.name)
            if parent_digits:
                year = parent_digits[-1]
            else:
                year = digits[:4]
            date_str = f"{year}0101"
        else:
            continue
        cur = per_year.get(year)
        if cur is None or date_str > cur[0]:
            per_year[year] = (date_str, p)

    # Restreindre aux années couvrant la période demandée si date_deb/date_fin fournis
    if date_deb is not None and date_fin is not None:
        try:
            deb_ts = pd.to_datetime(date_deb)
            fin_ts = pd.to_datetime(date_fin)
            year_min, year_max = deb_ts.year, fin_ts.year
            filtered = {y: t for y, t in per_year.items() if year_min <= int(y) <= year_max}
            if filtered:
                per_year = filtered
        except Exception as e:
            raise ValueError(f"Erreur de conversion des dates : {e}")

    selected_paths = [tpl[1] for tpl in per_year.values()]
    frames: List[pd.DataFrame] = []
    for path in selected_paths:
        engine = _GPKG_ENGINE
        
        # Filtre à la lecture par département si possible (réduit le volume chargé)
        if dept_code is not None:
            try:
                gdf = gpd.read_file(
                    path,
                    engine=engine,
                    where=f"num_depart = '{str(dept_code).strip()}'",
                )
            except Exception as e:
                raise RuntimeError(f"Erreur de lecture du fichier GPKG avec filtre : {e}")
        else:
            try:
                gdf = gpd.read_file(path, engine=engine)
            except Exception as e:
                raise RuntimeError(f"Erreur de lecture du fichier GPKG : {e}")
        df = pd.DataFrame(gdf.drop(columns=["geometry"], errors="ignore"))
        df.columns = [str(c).split(",")[0].strip() for c in df.columns]
        
        # Validation des colonnes requises
        required_columns = ["date_ctrl", "dc_id", "num_depart"]
        for col in required_columns:
            if col not in df.columns:
                raise KeyError(f"La colonne '{col}' est absente des données point_ctrl_*")
        
        if "date_ctrl" in df.columns:
            df["date_ctrl"] = pd.to_datetime(
                df["date_ctrl"], dayfirst=True, errors="coerce", format="mixed"
            )

        # Alias de colonnes : le reste du code travaille avec des noms
        # uniformisés, malgré les variations entre GPKG/SHP et les versions.
        # - nom_dossier / nom_dossie
        # - type_action / type_actio
        # - Résultat / RESULTAT / resultat
        # - nom_commun / nom_commune
        # - type_usage / type_usager
        # - nature_con / nature_controle
        # - plan_contr / plan_controle
        # - avis_pasbi / avis_patbi / avis_patbiodiv

        # Nom du dossier
        if "nom_dossier" in df.columns and "nom_dossie" not in df.columns:
            df["nom_dossie"] = df["nom_dossier"]

        # Type d'action
        if "type_action" in df.columns and "type_actio" not in df.columns:
            df["type_actio"] = df["type_action"]

        # Variantes du champ résultat
        if "resultat" not in df.columns:
            if "Résultat" in df.columns:
                df["resultat"] = df["Résultat"]
            elif "RESULTAT" in df.columns:
                df["resultat"] = df["RESULTAT"]

        # Nom de commune : harmoniser vers nom_commun
        if "nom_commune" in df.columns and "nom_commun" not in df.columns:
            df["nom_commun"] = df["nom_commune"]

        # Type d'usager / usage : créer les deux alias pour faciliter les usages
        if "type_usager" in df.columns and "type_usage" not in df.columns:
            df["type_usage"] = df["type_usager"]
        if "type_usage" in df.columns and "type_usager" not in df.columns:
            df["type_usager"] = df["type_usage"]

        # Nature du contrôle : nature_con (tronqué) / nature_controle (complet)
        if "nature_controle" in df.columns and "nature_con" not in df.columns:
            df["nature_con"] = df["nature_controle"]
        if "nature_con" in df.columns and "nature_controle" not in df.columns:
            df["nature_controle"] = df["nature_con"]

        # Plan de contrôle : plan_contr / plan_controle
        if "plan_controle" in df.columns and "plan_contr" not in df.columns:
            df["plan_contr"] = df["plan_controle"]
        if "plan_contr" in df.columns and "plan_controle" not in df.columns:
            df["plan_controle"] = df["plan_contr"]

        # Avis patrimoine / biodiversité
        avis_src = None
        for cand in ("avis_patbiodiv", "avis_patbi", "avis_pasbi"):
            if cand in df.columns:
                avis_src = cand
                break
        if avis_src is not None:
            if "avis_patbiodiv" not in df.columns:
                df["avis_patbiodiv"] = df[avis_src]
            if "avis_patbi" not in df.columns:
                df["avis_patbi"] = df[avis_src]
            if "avis_pasbi" not in df.columns:
                df["avis_pasbi"] = df[avis_src]

        frames.append(df)

    if not frames:
        raise FileNotFoundError(
            f"Aucun enregistrement valide trouvé dans les GPKG point_ctrl de {sources_sig}"
        )

    df_all = pd.concat(frames, ignore_index=True)
    dedup_cols = [c for c in ["dc_id", "date_ctrl", "x", "y"] if c in df_all.columns]
    if dedup_cols:
        df_all.drop_duplicates(subset=dedup_cols, keep="first", inplace=True)
    if "date_ctrl" not in df_all.columns:
        raise KeyError("La colonne 'date_ctrl' est absente des données point_ctrl_*")

    # Filtrage optionnel par département et période
    if dept_code is not None and "num_depart" in df_all.columns:
        df_all = df_all[df_all["num_depart"].astype(str).str.strip() == str(dept_code).strip()].copy()
    if date_deb is not None and date_fin is not None:
        try:
            deb_ts = pd.to_datetime(date_deb)
            fin_ts = pd.to_datetime(date_fin)
        except ValueError as e:
            raise ValueError(f"Format de date invalide : {e}")
        df_all = filtre_periode(df_all, "date_ctrl", deb_ts, fin_ts)

    return df_all


def load_pej(
    root: Path,
    dept_code: Optional[str] = None,
    date_deb: Optional[Union[str, pd.Timestamp]] = None,
    date_fin: Optional[Union[str, pd.Timestamp]] = None,
) -> pd.DataFrame:
    """
    Charge le classeur ODS des procédures d'enquête judiciaire le plus récent
    (suivi_procedure_enq_judiciaire_YYYYMMDD.ods dans sources/) et prépare
    la colonne DATE_REF.

    Si date_deb et date_fin sont fournis, filtre les lignes sur cette période
    (réduit le volume en analyse ciblée).
    """
    sources = root / "sources"
    prefix = "suivi_procedure_enq_judiciaire_"
    path = _find_latest_dated_file(sources, prefix, (".ods",))
    df = pd.read_excel(path, dtype=str, engine="odf")
    # Alias pour compatibilité si le classeur utilise "NATINF" au lieu de "NATINF_PEJ"
    if "NATINF" in df.columns and "NATINF_PEJ" not in df.columns:
        df["NATINF_PEJ"] = df["NATINF"]
    df["DATE_CONSTATATION"] = pd.to_datetime(df["DATE_CONSTATATION"], errors="coerce")
    df["DATE_OUVERTURE_PROCEDURE"] = pd.to_datetime(
        df["DATE_OUVERTURE_PROCEDURE"], errors="coerce"
    )
    df["RECAP_DATE_INIT_PJ"] = pd.to_datetime(
        df.get("RECAP_DATE_INIT_PJ", pd.Series(dtype=str)), errors="coerce"
    )
    df["DATE_REF"] = (
        df["DATE_CONSTATATION"]
        .fillna(df["DATE_OUVERTURE_PROCEDURE"])
        .fillna(df["RECAP_DATE_INIT_PJ"])
    )
    if date_deb is not None and date_fin is not None:
        deb_ts = pd.to_datetime(date_deb)
        fin_ts = pd.to_datetime(date_fin)
        df = filtre_periode(df, "DATE_REF", deb_ts, fin_ts)
    return df


def load_pa(
    root: Path,
    dept_code: Optional[str] = None,
    date_deb: Optional[Union[str, pd.Timestamp]] = None,
    date_fin: Optional[Union[str, pd.Timestamp]] = None,
) -> pd.DataFrame:
    """
    Charge le classeur ODS des procédures administratives le plus récent
    (suivi_procedure_administrative_YYYYMMDD.ods dans sources/) et prépare
    la colonne DATE_REF.

    Si date_deb et date_fin sont fournis, filtre les lignes sur cette période.
    """
    sources = root / "sources"
    path = _find_latest_dated_file(
        sources, "suivi_procedure_administrative_", (".ods",)
    )
    df = pd.read_excel(path, dtype=str, engine="odf")
    df["DATE_CONTROLE"] = pd.to_datetime(df["DATE_CONTROLE"], errors="coerce")
    df["DATE_DOSSIER"] = pd.to_datetime(df["DATE_DOSSIER"], errors="coerce")
    df["DATE_REF"] = df["DATE_CONTROLE"].fillna(df["DATE_DOSSIER"])
    if date_deb is not None and date_fin is not None:
        deb_ts = pd.to_datetime(date_deb)
        fin_ts = pd.to_datetime(date_fin)
        df = filtre_periode(df, "DATE_REF", deb_ts, fin_ts)
    return df


def load_pnf(root: Path) -> pd.DataFrame:
    """Charge la liste des communes PNF (référentiel). Cherche dans ref/ puis sources/."""
    for base in ("ref", "sources"):
        path = root / base / "communes_PNF.csv"
        if path.exists():
            df = pd.read_csv(path, sep=",", dtype=str)
            df["CODE_INSEE"] = df["CODE_INSEE"].astype(str).str.zfill(5)
            return df
    raise FileNotFoundError(
        "Aucun fichier communes_PNF.csv trouvé dans ref/ ni sources/."
    )


def load_tub(root: Path) -> pd.DataFrame:
    """Charge la liste des communes TUB (référentiel). Cherche dans ref/ puis sources/."""
    for base in ("ref", "sources"):
        path = root / base / "tub_communes.csv"
        if path.exists():
            df = pd.read_csv(path, sep=";", dtype=str, encoding="latin-1")
            df["INSEE_COM"] = df["INSEE_COM"].astype(str).str.zfill(5)
            return df
    raise FileNotFoundError(
        "Aucun fichier tub_communes.csv trouvé dans ref/ ni sources/."
    )


def load_ref_themes_ctrl(root: Path) -> List[dict]:
    """
    Charge le référentiel des thèmes des contrôles (ref/ref_themes_ctrl.csv).
    Retourne une liste de dictionnaires {"id": ..., "label": ..., "ordre": ...}
    triée par ordre. Si le fichier est absent ou invalide, retourne une liste vide.
    """
    path = root / "ref" / "ref_themes_ctrl.csv"
    if not path.exists():
        return []
    try:
        df = pd.read_csv(path, sep=";", dtype=str, encoding="utf-8")
        if df.empty:
            return []
        # Colonnes attendues : id, label, ordre
        if "id" not in df.columns:
            return []
        out = []
        for _, row in df.iterrows():
            id_val = str(row.get("id", "")).strip()
            if not id_val:
                continue
            label_val = str(row.get("label", id_val)).strip()
            ordre_val = row.get("ordre", "999")
            try:
                ordre_int = int(ordre_val) if ordre_val else 999
            except (ValueError, TypeError):
                ordre_int = 999
            out.append({"id": id_val, "label": label_val, "ordre": ordre_int})
        out.sort(key=lambda x: x["ordre"])
        return out
    except Exception:
        return []


def load_tub_pnf_codes(root: Path) -> Tuple[set, set]:
    """
    Charge les référentiels TUB et PNF et retourne les ensembles de codes INSEE
    (tub_codes, pnf_codes) pour les agrégations par zone. Utile pour réutiliser
    la même logique dans plusieurs bilans.
    """
    tub = load_tub(root)
    pnf = load_pnf(root)
    return set(tub["INSEE_COM"].unique()), set(pnf["CODE_INSEE"].unique())


def load_communes_noms(root: Path) -> dict:
    """
    Charge la table de correspondance code INSEE → nom de commune.

    Source : ref/sig/communes_21/communes.csv (INSEE_COM, NOM_COM).
    Retourne un dictionnaire {code_insee_5chars: nom_commune}.
    Si le fichier est absent, retourne un dict vide (les PDF afficheront le code).
    """
    path = root / "ref" / "sig" / "communes_21" / "communes.csv"
    if not path.exists():
        return {}
    df = pd.read_csv(path, sep=";", dtype=str)
    # Colonnes possibles : première ligne = en-tête type "INSEE_COM,C,5" et "NOM_COM,C,50"
    code_col = next((c for c in df.columns if c.strip().startswith("INSEE_COM")), None)
    nom_col = next((c for c in df.columns if c.strip().startswith("NOM_COM")), None)
    if code_col is None or nom_col is None:
        code_col, nom_col = df.columns[0], df.columns[1]
    df[code_col] = df[code_col].astype(str).str.strip().str.zfill(5)
    return dict(zip(df[code_col], df[nom_col].astype(str).str.strip()))


def load_communes_centroides(root: Path) -> pd.DataFrame:
    """
    Charge la table des communes de France avec centroïdes et renvoie un
    DataFrame minimal (code_insee, lat, lon) pour les jointures.

    Source par défaut : ref/sig/communes-france-2025.csv
    - code_insee : code INSEE commune (5 caractères, zfill(5))
    - lat / lon : coordonnées du centroïde (colonnes latitude_centre / longitude_centre)
    """
    ref_dir = root / "ref" / "sig"
    csv_path = ref_dir / "communes-france-2025.csv"

    if csv_path.exists():
        df = pd.read_csv(csv_path, sep=",", dtype=str)
        # Harmonisation du code INSEE
        insee_col = None
        for cand in ["code_insee", "CODE_INSEE", "insee"]:
            if cand in df.columns:
                insee_col = cand
                break
        if insee_col is None:
            raise KeyError(
                "Aucune colonne de code INSEE trouvée dans communes-france-2025.csv "
                "(attendu: code_insee / CODE_INSEE / insee)"
            )

        lat_col = None
        lon_col = None
        for cand in ["latitude_centre", "LATITUDE_CENTRE", "lat_centre"]:
            if cand in df.columns:
                lat_col = cand
                break
        for cand in ["longitude_centre", "LONGITUDE_CENTRE", "lon_centre"]:
            if cand in df.columns:
                lon_col = cand
                break

        if lat_col is None or lon_col is None:
            raise KeyError(
                "Colonnes de centroïdes manquantes dans communes-france-2025.csv "
                "(attendu: latitude_centre / longitude_centre ou équivalents)"
            )

        out = pd.DataFrame(
            {
                "code_insee": df[insee_col].astype(str).str.zfill(5),
                "lat": pd.to_numeric(df[lat_col], errors="coerce"),
                "lon": pd.to_numeric(df[lon_col], errors="coerce"),
            }
        )
        # On élimine les lignes sans coordonnées valides
        out = out.dropna(subset=["lat", "lon"])
        return out

    # Fallback possible : shapefile / gpkg de communes avec géométrie, si disponible.
    # On extrait alors le centroïde de chaque polygone.
    for base_name in ["communes-france-2025", "communes_france_2025"]:
        for ext in (".gpkg", ".shp"):
            vec_path = ref_dir / f"{base_name}{ext}"
            if vec_path.exists():
                gdf = gpd.read_file(vec_path)
                insee_col = None
                for cand in ["code_insee", "CODE_INSEE", "insee", "INSEE_COM"]:
                    if cand in gdf.columns:
                        insee_col = cand
                        break
                if insee_col is None:
                    raise KeyError(
                        f"Aucune colonne de code INSEE trouvée dans {vec_path} "
                        "(attendu: code_insee / CODE_INSEE / insee / INSEE_COM)"
                    )
                # S'assurer que la géométrie est présente
                if "geometry" not in gdf.columns:
                    raise KeyError(f"Aucune colonne 'geometry' dans {vec_path}")

                # Centroïdes géométriques
                gdf = gdf.set_geometry("geometry")
                centroids = gdf.geometry.centroid
                out = pd.DataFrame(
                    {
                        "code_insee": gdf[insee_col].astype(str).str.zfill(5),
                        "lat": centroids.y,
                        "lon": centroids.x,
                    }
                )
                out = out.dropna(subset=["lat", "lon"])
                return out

    raise FileNotFoundError(
        "Impossible de trouver une table de centroïdes communes : "
        "ni ref/sig/communes-france-2025.csv ni shapefile/GeoPackage équivalent."
    )


def load_natinf_ref(root: Path) -> pd.DataFrame:
    """Charge le référentiel NATINF (ref/liste_natinf.csv) pour libeller les exports."""
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
                # Colonnes spécifiques du fichier standard OFB :
                #   - « Nature de l'infraction » (ex. Délit, Contravention de classe 5)
                #   - « Qualification de l'infraction » (nom de l'infraction)
                nature_col = None
                qualif_col = None
                for c in df.columns:
                    cl = c.lower()
                    if c == "numero_natinf":
                        continue
                    if "nature de l'infraction" in cl:
                        nature_col = c
                    elif "qualification de l'infraction" in cl:
                        qualif_col = c
                # Renommer d'abord les colonnes détaillées (avant tout autre rename)
                if nature_col:
                    df = df.rename(columns={nature_col: "nature_infraction"})
                if qualif_col:
                    df = df.rename(columns={qualif_col: "qualification_infraction"})
                # Fallback libelle_natinf pour les usages existants : privilégier la qualification
                lib_col = None
                if "qualification_infraction" in df.columns:
                    df["libelle_natinf"] = df["qualification_infraction"].fillna("")
                else:
                    for c in df.columns:
                        if c == "numero_natinf":
                            continue
                        if "nature" in c.lower() or "infraction" in c.lower():
                            lib_col = c
                            break
                    if lib_col:
                        df = df.rename(columns={lib_col: "libelle_natinf"})
                if "libelle_natinf" not in df.columns:
                    df["libelle_natinf"] = ""
                out_cols = ["numero_natinf", "libelle_natinf"]
                if "nature_infraction" in df.columns:
                    out_cols.append("nature_infraction")
                if "qualification_infraction" in df.columns:
                    out_cols.append("qualification_infraction")
                return df[out_cols].drop_duplicates()
            except Exception:
                continue
    return pd.DataFrame(columns=["numero_natinf", "libelle_natinf"])


def load_pve(
    root: Path,
    dept_code: Optional[str] = None,
    date_deb: Optional[Union[str, pd.Timestamp]] = None,
    date_fin: Optional[Union[str, pd.Timestamp]] = None,
) -> pd.DataFrame:
    """
    Charge le tableau Stats_PVe_OFB (format .csv ou .ods) et homogénéise
    les principales colonnes utilisées dans les analyses/cartes.

    Le fichier est recherché dynamiquement dans sources/ : on prend le plus récent
    parmi les fichiers dont le nom commence par "Stats_PVe_OFB" (extensions .csv ou .ods),
    en se basant sur la date de modification du fichier.

    Colonnes normalisées :
      * INF-INSEE : chaîne à 5 chiffres (zfill(5)) si présente
      * INF-DEPARTEMENT / INF-DEPART : alias réciproques
      * INF-DATE-INTG : datetime (date d'intégration, jour/mois/année)

    Si dept_code et/ou date_deb/date_fin sont fournis, filtre les lignes en conséquence.
    """
    sources = root / "sources"
    if not sources.exists():
        raise FileNotFoundError(f"Le dossier sources n'existe pas : {sources}")

    candidates: List[Path] = []
    for ext in (".csv", ".ods"):
        candidates.extend(sources.glob(f"Stats_PVe_OFB*{ext}"))
    if not candidates:
        raise FileNotFoundError(
            f"Aucun fichier Stats_PVe_OFB*.csv ou Stats_PVe_OFB*.ods trouvé dans {sources}."
        )
    # Fichier le plus récent (date de modification)
    path = max(candidates, key=lambda p: p.stat().st_mtime)

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, sep=";", dtype=str, encoding="latin1")
    else:
        df = pd.read_excel(path, dtype=str, engine="odf")

    # Normalisation du code commune
    if "INF-INSEE" in df.columns:
        df["INF-INSEE"] = (
            df["INF-INSEE"]
            .astype(str)
            .str.extract(r"(\d{1,5})", expand=False)
            .fillna("")
            .str.zfill(5)
        )

    # Alias de département INF-DEPART / INF-DEPARTEMENT
    if "INF-DEPARTEMENT" in df.columns and "INF-DEPART" not in df.columns:
        df["INF-DEPART"] = df["INF-DEPARTEMENT"]
    elif "INF-DEPART" in df.columns and "INF-DEPARTEMENT" not in df.columns:
        df["INF-DEPARTEMENT"] = df["INF-DEPART"]

    # Date d'intégration
    if "INF-DATE-INTG" in df.columns:
        df["INF-DATE-INTG"] = pd.to_datetime(
            df["INF-DATE-INTG"], dayfirst=True, errors="coerce"
        )

    if dept_code is not None:
        dept_col = "INF-DEPART" if "INF-DEPART" in df.columns else "INF-DEPARTEMENT"
        if dept_col in df.columns:
            df = df[df[dept_col].astype(str).str.strip() == str(dept_code).strip()].copy()
    if date_deb is not None and date_fin is not None and "INF-DATE-INTG" in df.columns:
        deb_ts = pd.to_datetime(date_deb)
        fin_ts = pd.to_datetime(date_fin)
        df = filtre_periode(df, "INF-DATE-INTG", deb_ts, fin_ts)

    return df


def get_points_infrac_pj_path(root: Path) -> Path:
    """
    Chemin du fichier des points infractions PJ (GPKG ou shapefile) le plus récent.

    On recherche en priorité les GeoPackage dans sources/sig/points_infractions_pj/,
    avec un nom du type localisation_infrac_FAITS_YYYYMMDD.gpkg, puis, à défaut,
    un shapefile localisation_infrac_FAITS_YYYYMMDD.shp.
    """
    base_dir = root / "sources" / "sig" / "points_infractions_pj"
    if not base_dir.exists():
        # Compatibilité ancienne arborescence : sources/points_infractions_pj
        base_dir = root / "sources" / "points_infractions_pj"

    try:
        return _find_latest_dated_file(
            base_dir, "localisation_infrac_FAITS_", (".gpkg", ".shp")
        )
    except FileNotFoundError:
        # On laisse l'appelant gérer l'absence de fichier
        return base_dir / "localisation_infrac_FAITS_00000000.gpkg"


def load_points_infrac_pj(
    root: Path, natinf_list: List[str], dept_code: str
) -> gpd.GeoDataFrame:
    """Charge le shapefile/gpkg des points d'infractions PJ, filtre SD + NATINF."""
    path = get_points_infrac_pj_path(root)
    if not path.exists():
        raise FileNotFoundError(path)
    natinf_vals = [int(n) for n in natinf_list]
    # Filtre à la lecture si possible (pyogrio) pour éviter de charger tout le GPKG
    try:
        where_clause = f"entite = 'SD{dept_code}' AND natinf IN ({','.join(map(str, natinf_vals))})"
        gdf = gpd.read_file(path, engine=_GPKG_ENGINE, where=where_clause)
    except Exception:
        gdf = gpd.read_file(path)
        gdf["natinf"] = pd.to_numeric(gdf["natinf"], errors="coerce")
        mask = (gdf["entite"] == f"SD{dept_code}") & (gdf["natinf"].isin(natinf_vals))
        gdf = gdf.loc[mask].copy()

    # Harmonisation du CRS : si absent, on considère que les coordonnées sont en WGS84.
    if gdf.crs is None:
        gdf = gdf.set_crs(epsg=4326)

    if "natinf" in gdf.columns and gdf["natinf"].dtype == object:
        gdf["natinf"] = pd.to_numeric(gdf["natinf"], errors="coerce")

    # Colonne de commune : le schéma réel utilise "commune_fait" (et non "commune_fa").
    commune_col = None
    if "commune_fait" in gdf.columns:
        commune_col = "commune_fait"
    elif "commune_fa" in gdf.columns:
        commune_col = "commune_fa"

    cols = ["dossier", "natinf", "x_infrac", "y_infrac", "geometry"]
    if commune_col is not None:
        cols.insert(2, commune_col)
    cols = [c for c in cols if c in gdf.columns]
    return gdf[cols].copy()


def load_pj_with_geometry(
    root: Path,
    natinf_list: List[str],
    dept_code: str,
    date_deb: Optional[Union[str, pd.Timestamp]] = None,
    date_fin: Optional[Union[str, pd.Timestamp]] = None,
    pej_df: Optional[pd.DataFrame] = None,
) -> gpd.GeoDataFrame:
    """
    Charge les procédures judiciaires (ODS) et les associe à la géométrie des faits
    (GPKG points PJ). Ne conserve que les dossiers présents dans le GPKG (SD + NATINF).

    Retourne un GeoDataFrame avec les colonnes PJ utiles + geometry.
    Si date_deb/date_fin sont fournis, le PEJ est filtré sur cette période avant jointure.
    Si pej_df est fourni, il est utilisé à la place de recharger l'ODS (évite double lecture).
    """
    if pej_df is not None:
        pej = pej_df
    else:
        pej = load_pej(root, date_deb=date_deb, date_fin=date_fin)
    pts_pj = load_points_infrac_pj(root, natinf_list, dept_code)
    dossiers_geom = set(pts_pj["dossier"].unique())
    pej_with_geom = pej[pej["DC_ID"].isin(dossiers_geom)].copy()
    if pej_with_geom.empty:
        return gpd.GeoDataFrame(columns=list(pej.columns) + ["geometry"], crs=pts_pj.crs)

    pts_dedup = pts_pj.drop_duplicates(subset="dossier", keep="first")
    merged = pej_with_geom.merge(
        pts_dedup[["dossier", "geometry"]],
        left_on="DC_ID",
        right_on="dossier",
        how="left",
    )
    if "dossier" in merged.columns:
        merged = merged.drop(columns=["dossier"])
    return gpd.GeoDataFrame(merged, geometry="geometry", crs=pts_pj.crs)
