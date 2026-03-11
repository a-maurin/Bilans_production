"""Chargement des données OSCEAN (points de contrôle, PEJ, PA, PVe, PNF, TUB, infractions PJ)."""
from pathlib import Path
from typing import List, Optional, Tuple, Union

import geopandas as gpd
import pandas as pd
import re


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

    candidates = list(sources_sig.glob("point_ctrl_*_wgs84.gpkg"))
    if not candidates:
        raise FileNotFoundError(
            f"Aucun fichier point_ctrl_YYYYMMDD_wgs84.gpkg trouvé dans {sources_sig}"
        )

    # Sélectionner, pour chaque année, le fichier au suffixe de date le plus récent.
    per_year: dict[str, tuple[str, Path]] = {}
    for p in candidates:
        m = re.match(r"point_ctrl_(\d{8})_wgs84\.gpkg", p.name)
        if not m:
            continue
        date_str = m.group(1)
        year = date_str[:4]
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
        except Exception:
            pass

    selected_paths = [tpl[1] for tpl in per_year.values()]
    frames: List[pd.DataFrame] = []
    for path in selected_paths:
        # Filtre à la lecture par département si possible (réduit le volume chargé)
        if dept_code is not None:
            try:
                gdf = gpd.read_file(
                    path,
                    engine="pyogrio",
                    where=f"num_depart = '{str(dept_code).strip()}'",
                )
            except Exception:
                gdf = gpd.read_file(path)
        else:
            gdf = gpd.read_file(path)
        df = pd.DataFrame(gdf.drop(columns=["geometry"], errors="ignore"))
        df.columns = [str(c).split(",")[0].strip() for c in df.columns]
        if "date_ctrl" in df.columns:
            df["date_ctrl"] = pd.to_datetime(
                df["date_ctrl"], dayfirst=True, errors="coerce", format="mixed"
            )
        # Alias de colonnes : le script attend nom_dossie / type_actio (noms possibles dans GPKG : nom_dossier / type_action)
        if "nom_dossier" in df.columns and "nom_dossie" not in df.columns:
            df["nom_dossie"] = df["nom_dossier"]
        if "type_action" in df.columns and "type_actio" not in df.columns:
            df["type_actio"] = df["type_action"]
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
        deb_ts = pd.to_datetime(date_deb)
        fin_ts = pd.to_datetime(date_fin)
        df_all = df_all[(df_all["date_ctrl"] >= deb_ts) & (df_all["date_ctrl"] <= fin_ts)].copy()

    return df_all


def load_pej(
    root: Path,
    dept_code: Optional[str] = None,
    date_deb: Optional[Union[str, pd.Timestamp]] = None,
    date_fin: Optional[Union[str, pd.Timestamp]] = None,
) -> pd.DataFrame:
    """
    Charge le classeur ODS des procédures judiciaires le plus récent
    (suivi_procedure_judiciaire_YYYYMMDD.ods ou
    suivi_procedure_enq_judiciaire_YYYYMMDD.ods dans sources/) et prépare
    la colonne DATE_REF.

    Si date_deb et date_fin sont fournis, filtre les lignes sur cette période
    (réduit le volume en analyse ciblée).
    """
    sources = root / "sources"
    # On tolère les deux conventions de nommage : avec ou sans "enq"
    prefixes = (
        "suivi_procedure_judiciaire_",
        "suivi_procedure_enq_judiciaire_",
    )
    path: Path | None = None
    last_err: Exception | None = None
    for prefix in prefixes:
        try:
            path = _find_latest_dated_file(sources, prefix, (".ods",))
            break
        except FileNotFoundError as e:
            last_err = e
            continue
    if path is None:
        # On remonte la dernière erreur rencontrée (message explicite)
        if last_err is not None:
            raise last_err
        raise FileNotFoundError(
            f"Aucun fichier suivi_procedure*_judiciaire_YYYYMMDD.ods trouvé dans {sources}."
        )
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
        df = df[(df["DATE_REF"] >= deb_ts) & (df["DATE_REF"] <= fin_ts)].copy()
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
        df = df[(df["DATE_REF"] >= deb_ts) & (df["DATE_REF"] <= fin_ts)].copy()
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


def load_tub_pnf_codes(root: Path) -> Tuple[set, set]:
    """
    Charge les référentiels TUB et PNF et retourne les ensembles de codes INSEE
    (tub_codes, pnf_codes) pour les agrégations par zone. Utile pour réutiliser
    la même logique dans plusieurs bilans.
    """
    tub = load_tub(root)
    pnf = load_pnf(root)
    return set(tub["INSEE_COM"].unique()), set(pnf["CODE_INSEE"].unique())


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
        df = pd.read_csv(path, sep=";", dtype=str)
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
        df = df[(df["INF-DATE-INTG"] >= deb_ts) & (df["INF-DATE-INTG"] <= fin_ts)].copy()

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
        gdf = gpd.read_file(path, engine="pyogrio", where=where_clause)
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
