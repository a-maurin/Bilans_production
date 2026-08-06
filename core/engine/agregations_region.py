# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, mais SANS AUCUNE GARANTIE ;
# sans même la garantie implicite de QUALITÉ MARCHANDE ou D'ADÉQUATION À UN USAGE PARTICULIER.
# Voir la Licence Publique Générale GNU pour plus de détails.
#
# CONDITIONS SUPPLÉMENTAIRES D'ATTRIBUTION (SECTION 7(b) DE LA GPL v3) :
# Conformément à la section 7(b) de la GNU GPL v3, vous devez expressément conserver
# intactes et lisibles toutes les mentions d'auteur, notices de copyright et la présente
# clause dans chaque fichier source ou interface utilisateur redistribué. Toute version modifiée
# doit clairement indiquer qu'elle a été altérée et ne doit en aucun cas supprimer le nom
# de l'auteur original (Aguirre MAURIN).

"""
========================================================================================
MODULE : CALCULS D'AGREGATION A L'ECHELLE REGIONALE (`agregations_region.py`)
========================================================================================
Ce module gère le regroupement et la ventilation des données de police au niveau régional
et interdépartemental (Direction Régionale, Brigade de Mission Interdépartementale BMI).

Rôles :
  1. Association des NATINF PVe aux thèmes de profil via le registre YAML.
  2. Ventilation interdépartementale des contrôles, opérations, PA, PEJ et PVe.
  3. Génération des matrices de répartition départementale pour la cartographie et les
     tableaux comparatifs régionaux dans les PDF.
========================================================================================
"""
import pandas as pd
from pathlib import Path
from typing import Any
from core.common.utilitaires_metier import get_departements_pour_perimetre

def _load_natinf_to_theme_map() -> dict[str, str]:
    """Lit les profils YAML pour associer chaque NATINF PVe à un id de profil (thème)."""
    from core.chemins_projet import PROJECT_ROOT
    import yaml
    
    mapping = {}
    profiles_dir = PROJECT_ROOT / "config" / "profils_bilan"
    if not profiles_dir.exists():
        return mapping
    
    for p in profiles_dir.glob("*.yaml"):
        if p.stem in ("_defaults", "schema_ui"):
            continue
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f)
                if isinstance(data, dict):
                    natinfs = data.get("natinf_pve", [])
                    if isinstance(natinfs, list):
                        for n in natinfs:
                            mapping[str(n).strip()] = p.stem
        except Exception:
            continue
    return mapping


def analyse_region_par_departement(
    point: pd.DataFrame,
    pa: pd.DataFrame,
    pej: pd.DataFrame,
    pve: pd.DataFrame,
    echelle: str,
    code: str,
    out_dir: Path,
    pej_global: pd.DataFrame | None = None,
    profil_id: str = "global",
) -> None:
    if str(echelle).strip().lower() not in ("region", "bmi") and str(profil_id).strip().lower() not in ("pnf_v2", "pnf"):
        return
        
    if str(profil_id).strip().lower() in ("pnf_v2", "pnf"):
        dept_codes = ["21", "52"]
    else:
        dept_codes = get_departements_pour_perimetre(echelle, code)
        if not dept_codes or "FR" in dept_codes:
            return
        
    rows = []
    
    # 1. Traitement des points de contrôle (Localisations et Opérations)
    if not point.empty:
        # Assurer qu'on a un num_depart propre et filtré
        pt = point.copy()
        if "num_depart" not in pt.columns:
            pt["num_depart"] = "Inconnu"
        pt["num_depart"] = pt["num_depart"].astype(str).str.strip().str.split('.').str[0].str.zfill(2)
        pt = pt[pt["num_depart"].isin(dept_codes)]
        
        pt["domaine"] = pt["domaine"].fillna("Hors domaine").astype(str) if "domaine" in pt.columns else "Hors domaine"
        pt["theme"] = pt["theme"].fillna("Hors thème").astype(str) if "theme" in pt.columns else (pt["thematique"].fillna("Hors thème").astype(str) if "thematique" in pt.columns else "Hors thème")
        
        if not pt.empty:
            # Localisations
            locs = pt.groupby(["domaine", "theme", "num_depart"]).size().reset_index(name="nb_localisations")
            
            # Opérations (fc_id uniques)
            if "fc_id" in pt.columns:
                ops = pt.groupby(["domaine", "theme", "num_depart"])["fc_id"].nunique().reset_index(name="nb_operations")
                locs = pd.merge(locs, ops, on=["domaine", "theme", "num_depart"], how="outer")
            else:
                locs["nb_operations"] = 0
                
            for _, r in locs.iterrows():
                rows.append({
                    "domaine": r["domaine"],
                    "theme": r["theme"],
                    "departement": r["num_depart"],
                    "metrique": "nb_localisations",
                    "valeur": r["nb_localisations"]
                })
                rows.append({
                    "domaine": r["domaine"],
                    "theme": r["theme"],
                    "departement": r["num_depart"],
                    "metrique": "nb_operations",
                    "valeur": r["nb_operations"]
                })
            
    # 2. PEJ
    if not pej.empty:
        pj = pej.copy()
        pj["domaine"] = pj["DOMAINE"].fillna("Hors domaine").astype(str) if "DOMAINE" in pj.columns else "Hors domaine"
        pj["theme"] = pj["THEME"].fillna("Hors thème").astype(str) if "THEME" in pj.columns else "Hors thème"
        pj["departement"] = "Inconnu"
        if "ENTITE_ORIGINE_PROCEDURE" in pj.columns:
            # Extraction du département de SDXX
            pj["departement"] = pj["ENTITE_ORIGINE_PROCEDURE"].astype(str).str.extract(r'(\d+)')[0]
            pj["departement"] = pj["departement"].fillna("Inconnu").astype(str).str.strip().str.zfill(2)
        pj = pj[pj["departement"].isin(dept_codes)]
            
        if "DATE_REF" in pj.columns and "DC_ID" in pj.columns:
            pj = pj.sort_values("DATE_REF", ascending=False).drop_duplicates("DC_ID")
            
        if not pj.empty:
            pejs = pj.groupby(["domaine", "theme", "departement"]).size().reset_index(name="nb_pej")
            for _, r in pejs.iterrows():
                rows.append({
                    "domaine": r["domaine"],
                    "theme": r["theme"],
                    "departement": r["departement"],
                    "metrique": "nb_pej",
                    "valeur": r["nb_pej"]
                })
            
    # 3. PA
    if not point.empty and "resultat" in point.columns:
        from core.common.utilitaires_metier import filter_points_induisant_pa
        pt_pa = filter_points_induisant_pa(point)
        if not pt_pa.empty:
            pt_pa = pt_pa.copy()
            if "num_depart" not in pt_pa.columns:
                pt_pa["num_depart"] = "Inconnu"
            pt_pa["num_depart"] = pt_pa["num_depart"].astype(str).str.strip().str.split('.').str[0].str.zfill(2)
            pt_pa = pt_pa[pt_pa["num_depart"].isin(dept_codes)]
            
            if not pt_pa.empty:
                pt_pa["domaine"] = pt_pa["domaine"].fillna("Hors domaine").astype(str) if "domaine" in pt_pa.columns else "Hors domaine"
                pt_pa["theme"] = pt_pa["theme"].fillna("Hors thème").astype(str) if "theme" in pt_pa.columns else (pt_pa["thematique"].fillna("Hors thème").astype(str) if "thematique" in pt_pa.columns else "Hors thème")
                pas = pt_pa.groupby(["domaine", "theme", "num_depart"]).size().reset_index(name="nb_pa")
                for _, r in pas.iterrows():
                    rows.append({
                        "domaine": r["domaine"],
                        "theme": r["theme"],
                        "departement": r["num_depart"],
                        "metrique": "nb_pa",
                        "valeur": r["nb_pa"]
                    })
                
    # 4. PVe
    if not pve.empty:
        pv = pve.copy()
        pv["domaine"] = pv["DOMAINE"].fillna("Hors domaine").astype(str) if "DOMAINE" in pv.columns else "Hors domaine"
        
        natinf_map = _load_natinf_to_theme_map()
        def _get_theme_from_natinf(val):
            if pd.isna(val):
                return "Hors thème"
            tokens = [t.strip() for t in str(val).replace("_", " ").replace("-", " ").split() if t.strip()]
            for tok in tokens:
                if tok in natinf_map:
                    return natinf_map[tok]
            return "Hors thème"
            
        natinf_col = "INF-NATINF" if "INF-NATINF" in pv.columns else ("NATINF" if "NATINF" in pv.columns else None)
        if natinf_col:
            pv["theme"] = pv[natinf_col].apply(_get_theme_from_natinf)
        else:
            pv["theme"] = "Hors thème"
            
        pv["departement"] = "Inconnu"
        if "INF-INSEE" in pv.columns:
            s_insee = pv["INF-INSEE"].astype(str).str.strip().str.zfill(5)
            pv["departement"] = s_insee.where(~s_insee.str.startswith("97"), s_insee.str[:3])
            pv["departement"] = pv["departement"].where(s_insee.str.startswith("97"), s_insee.str[:2])
        elif "INSEE_DEP" in pv.columns:
            pv["departement"] = pv["INSEE_DEP"].astype(str)
            
        pv["departement"] = pv["departement"].astype(str).str.strip().str.zfill(2)
        pv = pv[pv["departement"].isin(dept_codes)]

        if not pv.empty:
            pves = pv.groupby(["domaine", "theme", "departement"]).size().reset_index(name="nb_pve")
            for _, r in pves.iterrows():
                rows.append({
                    "domaine": r["domaine"],
                    "theme": r["theme"],
                    "departement": r["departement"],
                    "metrique": "nb_pve",
                    "valeur": r["nb_pve"]
                })

    if not rows:
        pd.DataFrame(columns=["domaine", "theme", "departement", "metrique", "valeur"]).to_csv(out_dir / "region_detail_par_dept.csv", sep=";", index=False)
        return
        
    df = pd.DataFrame(rows)
    # Pivot
    df_pivot = df.pivot_table(index=["domaine", "theme", "departement"], columns="metrique", values="valeur", aggfunc="sum").fillna(0).reset_index()
    
    # Ensure all columns exist
    for col in ["nb_operations", "nb_localisations", "nb_pej", "nb_pa", "nb_pve"]:
        if col not in df_pivot.columns:
            df_pivot[col] = 0
            
    df_pivot.to_csv(out_dir / "region_detail_par_dept.csv", sep=";", index=False)

    if str(profil_id).strip().lower() in ("pnf_v2", "pnf"):
        generer_csv_pnf_coeur_vs_aoa(point, pej, pve, out_dir)

    if pej_global is not None and not pej_global.empty:
        calculer_ratio_pej_departement(pej, pej_global, echelle, code, out_dir, profil_id)


def _extract_gdf(df: pd.DataFrame, out_dir: Path, pattern: str = "controles_*.gpkg") -> Any:
    """Helper pour convertir un DataFrame en GeoDataFrame ou charger le GPKG généré."""
    try:
        import geopandas as gpd
    except ImportError:
        return None

    if df is None or df.empty:
        return None

    # 1. Si c'est déjà un GeoDataFrame avec géométrie valide
    if isinstance(df, gpd.GeoDataFrame) and "geometry" in df.columns and df.geometry.notna().any():
        return df

    # 2. Vérification des colonnes de géométrie explicites ou coordonnées
    if "geometry" in df.columns:
        try:
            return gpd.GeoDataFrame(df.copy(), geometry="geometry")
        except Exception:
            pass

    x_col = next((c for c in ["x", "X", "lon", "longitude", "x_faits", "x_l93", "x_2154", "inf_gps_long"] if c in df.columns), None)
    y_col = next((c for c in ["y", "Y", "lat", "latitude", "y_faits", "y_l93", "y_2154", "inf_gps_lat"] if c in df.columns), None)
    if x_col and y_col:
        try:
            s_x = pd.to_numeric(df[x_col], errors="coerce")
            s_y = pd.to_numeric(df[y_col], errors="coerce")
            valid = s_x.notna() & s_y.notna() & (s_x != 0) & (s_y != 0)
            if valid.any():
                crs = "EPSG:2154" if float(s_x.dropna().iloc[0]) > 1000 else "EPSG:4326"
                return gpd.GeoDataFrame(df[valid].copy(), geometry=gpd.points_from_xy(s_x[valid], s_y[valid]), crs=crs)
        except Exception:
            pass

    # 3. Fallback : Chargement du fichier GPKG de sortie s'il existe
    gpkg_files = list(out_dir.glob(pattern)) + list(out_dir.rglob(pattern))
    from core.chemins_projet import PROJECT_ROOT
    carto_dir = PROJECT_ROOT / "data" / "sources" / "sig" / "CARTO"
    gpkg_files += list(carto_dir.glob(pattern))
    if gpkg_files:
        try:
            gpkg_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            gdf_gpkg = gpd.read_file(gpkg_files[0])
            if gdf_gpkg is not None and not gdf_gpkg.empty and "geometry" in gdf_gpkg.columns:
                return gdf_gpkg
        except Exception:
            pass

    return None


def generer_csv_pnf_coeur_vs_aoa(
    point: pd.DataFrame,
    pej: pd.DataFrame,
    pve: pd.DataFrame,
    out_dir: Path,
) -> None:
    """
    Génère le fichier CSV `pnf_coeur_vs_aoa.csv` ventilé par territoire
    ('Cœur de parc', 'AOA (Hors cœur)', 'Non géolocalisé (AOA)') et par département (21, 52).
    """
    try:
        from core.common.chargeurs_donnees import load_pnf_coeur_gdf
        from core.chemins_projet import PROJECT_ROOT
        try:
            import geopandas as gpd
        except ImportError:
            gpd = None

        gdf_coeur = load_pnf_coeur_gdf(PROJECT_ROOT) if gpd is not None else None
        if gdf_coeur is not None and not gdf_coeur.empty and gdf_coeur.crs is None:
            gdf_coeur = gdf_coeur.set_crs("EPSG:2154")

        dept_codes = ["21", "52"]
        records = []

        # 1. Point (Opérations & Localisations)
        if not point.empty:
            pt = point.copy()
            if "num_depart" not in pt.columns:
                pt["num_depart"] = "Inconnu"
            pt["num_depart"] = pt["num_depart"].astype(str).str.strip().str.split('.').str[0].str.zfill(2)
            pt = pt[pt["num_depart"].isin(dept_codes)].copy()

            if not pt.empty:
                pt["zonage"] = "Non géolocalisé (AOA)"
                if gpd is not None and gdf_coeur is not None and not gdf_coeur.empty:
                    try:
                        gdf_pts = _extract_gdf(pt, out_dir, pattern="controles_*.gpkg")
                        if gdf_pts is not None and not gdf_pts.empty:
                            if gdf_pts.crs is None or gdf_pts.crs.to_epsg() != 2154:
                                if gdf_pts.crs is not None:
                                    gdf_pts = gdf_pts.to_crs("EPSG:2154")
                                else:
                                    gdf_pts = gdf_pts.set_crs("EPSG:2154")
                            sj = gpd.sjoin(gdf_pts, gdf_coeur, predicate="within", how="left")
                            is_in_coeur = sj["index_right"].notna()

                            # If length matches pt:
                            if len(is_in_coeur) == len(pt):
                                pt["zonage"] = "AOA (Hors cœur)"
                                pt.loc[is_in_coeur.values, "zonage"] = "Cœur de parc"
                            else:
                                # Join on fc_id if present
                                if "fc_id" in sj.columns:
                                    coeur_fc_ids = set(sj[sj["index_right"].notna()]["fc_id"].dropna())
                                    aoa_fc_ids = set(sj["fc_id"].dropna())
                                    pt["zonage"] = pt["fc_id"].apply(lambda fid: "Cœur de parc" if fid in coeur_fc_ids else ("AOA (Hors cœur)" if fid in aoa_fc_ids else "Non géolocalisé (AOA)"))
                                else:
                                    pt["zonage"] = "AOA (Hors cœur)"
                                    pt.loc[is_in_coeur, "zonage"] = "Cœur de parc"
                    except Exception:
                        pass

                # Aggregation Localisations & Operations
                for (zonage, dept), sub in pt.groupby(["zonage", "num_depart"]):
                    nb_locs = len(sub)
                    nb_ops = sub["fc_id"].nunique() if "fc_id" in sub.columns else nb_locs
                    records.append({"zonage": zonage, "departement": dept, "metrique": "nb_localisations", "valeur": nb_locs})
                    records.append({"zonage": zonage, "departement": dept, "metrique": "nb_operations", "valeur": nb_ops})

        # 2. PEJ
        if not pej.empty:
            pj = pej.copy()
            pj["departement"] = "Inconnu"
            if "ENTITE_ORIGINE_PROCEDURE" in pj.columns:
                pj["departement"] = pj["ENTITE_ORIGINE_PROCEDURE"].astype(str).str.extract(r'(\d+)')[0]
                pj["departement"] = pj["departement"].fillna("Inconnu").astype(str).str.strip().str.zfill(2)
            pj = pj[pj["departement"].isin(dept_codes)].copy()

            if "DATE_REF" in pj.columns and "DC_ID" in pj.columns:
                pj = pj.sort_values("DATE_REF", ascending=False).drop_duplicates("DC_ID")

            if not pj.empty:
                pj["zonage"] = "AOA (Hors cœur)"
                if gpd is not None and gdf_coeur is not None and not gdf_coeur.empty:
                    try:
                        from core.common.chargeurs_donnees import merge_pej_faits_locations, get_communes_centroids_dicts
                        pj_merged = merge_pej_faits_locations(pj, project_root, echelle, code)
                        dict_x, dict_y = get_communes_centroids_dicts(project_root)
                        
                        lon_col = next((c for c in ["x_faits", "x_infrac", "x_longitude", "longitude", "lon"] if c in pj_merged.columns), None)
                        lat_col = next((c for c in ["y_faits", "y_infrac", "y_latitude", "latitude", "lat"] if c in pj_merged.columns), None)
                        
                        df_geo = pj_merged.copy()
                        df_geo["_lon"] = pd.to_numeric(df_geo[lon_col], errors="coerce") if lon_col else np.nan
                        df_geo["_lat"] = pd.to_numeric(df_geo[lat_col], errors="coerce") if lat_col else np.nan
                        
                        missing = df_geo["_lon"].isna() | df_geo["_lat"].isna() | (df_geo["_lon"] == 0) | (df_geo["_lat"] == 0)
                        if missing.any() and "INSEE_COM" in df_geo.columns:
                            s_insee = df_geo["INSEE_COM"].astype(str).str.zfill(5)
                            df_geo.loc[missing, "_lon"] = s_insee.map(dict_x)
                            df_geo.loc[missing, "_lat"] = s_insee.map(dict_y)
                            
                        valid_geo = df_geo["_lon"].notna() & df_geo["_lat"].notna() & (df_geo["_lon"] != 0)
                        if valid_geo.any():
                            gdf_pj = gpd.GeoDataFrame(
                                df_geo[valid_geo],
                                geometry=gpd.points_from_xy(df_geo.loc[valid_geo, "_lon"], df_geo.loc[valid_geo, "_lat"]),
                                crs="EPSG:4326"
                            ).to_crs("EPSG:2154")
                            sj = gpd.sjoin(gdf_pj, gdf_coeur, predicate="within", how="left")
                            is_in = sj["index_right"].notna()
                            pj_indices = pj.index[valid_geo.values]
                            pj.loc[pj_indices[is_in.values], "zonage"] = "Cœur de parc"
                    except Exception as e_pj:
                        logger.warning(f"Spatial join PEJ Cœur vs AOA : {e_pj}")

                for (zonage, dept), sub in pj.groupby(["zonage", "departement"]):
                    records.append({"zonage": zonage, "departement": dept, "metrique": "nb_pej", "valeur": len(sub)})

        # 3. PA
        if not point.empty and "resultat" in point.columns:
            from core.common.utilitaires_metier import filter_points_induisant_pa
            pt_pa = filter_points_induisant_pa(point)
            if not pt_pa.empty:
                pt_pa = pt_pa.copy()
                if "num_depart" not in pt_pa.columns:
                    pt_pa["num_depart"] = "Inconnu"
                pt_pa["num_depart"] = pt_pa["num_depart"].astype(str).str.strip().str.split('.').str[0].str.zfill(2)
                pt_pa = pt_pa[pt_pa["num_depart"].isin(dept_codes)].copy()

                if not pt_pa.empty:
                    pt_pa["zonage"] = "AOA (Hors cœur)"
                    if gpd is not None and gdf_coeur is not None and not gdf_coeur.empty and not gdf_pts.empty:
                        try:
                            pts_pa_gdf = gdf_pts[gdf_pts.index.isin(pt_pa.index)] if hasattr(gdf_pts, "index") else None
                            if pts_pa_gdf is not None and not pts_pa_gdf.empty:
                                if pts_pa_gdf.crs is None or pts_pa_gdf.crs.to_epsg() != 2154:
                                    pts_pa_gdf = pts_pa_gdf.to_crs("EPSG:2154")
                                sj = gpd.sjoin(pts_pa_gdf, gdf_coeur, predicate="within", how="left")
                                is_in = sj["index_right"].notna()
                                pa_in_indices = pts_pa_gdf.index[is_in.values]
                                pt_pa.loc[pt_pa.index.isin(pa_in_indices), "zonage"] = "Cœur de parc"
                        except Exception as e_pa:
                            logger.warning(f"Spatial join PA Cœur vs AOA : {e_pa}")

                    for (zonage, dept), sub in pt_pa.groupby(["zonage", "num_depart"]):
                        records.append({"zonage": zonage, "departement": dept, "metrique": "nb_pa", "valeur": len(sub)})

        # 4. PVe
        if not pve.empty:
            pv = pve.copy()
            pv["departement"] = "Inconnu"
            if "INF-INSEE" in pv.columns:
                s_insee = pv["INF-INSEE"].astype(str).str.strip().str.zfill(5)
                pv["departement"] = s_insee.where(~s_insee.str.startswith("97"), s_insee.str[:3])
                pv["departement"] = pv["departement"].where(s_insee.str.startswith("97"), s_insee.str[:2])
            elif "INSEE_DEP" in pv.columns:
                pv["departement"] = pv["INSEE_DEP"].astype(str)
            pv["departement"] = pv["departement"].astype(str).str.strip().str.zfill(2)
            pv = pv[pv["departement"].isin(dept_codes)].copy()

            if not pv.empty:
                pv["zonage"] = "AOA (Hors cœur)"
                if gpd is not None and gdf_coeur is not None and not gdf_coeur.empty:
                    try:
                        from core.common.chargeurs_donnees import get_communes_centroids_dicts
                        dict_x, dict_y = get_communes_centroids_dicts(project_root)
                        s_insee = pv["INF-INSEE"].astype(str).str.strip().str.zfill(5) if "INF-INSEE" in pv.columns else pd.Series(dtype=str)
                        pv["_lon"] = s_insee.map(dict_x)
                        pv["_lat"] = s_insee.map(dict_y)
                        
                        valid_pv = pv["_lon"].notna() & pv["_lat"].notna()
                        if valid_pv.any():
                            gdf_pv = gpd.GeoDataFrame(
                                pv[valid_pv],
                                geometry=gpd.points_from_xy(pv.loc[valid_pv, "_lon"], pv.loc[valid_pv, "_lat"]),
                                crs="EPSG:4326"
                            ).to_crs("EPSG:2154")
                            sj = gpd.sjoin(gdf_pv, gdf_coeur, predicate="within", how="left")
                            is_in = sj["index_right"].notna()
                            pv_indices = pv.index[valid_pv.values]
                            pv.loc[pv_indices[is_in.values], "zonage"] = "Cœur de parc"
                    except Exception as e_pv:
                        logger.warning(f"Spatial join PVe Cœur vs AOA : {e_pv}")

                for (zonage, dept), sub in pv.groupby(["zonage", "departement"]):
                    records.append({"zonage": zonage, "departement": dept, "metrique": "nb_pve", "valeur": len(sub)})

        df_res = pd.DataFrame(records)
        if df_res.empty:
            pd.DataFrame(columns=["zonage", "departement", "nb_operations", "nb_localisations", "nb_pej", "nb_pa", "nb_pve"]).to_csv(out_dir / "pnf_coeur_vs_aoa.csv", sep=";", index=False)
            return

        df_piv = df_res.pivot_table(index=["zonage", "departement"], columns="metrique", values="valeur", aggfunc="sum").fillna(0).reset_index()
        for col in ["nb_operations", "nb_localisations", "nb_pej", "nb_pa", "nb_pve"]:
            if col not in df_piv.columns:
                df_piv[col] = 0

        df_piv.to_csv(out_dir / "pnf_coeur_vs_aoa.csv", sep=";", index=False)
    except Exception:
        pd.DataFrame(columns=["zonage", "departement", "nb_operations", "nb_localisations", "nb_pej", "nb_pa", "nb_pve"]).to_csv(out_dir / "pnf_coeur_vs_aoa.csv", sep=";", index=False)


def calculer_ratio_pej_departement(
    pej_filtered: pd.DataFrame,
    pej_global: pd.DataFrame,
    echelle: str,
    code: str,
    out_dir: Path,
    profil_id: str,
) -> None:
    """Calcule le ratio d'enquêtes thématiques par rapport au global par département."""
    if str(echelle).strip().lower() not in ("region", "bmi"):
        return
        
    def _prepare_pej(df: pd.DataFrame) -> pd.DataFrame:
        d = df.copy()
        d["departement"] = "Inconnu"
        if "ENTITE_ORIGINE_PROCEDURE" in d.columns:
            d["departement"] = d["ENTITE_ORIGINE_PROCEDURE"].astype(str).str.extract(r'(\d+)')[0]
            d["departement"] = d["departement"].fillna("Inconnu")
        if "DATE_REF" in d.columns and "DC_ID" in d.columns:
            d = d.sort_values("DATE_REF", ascending=False).drop_duplicates("DC_ID")
        return d

    df_glob = _prepare_pej(pej_global)
    df_filt = _prepare_pej(pej_filtered)
    
    glob_counts = df_glob.groupby("departement").size().reset_index(name="total_pej")
    filt_counts = df_filt.groupby("departement").size().reset_index(name="ppp_pej")
    
    merged = pd.merge(glob_counts, filt_counts, on="departement", how="outer").fillna(0)
    merged["total_pej"] = merged["total_pej"].astype(int)
    merged["ppp_pej"] = merged["ppp_pej"].astype(int)
    
    merged = merged[merged["departement"] != "Inconnu"]
    
    merged["ratio_pourcent"] = 0.0
    mask = merged["total_pej"] > 0
    merged.loc[mask, "ratio_pourcent"] = (merged.loc[mask, "ppp_pej"] / merged.loc[mask, "total_pej"]) * 100.0
    merged["ratio_pourcent"] = merged["ratio_pourcent"].round(1)
    
    merged = merged.sort_values("departement")
    
    csv_name = f"{profil_id}_ratio_pej_departement.csv"
    merged.to_csv(out_dir / csv_name, sep=";", index=False)

