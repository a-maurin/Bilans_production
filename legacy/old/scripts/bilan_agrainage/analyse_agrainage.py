"""Bilan agrainage — contrôles OSCEAN, PVe, PEJ."""
import argparse
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import List

import geopandas as gpd
import pandas as pd
from PIL import Image as PILImage
from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image as RLImage,
    KeepTogether,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
)

# Bootstrap : exécution indépendante
_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from scripts.paths import get_cartes_dir, get_out_dir, get_sig_dir, get_sources_sig_dir
from scripts.common.loaders import (
    load_communes_centroides,
    load_pa,
    load_pej,
    load_point_ctrl,
    load_pj_with_geometry,
    load_pnf,
    load_pve,
    load_points_infrac_pj,
    load_tub,
    load_natinf_ref,
)
from scripts.common.utils import (
    contient_natinf,
    _zone_count,
    _zone_summary,
    _detect_insee_column,
)
from scripts.common.prompt_periode import ask_periode_dept
from scripts.common.ofb_charte import (
    COLOR_GREY,
    COLOR_PRIMARY,
    COLOR_SECONDARY,
    FONT_FAMILY,
    IMG_BACKGROUND,
    IMG_FOOTER_DECO,
    IMG_LOGO_BANNER,
    MARGIN_BOTTOM,
    MARGIN_LEFT,
    MARGIN_RIGHT,
    MARGIN_TOP,
    PAGE_H,
    PAGE_W,
    Spinner,
    _get_styles,
)
from scripts.common.pdf_utils import key_figures_table, ofb_table
from scripts.common.charts import chart_bar, chart_bar_grouped, chart_pie

# ---------------------------------------------------------------------------
# Période et paramètres
# ---------------------------------------------------------------------------
DATE_DEB = pd.Timestamp("2025-01-01")
DATE_FIN = pd.Timestamp("2026-02-05")

DEPT_CODE = "21"
NATINF_PVE: List[str] = ["27742", "25001"]
NATINF_PEJ: List[str] = ["27742", "25001"]

# Couches de référence SIG (communes, pochoir) : attendues dans ref/sig
COMMUNES_SHP = get_sig_dir() / "communes_21" / "communes.shp"
POCHOIR_GPKG = get_sig_dir() / "Pochoir_SD21.gpkg"


# ---------------------------------------------------------------------------
# Analyses (spécifiques au bilan agrainage)
# ---------------------------------------------------------------------------

def analyse_controles_agrainage(
    point: pd.DataFrame, tub: pd.DataFrame, pnf: pd.DataFrame, out_dir: Path
) -> pd.DataFrame:
    """
    Contrôles agrainage issus des points de contrôle OSCEAN (point déjà filtré par le loader).
    1 contrôle = 1 ligne (1 localisation) ; un même dc_id peut avoir plusieurs lignes.
    """
    tub_codes = set(tub["INSEE_COM"].unique())
    pnf_codes = set(pnf["CODE_INSEE"].unique())

    pt = point.copy()

    nom = pt["nom_dossie"].astype(str).str.lower()
    m_nom = nom.str.contains("agrain", na=False)
    m_snc = pt["type_actio"].astype(str).str.contains(
        "police sanitaire de la faune sauvage", case=False, na=False,
    )
    exclusions = r"tubercul|grippe|pi[eé]geage"
    m_excl = nom.str.contains(exclusions, na=False)
    pt["is_agrain"] = m_nom | (m_snc & ~m_excl)
    agrain = pt[pt["is_agrain"]].copy()
    agrain["insee_comm"] = agrain["insee_comm"].astype(str).str.zfill(5)

    # 1 contrôle = 1 ligne (1 localisation) : décompte et tableaux basés sur agrain
    nb_total = len(agrain)
    tab_resultats = (
        agrain["resultat"]
        .value_counts()
        .rename_axis("resultat")
        .to_frame("nb")
        .reset_index()
    )
    tab_resultats["taux"] = tab_resultats["nb"] / float(nb_total or 1)
    tab_resultats.to_csv(
        out_dir / "controles_agrainage_resultats.csv", sep=";", index=False
    )

    summary = _zone_summary(agrain, "insee_comm", tub_codes, pnf_codes)
    summary.to_csv(
        out_dir / "controles_agrainage_par_zone.csv", sep=";", index=False
    )

    cols_export = [
        c for c in [
            "fid", "dc_id", "nom_dossie", "date_ctrl", "num_depart",
            "nom_commun", "insee_comm", "dr", "entit_ctrl",
            "type_usage", "domaine", "theme", "type_actio",
            "fc_type", "resultat", "code_pej", "code_pa",
            "natinf_pej", "x", "y",
        ] if c in agrain.columns
    ]
    agrain[cols_export].to_csv(
        out_dir / "controles_agrainage_points.csv", sep=";", index=False
    )

    return agrain


def export_point_ctrl_sig(root: Path) -> None:
    """
    S'assure que le dossier sources/sig existe pour les couches de points de contrôle.

    Les couches de contrôles sont désormais les GPKG point_ctrl_YYYYMMDD_wgs84.gpkg
    déjà présents dans sources/sig (pas de génération depuis des CSV).
    """
    sig_dir = get_sources_sig_dir()
    sig_dir.mkdir(parents=True, exist_ok=True)


def analyse_pve_agrainage(
    root: Path, tub: pd.DataFrame, pnf: pd.DataFrame, out_dir: Path
) -> pd.DataFrame:
    """PVe avec NATINF 27742 (pve déjà filtré par le loader si appelé avec dept/période)."""
    tub_codes = set(tub["INSEE_COM"].unique())
    pnf_codes = set(pnf["CODE_INSEE"].unique())

    pve = load_pve(root, dept_code=DEPT_CODE, date_deb=DATE_DEB, date_fin=DATE_FIN)
    pve_p = pve.copy()
    pve_p["is_agrain"] = pve_p["INF-NATINF"].apply(
        lambda x: contient_natinf(x, NATINF_PVE)
    )
    pve_agrain = pve_p[pve_p["is_agrain"]].copy()

    # Jointure avec la table des centroïdes de communes pour compléter / construire
    # les coordonnées de chaque PVe à partir du code INSEE.
    communes_centroides = load_communes_centroides(root)

    if "INF-INSEE" in pve_agrain.columns:
        pve_agrain["INF-INSEE"] = (
            pve_agrain["INF-INSEE"].astype(str).str.zfill(5)
        )
        pve_agrain = pve_agrain.merge(
            communes_centroides,
            left_on="INF-INSEE",
            right_on="code_insee",
            how="left",
        )
    else:
        # Même sans INF-INSEE explicite, on conserve tout de même le DataFrame
        pve_agrain["code_insee"] = ""
        pve_agrain["lat"] = pd.NA
        pve_agrain["lon"] = pd.NA

    # Priorité aux coordonnées GPS natives si elles existent et sont valides,
    # sinon on utilise les centroïdes de communes.
    has_native_coords = (
        {"inf_gps_lat", "inf_gps_long"}.issubset(pve_agrain.columns)
    )
    if has_native_coords:
        pve_agrain["inf_gps_lat"] = pd.to_numeric(
            pve_agrain["inf_gps_lat"], errors="coerce"
        )
        pve_agrain["inf_gps_long"] = pd.to_numeric(
            pve_agrain["inf_gps_long"], errors="coerce"
        )

    def _coord_from_row(row):
        if has_native_coords and pd.notna(row.get("inf_gps_lat")) and pd.notna(
            row.get("inf_gps_long")
        ):
            return row["inf_gps_lat"], row["inf_gps_long"]
        return row.get("lat"), row.get("lon")

    coords = pve_agrain.apply(_coord_from_row, axis=1, result_type="expand")
    pve_agrain["lat_final"] = coords[0]
    pve_agrain["lon_final"] = coords[1]

    # PVe localisables : ceux pour lesquels on a réussi à obtenir des coordonnées
    mask_localisable = pve_agrain["lat_final"].notna() & pve_agrain["lon_final"].notna()
    pve_loc = pve_agrain[mask_localisable].copy()
    pve_non_loc = pve_agrain[~mask_localisable].copy()

    nb_dept = len(pve_agrain)

    # Comptages par zone basés sur les codes INSEE (comportement "fallback"
    # généralisé, sans dépendance au shapefile PVE).
    if "INF-INSEE" in pve_loc.columns:
        pve_loc["INF-INSEE"] = pve_loc["INF-INSEE"].astype(str).str.zfill(5)
        nb_tub = int(pve_loc["INF-INSEE"].isin(tub_codes).sum())
        nb_pnf = int(pve_loc["INF-INSEE"].isin(pnf_codes).sum())
    else:
        nb_tub = 0
        nb_pnf = 0

    zone_pve = pd.DataFrame([
        {"zone": "Département", "nb": nb_dept},
        {"zone": "Zone TUB", "nb": nb_tub},
        {"zone": "PNF", "nb": nb_pnf},
    ])
    zone_pve.to_csv(out_dir / "pve_agrainage_par_zone.csv", sep=";", index=False)

    # Export des PVe avec coordonnées (points) pour la cartographie.
    cols_pve = [
        c
        for c in [
            "INF-ID",
            "INF-DATE-INTG",
            "INF-NATINF",
            "INF-TYP-INF-STAT-LIB",
            "INF-DEPARTEMENT",
            "INF-INSEE",
            "INF-CP",
            "lat_final",
            "lon_final",
        ]
        if c in pve_loc.columns
    ]
    pve_loc[cols_pve].to_csv(
        out_dir / "pve_agrainage_points.csv", sep=";", index=False
    )

    # Export SIG : points PVe (WGS84) pour le projet QGIS / production cartographique.
    if not pve_loc.empty:
        try:
            gdf_pve = gpd.GeoDataFrame(
                pve_loc.copy(),
                geometry=gpd.points_from_xy(
                    pve_loc["lon_final"], pve_loc["lat_final"]
                ),
                crs="EPSG:4326",
            )
            sig_dir = get_sources_sig_dir()
            sig_dir.mkdir(parents=True, exist_ok=True)
            gpkg_path = sig_dir / "pve_agrainage_points_centroides.gpkg"
            gdf_pve.to_file(
                gpkg_path,
                layer="pve_agrainage_points_centroides",
                driver="GPKG",
            )
        except Exception as e:
            print(
                f"  [WARN] Impossible de générer le GPKG des points PVe agrainage : {e}"
            )

    # Export des PVe non localisables (contrôle qualité)
    if not pve_non_loc.empty:
        cols_nl = [
            c
            for c in [
                "INF-ID",
                "INF-DATE-INTG",
                "INF-NATINF",
                "INF-TYP-INF-STAT-LIB",
                "INF-DEPARTEMENT",
                "INF-INSEE",
                "INF-CP",
            ]
            if c in pve_non_loc.columns
        ]
        pve_non_loc[cols_nl].to_csv(
            out_dir / "pve_agrainage_sans_localisation.csv",
            sep=";",
            index=False,
        )

    pd.DataFrame([{"nb_pve_agrainage": len(pve_agrain)}]).to_csv(
        out_dir / "pve_agrainage_resume.csv", sep=";", index=False
    )

    return pve_agrain


def analyse_pa_agrainage(
    root: Path, point_ctrl_agrain: pd.DataFrame, out_dir: Path
) -> None:
    """
    PA agrainage : PA du département (DC_ID dans contrôles agrainage ou entité SD21),
    filtrées par thème/type d'action contenant « agrainage ».
    """
    pa = load_pa(root, date_deb=DATE_DEB, date_fin=DATE_FIN)
    dc_ids_ctrl = (
        set(point_ctrl_agrain["dc_id"].dropna().astype(str).unique())
        if not point_ctrl_agrain.empty and "dc_id" in point_ctrl_agrain.columns
        else set()
    )
    entity_sd = "SD" + DEPT_CODE
    dept_mask = pa["DC_ID"].astype(str).isin(dc_ids_ctrl)
    if "ENTITE_ORIGINE_PROCEDURE" in pa.columns:
        dept_mask = dept_mask | (pa["ENTITE_ORIGINE_PROCEDURE"].astype(str).str.strip() == entity_sd)
    pa_dept = pa[dept_mask].copy()

    theme = pa_dept.get("THEME", pd.Series(dtype=object)).fillna("").astype(str).str.lower()
    type_act = pa_dept.get("TYPE_ACTION", pd.Series(dtype=object)).fillna("").astype(str).str.lower()
    mask_agrain = (
        theme.str.contains("agrainage", regex=False)
        | type_act.str.contains("agrainage", regex=False)
        | theme.str.contains("agrain", regex=False)
        | type_act.str.contains("agrain", regex=False)
    )
    pa_agrain = pa_dept[mask_agrain].copy()

    id_col = "DC_ID" if "DC_ID" in pa_agrain.columns else None
    nb_pa_agrain = pa_agrain[id_col].nunique() if id_col else len(pa_agrain)

    pa_agrain.to_csv(out_dir / "pa_agrainage.csv", sep=";", index=False)

    if not pa_agrain.empty and "DOMAINE" in pa_agrain.columns and "THEME" in pa_agrain.columns:
        pa_par_theme = (
            pa_agrain.groupby(["DOMAINE", "THEME"])
            .size()
            .rename("nb_pa")
            .reset_index()
        )
        pa_par_theme.to_csv(out_dir / "pa_agrainage_par_theme.csv", sep=";", index=False)
    else:
        pa_par_theme = pd.DataFrame(columns=["DOMAINE", "THEME", "nb_pa"])
        pa_par_theme.to_csv(out_dir / "pa_agrainage_par_theme.csv", sep=";", index=False)

    pd.DataFrame([{"nb_pa_agrainage": nb_pa_agrain}]).to_csv(
        out_dir / "pa_agrainage_resume.csv", sep=";", index=False
    )


def analyse_pej_agrainage(
    root: Path, point_ctrl_agrain: pd.DataFrame,
    tub: pd.DataFrame, pnf: pd.DataFrame, out_dir: Path
) -> pd.DataFrame:
    """
    PEJ avec NATINF 27742 ou 25001, département Côte-d'Or.
    Périmètre unifié : PEJ SD21 + NATINF agrainage (avec ou sans géométrie),
    dédupliqués par DC_ID, pour que le total ne dépende pas du GPKG points PJ.
    """
    tub_codes = set(tub["INSEE_COM"].unique())
    pnf_codes = set(pnf["CODE_INSEE"].unique())

    pej = load_pej(root, date_deb=DATE_DEB, date_fin=DATE_FIN)
    entity_sd = "SD" + DEPT_CODE
    if "ENTITE_ORIGINE_PROCEDURE" in pej.columns:
        pej_dept = pej[pej["ENTITE_ORIGINE_PROCEDURE"] == entity_sd].copy()
    else:
        pej_dept = pej.copy()
    # Test NATINF agrainage (vectorisé pour éviter .apply sur de nombreuses lignes)
    _natinf_pattern = "|".join(rf"(^|_){re.escape(c)}(_|$)" for c in NATINF_PEJ)
    pej_dept["is_agrain"] = (
        pej_dept["NATINF_PEJ"].fillna("").astype(str).str.contains(_natinf_pattern, regex=True)
    )
    pej_agrain = (
        pej_dept[pej_dept["is_agrain"]]
        .copy()
        .sort_values("DATE_REF", ascending=False)
        .drop_duplicates(subset="DC_ID", keep="first")
    )
    pej_agrain["insee_comm"] = ""

    # Enrichissement insee_comm via géométrie (GPKG points PJ) si disponible
    try:
        pej_gdf = load_pj_with_geometry(
            root, NATINF_PEJ, DEPT_CODE, date_deb=DATE_DEB, date_fin=DATE_FIN, pej_df=pej
        )
        if not pej_gdf.empty:
            _natinf_pattern = "|".join(rf"(^|_){re.escape(c)}(_|$)" for c in NATINF_PEJ)
            pej_gdf["is_agrain"] = (
                pej_gdf["NATINF_PEJ"].fillna("").astype(str).str.contains(_natinf_pattern, regex=True)
            )
            pej_with_geom = (
                pej_gdf[pej_gdf["is_agrain"]]
                .sort_values("DATE_REF", ascending=False)
                .drop_duplicates(subset="DC_ID", keep="first")
            )
            pej_with_geom = pej_with_geom[pej_with_geom["geometry"].notna()].copy()
            if not pej_with_geom.empty and COMMUNES_SHP.exists():
                communes = gpd.read_file(COMMUNES_SHP)
                insee_col = _detect_insee_column(communes)
                communes[insee_col] = communes[insee_col].astype(str).str.zfill(5)
                crs_points = pej_gdf.crs if pej_gdf.crs else "EPSG:4326"
                pej_join = gpd.GeoDataFrame(
                    pej_with_geom, geometry="geometry", crs=crs_points
                )
                if pej_join.crs != communes.crs:
                    pej_join = pej_join.to_crs(communes.crs)
                joined = gpd.sjoin(
                    pej_join, communes[[insee_col, "geometry"]], how="left", predicate="within"
                )
                insee_map = joined.set_index("DC_ID")[insee_col].to_dict()
                pej_agrain["insee_comm"] = pej_agrain["DC_ID"].map(insee_map).fillna("")
    except Exception as e:
        print(f"  [WARN] Enrichissement géométrique PEJ non disponible : {e}")

    pej_agrain["insee_comm"] = pej_agrain["insee_comm"].fillna("").astype(str).str.zfill(5)

    valid_insee = pej_agrain["insee_comm"].str.match(r"^[0-9]{5}$") & (pej_agrain["insee_comm"] != "00000")
    pej_with_commune = pej_agrain[valid_insee]
    if not pej_with_commune.empty:
        zone_pej = _zone_count(pej_with_commune, "insee_comm", tub_codes, pnf_codes)
    else:
        zone_pej = pd.DataFrame([
            {"zone": "Département", "nb": len(pej_agrain)},
            {"zone": "Zone TUB", "nb": 0},
            {"zone": "PNF", "nb": 0},
        ])
    zone_pej.loc[zone_pej["zone"] == "Département", "nb"] = len(pej_agrain)
    zone_pej.to_csv(out_dir / "pej_agrainage_par_zone.csv", sep=";", index=False)

    cols_pej = [
        c for c in [
            "DC_ID", "DATE_REF", "NATINF_PEJ", "DOMAINE", "THEME",
            "TYPE_ACTION", "DUREE_PEJ", "CLOTUR_PEJ", "SUITE",
            "ENTITE_ORIGINE_PROCEDURE", "insee_comm",
        ] if c in pej_agrain.columns
    ]
    pej_agrain[cols_pej].to_csv(
        out_dir / "pej_agrainage.csv", sep=";", index=False
    )

    duree_moy = (
        pd.to_numeric(pej_agrain["DUREE_PEJ"], errors="coerce").mean()
        if "DUREE_PEJ" in pej_agrain.columns
        else None
    )
    pd.DataFrame([{
        "nb_pej_agrainage": len(pej_agrain),
        "duree_moy_pej": duree_moy,
    }]).to_csv(out_dir / "pej_agrainage_resume.csv", sep=";", index=False)

    return pej_agrain


def generer_synthese(out_dir: Path) -> pd.DataFrame:
    """Tableau de synthèse croisant les 3 sources par zone."""
    ctrl = pd.read_csv(out_dir / "controles_agrainage_par_zone.csv", sep=";")
    pve = pd.read_csv(out_dir / "pve_agrainage_par_zone.csv", sep=";")
    pej = pd.read_csv(out_dir / "pej_agrainage_par_zone.csv", sep=";")

    synth = ctrl[["zone", "nb_total", "nb_infraction"]].rename(
        columns={"nb_total": "ctrl_total", "nb_infraction": "ctrl_infraction"}
    )
    synth = synth.merge(pve[["zone", "nb"]].rename(columns={"nb": "pve_nb"}), on="zone", how="outer")
    synth = synth.merge(pej[["zone", "nb"]].rename(columns={"nb": "pej_nb"}), on="zone", how="outer")
    synth = synth.fillna(0)
    for c in ["ctrl_total", "ctrl_infraction", "pve_nb", "pej_nb"]:
        synth[c] = synth[c].astype(int)
    synth.to_csv(out_dir / "synthese_agrainage.csv", sep=";", index=False)
    return synth


def generate_pdf_report(root: Path, out_dir: Path) -> None:
    styles = _get_styles()
    tmp_dir = Path(tempfile.mkdtemp(prefix="ofb_agrainage_"))
    try:
        _generate_pdf_content(root, out_dir, tmp_dir, styles)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _generate_pdf_content(root: Path, out_dir: Path, tmp_dir: Path, styles) -> None:
    """Corps de la génération PDF, séparé pour garantir le nettoyage de tmp_dir."""
    avail_w = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT

    # Libellés NATINF (agrainage) pour les textes descriptifs.
    natinf_ref = load_natinf_ref(root)
    natinf_ref = (
        natinf_ref.set_index("numero_natinf")["libelle_natinf"]
        if natinf_ref is not None and not natinf_ref.empty
        else None
    )

    def _format_natinf_list(codes: List[str]) -> str:
        if not codes:
            return ""
        if natinf_ref is None:
            return ", ".join(codes)
        parts: List[str] = []
        for c in codes:
            lib = natinf_ref.get(str(c))
            if isinstance(lib, str) and lib.strip():
                parts.append(f"{c} – {lib.strip()}")
            else:
                parts.append(str(c))
        return "; ".join(parts)

    natinf_pve_txt = _format_natinf_list(NATINF_PVE)
    natinf_pej_txt = _format_natinf_list(NATINF_PEJ)
    tab_resultats = pd.read_csv(out_dir / "controles_agrainage_resultats.csv", sep=";")
    ctrl_zones = pd.read_csv(out_dir / "controles_agrainage_par_zone.csv", sep=";")
    pve_zones = pd.read_csv(out_dir / "pve_agrainage_par_zone.csv", sep=";")
    pej_zones = pd.read_csv(out_dir / "pej_agrainage_par_zone.csv", sep=";")
    synthese = pd.read_csv(out_dir / "synthese_agrainage.csv", sep=";")

    pve_resume = pd.read_csv(out_dir / "pve_agrainage_resume.csv", sep=";")
    pej_resume = pd.read_csv(out_dir / "pej_agrainage_resume.csv", sep=";")
    pa_resume = pd.read_csv(out_dir / "pa_agrainage_resume.csv", sep=";") if (out_dir / "pa_agrainage_resume.csv").exists() else pd.DataFrame()
    pa_par_theme = pd.read_csv(out_dir / "pa_agrainage_par_theme.csv", sep=";") if (out_dir / "pa_agrainage_par_theme.csv").exists() else pd.DataFrame()

    nb_ctrl_dept = int(ctrl_zones.loc[ctrl_zones["zone"] == "Département", "nb_total"].iloc[0]) if not ctrl_zones.empty else 0
    nb_ctrl_inf = int(ctrl_zones.loc[ctrl_zones["zone"] == "Département", "nb_infraction"].iloc[0]) if not ctrl_zones.empty else 0
    nb_pve = int(pve_resume["nb_pve_agrainage"].iloc[0]) if not pve_resume.empty else 0
    nb_pej = int(float(pej_resume["nb_pej_agrainage"].iloc[0])) if not pej_resume.empty else 0
    nb_pa_agrain = int(pa_resume["nb_pa_agrainage"].iloc[0]) if not pa_resume.empty else 0
    duree_moy = float(pej_resume["duree_moy_pej"].iloc[0]) if not pej_resume.empty and pd.notna(pej_resume["duree_moy_pej"].iloc[0]) else 0

    pdf_path = out_dir / "bilan_agrainage_Cote_dOr.pdf"

    sections = [
        ("sec1", "I. Localisations de contrôle agrainage"),
        ("sec2", "II. Infractions PVe"),
        ("sec_pa", "III. Procédures administratives (PA)"),
        ("sec3", "IV. Procédures judiciaires (PEJ)"),
        ("sec4", "V. Synthèse par zone"),
        ("sec5", "VI. Cartographie"),
        ("sec6", "Annexes"),
    ]

    content_frame = Frame(
        MARGIN_LEFT, MARGIN_BOTTOM,
        PAGE_W - MARGIN_LEFT - MARGIN_RIGHT,
        PAGE_H - MARGIN_TOP - MARGIN_BOTTOM,
        id="content",
    )

    def _header_footer(canvas, doc):
        canvas.saveState()
        if IMG_FOOTER_DECO.exists():
            canvas.drawImage(
                str(IMG_FOOTER_DECO), PAGE_W - 60 * mm, 0,
                width=60 * mm, height=7 * mm,
                preserveAspectRatio=True, mask="auto",
            )
        canvas.setStrokeColor(rl_colors.HexColor(COLOR_PRIMARY))
        canvas.setLineWidth(2)
        y_header = PAGE_H - 16 * mm
        canvas.line(MARGIN_LEFT, y_header, PAGE_W - MARGIN_RIGHT, y_header)
        canvas.setFont(f"{FONT_FAMILY}-Bold", 8)
        canvas.setFillColor(rl_colors.HexColor(COLOR_PRIMARY))
        canvas.drawString(MARGIN_LEFT, y_header + 3, "Bilan agrainage – Côte-d'Or")
        y_foot = 8 * mm
        canvas.setFont(f"{FONT_FAMILY}-Bold", 7)
        canvas.setFillColor(rl_colors.HexColor(COLOR_SECONDARY))
        canvas.drawString(MARGIN_LEFT, y_foot + 12, "Office français de la biodiversité")
        canvas.setFont(FONT_FAMILY, 7)
        canvas.drawString(MARGIN_LEFT, y_foot + 3,
                         "SD de la Côte-d'Or – 57, rue de Mulhouse – 21000 Dijon – www.ofb.gouv.fr")
        canvas.drawRightString(PAGE_W - MARGIN_RIGHT, y_foot + 3, f"{doc.page}")
        canvas.restoreState()

    def _title_page_template(canvas, doc):
        canvas.saveState()
        if IMG_BACKGROUND.exists():
            canvas.drawImage(str(IMG_BACKGROUND), 0, 0, width=PAGE_W, height=PAGE_H * 0.86,
                            preserveAspectRatio=False, mask="auto")
        if IMG_LOGO_BANNER.exists():
            canvas.drawImage(str(IMG_LOGO_BANNER), 0, PAGE_H * 0.86, width=PAGE_W, height=PAGE_H * 0.14,
                            preserveAspectRatio=False, mask="auto")
        cx = PAGE_W / 2
        canvas.setFont(f"{FONT_FAMILY}-Bold", 26)
        canvas.setFillColor(rl_colors.HexColor(COLOR_PRIMARY))
        canvas.drawCentredString(cx, PAGE_H * 0.62 + 14, "Bilan de l'activité de contrôle")
        canvas.drawCentredString(cx, PAGE_H * 0.62 - 20, "Agrainage")
        canvas.setFont(f"{FONT_FAMILY}-Bold", 22)
        canvas.drawCentredString(cx, PAGE_H * 0.50, "Côte-d'Or")
        canvas.setFont(FONT_FAMILY, 14)
        canvas.setFillColor(rl_colors.HexColor(COLOR_GREY))
        canvas.drawCentredString(cx, PAGE_H * 0.42, f"Période : {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y}")
        canvas.setFont(FONT_FAMILY, 11)
        canvas.drawCentredString(cx, PAGE_H * 0.36, "Données OSCEAN & PVe OFB")
        canvas.setFont(f"{FONT_FAMILY}-Bold", 7)
        canvas.setFillColor(rl_colors.HexColor(COLOR_SECONDARY))
        canvas.drawString(MARGIN_LEFT, 30, "Office français de la biodiversité")
        canvas.setFont(FONT_FAMILY, 7)
        canvas.drawString(MARGIN_LEFT, 20, "SD de la Côte-d'Or – www.ofb.gouv.fr")
        canvas.restoreState()

    title_template = PageTemplate(id="TitlePage", frames=[Frame(0, 0, PAGE_W, PAGE_H, id="full")], onPage=_title_page_template)
    normal_template = PageTemplate(id="Normal", frames=[content_frame], onPage=_header_footer)

    doc = BaseDocTemplate(
        str(pdf_path), pagesize=A4,
        title="Bilan agrainage – Côte-d'Or",
        author="Office français de la biodiversité",
    )
    doc.addPageTemplates([title_template, normal_template])

    story = []
    story.append(NextPageTemplate("Normal"))
    story.append(PageBreak())

    story.append(Paragraph("Sommaire", styles["Title"]))
    story.append(Spacer(1, 6 * mm))
    for anchor, sec_title in sections:
        story.append(Paragraph(f'<a href="#{anchor}" color="{COLOR_PRIMARY}">{sec_title}</a>', styles["TOCEntry"]))
    story.append(PageBreak())

    story.append(Paragraph('<a name="sec1"/>I. Localisations de contrôle agrainage', styles["Heading1"]))
    story.append(Paragraph(
        "Localisations de contrôle liées à l'agrainage identifiées dans les points de contrôle OSCEAN (champ « nom_dossie » contenant « agrain »).",
        styles["BodyText"],
    ))
    story.append(Spacer(1, 3 * mm))

    taux_inf = nb_ctrl_inf / nb_ctrl_dept if nb_ctrl_dept else 0
    story.append(key_figures_table([
        (str(nb_ctrl_dept), "Localisations de contrôle agrainage"),
        (str(nb_ctrl_inf), "Infractions"),
        (f"{taux_inf:.1%}", "Taux d'infraction"),
    ], styles))
    story.append(Spacer(1, 5 * mm))

    story.append(Paragraph("Tableau 1 : Résultats des contrôles agrainage (département)", styles["TableCaption"]))
    tbl_data = [["Résultat", "Nombre", "Taux"]]
    for _, row in tab_resultats.iterrows():
        taux_str = f"{row['taux']:.1%}" if pd.notna(row.get("taux")) else "n.d."
        tbl_data.append([str(row["resultat"]), str(int(row["nb"])), taux_str])
    story.append(ofb_table(tbl_data, col_widths=[avail_w * 0.50, avail_w * 0.25, avail_w * 0.25], col_aligns=["LEFT", "RIGHT", "RIGHT"]))
    story.append(Spacer(1, 3 * mm))

    pie_data = {str(row["resultat"]): int(row["nb"]) for _, row in tab_resultats.iterrows()}
    if pie_data:
        pie_path = chart_pie(pie_data, "Répartition des résultats", tmp_dir, "pie_resultats.png")
        _pimg = PILImage.open(pie_path)
        _target_w = avail_w * 0.75
        _target_h = _target_w * (_pimg.height / _pimg.width)
        _pimg.close()
        img_pie = RLImage(pie_path, width=_target_w, height=_target_h)
        img_pie.hAlign = "CENTER"
        story.append(img_pie)
    story.append(Spacer(1, 5 * mm))

    tbl_zone = [["Zone", "Nb contrôles", "Nb conformes", "Nb infractions", "Taux infraction"]]
    for _, row in ctrl_zones.iterrows():
        taux_str = f"{row['taux_infraction']:.1%}" if pd.notna(row.get("taux_infraction")) else "n.d."
        tbl_zone.append([str(row["zone"]), str(int(row["nb_total"])), str(int(row["nb_conforme"])), str(int(row["nb_infraction"])), taux_str])
    _desc_parts = []
    for _, _r in ctrl_zones.iterrows():
        _t = f"{_r['taux_infraction']:.1%}" if pd.notna(_r.get("taux_infraction")) else "n.d."
        _desc_parts.append(f"{_r['zone']} : {int(_r['nb_total'])} contrôle(s), {int(_r['nb_infraction'])} infraction(s) ({_t})")
    zone_block = [
        Paragraph("Tableau 2 : Localisations de contrôle agrainage par zone", styles["TableCaption"]),
        ofb_table(tbl_zone, col_widths=[avail_w * 0.22, avail_w * 0.20, avail_w * 0.20, avail_w * 0.20, avail_w * 0.18], col_aligns=["LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"]),
        Paragraph("Répartition par zone : " + " ; ".join(_desc_parts) + ".", styles["BodyText"]),
        Paragraph(f"Sources : OFB/OSCEAN \u2013 IGN/BD TOPO \u2013 ESRI World Topographic \u2013 MNHN/Espaces prot\u00e9g\u00e9s. Contr\u00f4les agrainage du {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y}, d\u00e9partement 21.", styles["BodySmall"]),
    ]
    story.append(KeepTogether(zone_block))
    story.append(Spacer(1, 8 * mm))

    story.append(Paragraph('<a name="sec2"/>II. Infractions PVe', styles["Heading1"]))
    story.append(Paragraph(
        "Infractions relevées par Procès-Verbal électronique (PVe) pour les NATINF d'agrainage : "
        f"{natinf_pve_txt} (Côte-d'Or).",
        styles["BodyText"],
    ))
    story.append(Spacer(1, 3 * mm))
    story.append(key_figures_table([(str(nb_pve), "PVe agrainage")], styles))
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph("Tableau 3 : PVe agrainage par zone", styles["TableCaption"]))
    tbl_pve = [["Zone", "Nombre de PVe"]]
    for _, row in pve_zones.iterrows():
        tbl_pve.append([str(row["zone"]), str(int(row["nb"]))])
    story.append(ofb_table(tbl_pve, col_widths=[avail_w * 0.50, avail_w * 0.50], col_aligns=["LEFT", "RIGHT"]))
    story.append(Spacer(1, 8 * mm))

    story.append(Paragraph('<a name="sec_pa"/>III. Procédures administratives (PA)', styles["Heading1"]))
    story.append(Paragraph("Procédures administratives (PA) liées à l'agrainage : thème ou type d'action contenant « agrainage ».", styles["BodyText"]))
    story.append(Spacer(1, 3 * mm))
    story.append(key_figures_table([(str(nb_pa_agrain), "PA agrainage")], styles))
    story.append(Spacer(1, 5 * mm))
    if not pa_par_theme.empty:
        tbl_pa = [["Domaine", "Thème", "Nombre"]]
        for _, row in pa_par_theme.iterrows():
            tbl_pa.append([str(row.get("DOMAINE", "")), str(row.get("THEME", "")), str(int(row["nb_pa"]))])
        tbl_pa.append(["", "Total", str(nb_pa_agrain)])
        story.append(Paragraph("Tableau 4 : PA agrainage par thème", styles["TableCaption"]))
        story.append(ofb_table(tbl_pa, col_widths=[avail_w * 0.35, avail_w * 0.45, avail_w * 0.20], col_aligns=["LEFT", "LEFT", "RIGHT"]))
    else:
        story.append(Paragraph("Aucune procédure administrative agrainage enregistrée sur la période.", styles["BodyText"]))
    story.append(Spacer(1, 8 * mm))

    duree_txt = f"{duree_moy:.0f} j" if duree_moy else "n.d."
    story.append(Paragraph('<a name="sec3"/>IV. Procédures judiciaires (PEJ)', styles["Heading1"]))
    story.append(Paragraph(
        "Infractions ayant fait l'objet d'une procédure d'enquête judiciaire (PEJ) pour les NATINF d'agrainage : "
        f"{natinf_pej_txt} (Côte-d'Or).",
        styles["BodyText"],
    ))
    story.append(Spacer(1, 3 * mm))
    story.append(key_figures_table([(str(nb_pej), "PEJ agrainage"), (duree_txt, "Durée moy. PEJ")], styles))
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph("Tableau 5 : PEJ agrainage par zone", styles["TableCaption"]))
    tbl_pej = [["Zone", "Nombre de PEJ"]]
    for _, row in pej_zones.iterrows():
        tbl_pej.append([str(row["zone"]), str(int(row["nb"]))])
    story.append(ofb_table(tbl_pej, col_widths=[avail_w * 0.50, avail_w * 0.50], col_aligns=["LEFT", "RIGHT"]))
    story.append(Spacer(1, 8 * mm))

    story.append(Paragraph('<a name="sec4"/>V. Synthèse par zone', styles["Heading1"]))
    story.append(Paragraph("Tableau croisé des différentes sources d'infractions d'agrainage par périmètre géographique (département, zone TUB, PNF).", styles["BodyText"]))
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph("Tableau 6 : Synthèse agrainage par zone et source", styles["TableCaption"]))
    tbl_synth = [["Zone", "Ctrl total", "Ctrl infractions", "PVe", "PEJ"]]
    for _, row in synthese.iterrows():
        tbl_synth.append([str(row["zone"]), str(int(row["ctrl_total"])), str(int(row["ctrl_infraction"])), str(int(row["pve_nb"])), str(int(row["pej_nb"]))])
    story.append(ofb_table(tbl_synth, col_widths=[avail_w * 0.22, avail_w * 0.20, avail_w * 0.20, avail_w * 0.19, avail_w * 0.19], col_aligns=["LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"]))
    story.append(Spacer(1, 5 * mm))

    zone_labels = synthese["zone"].tolist()
    series_synth = {
        "Ctrl infractions": synthese["ctrl_infraction"].astype(int).tolist(),
        "PVe": synthese["pve_nb"].astype(int).tolist(),
        "PEJ": synthese["pej_nb"].astype(int).tolist(),
    }
    if any(sum(v) > 0 for v in series_synth.values()):
        bar_synth_path = chart_bar_grouped(zone_labels, series_synth, "Infractions agrainage par zone et source", "Nombre", tmp_dir, "bar_synthese_zones.png")
        img_synth = RLImage(bar_synth_path, width=avail_w * 0.85, height=avail_w * 0.50)
        img_synth.hAlign = "CENTER"
        story.append(img_synth)
    story.append(Spacer(1, 3 * mm))

    total_sources = [nb_ctrl_inf, nb_pve, nb_pej]
    if any(v > 0 for v in total_sources):
        bar_src_path = chart_bar(["Ctrl OSCEAN", "PVe", "PEJ"], total_sources, "Infractions agrainage par source (département)", "Nombre", tmp_dir, "bar_sources.png", color="#53AB60")
        img_src = RLImage(bar_src_path, width=avail_w * 0.6, height=avail_w * 0.36)
        img_src.hAlign = "CENTER"
        story.append(img_src)
    story.append(PageBreak())

    story.append(Paragraph('<a name="sec5"/>VI. Cartographie', styles["Heading1"]))
    carte_path = get_cartes_dir() / "carte_agrainage.png"
    if carte_path.exists():
        story.append(Paragraph("Agrainage illicite (NATINF 27742 et 25001) — Côte-d'Or", styles["Heading2"]))
        story.append(Spacer(1, 3 * mm))
        img_carte = RLImage(str(carte_path), width=avail_w, height=avail_w * 0.75)
        img_carte.hAlign = "CENTER"
        story.append(img_carte)
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph("Sources : OFB/OSCEAN – IGN/BD TOPO – ESRI World Topographic – MNHN/Espaces protégés. Réalisation : OFB SD21 – Mars 2026.", styles["BodySmall"]))
    else:
        story.append(Paragraph("La carte n'a pas été trouvée dans out/generateur_de_cartes/carte_agrainage.png.", styles["BodyText"]))
    story.append(PageBreak())

    story.append(Paragraph('<a name="sec6"/>Annexes', styles["Heading1"]))
    story.append(Paragraph("Méthodologie", styles["Heading2"]))
    methodo_text = (
        f"<b>Période d'analyse :</b> du {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y}.<br/>"
        f"<b>Périmètre :</b> département de la Côte-d'Or (21).<br/>"
        "<b>Sources :</b> OSCEAN (points de contrôle, PEJ, PA) et PVe OFB.<br/>"
        "<b>Localisations de contrôle agrainage :</b> points de contrôle dont le champ « nom_dossie » contient « agrain ».<br/>"
        "<b>PVe agrainage :</b> NATINF 27742 et 25001.<br/>"
        "<b>PA agrainage :</b> procédures administratives (thème/type d'action contenant « agrainage »).<br/>"
        "<b>PEJ agrainage :</b> NATINF 27742 ou 25001.<br/>"
        "<b>Zone TUB :</b> communes de la zone TUB (source : référentiel TUB, ref/).<br/>"
        "<b>PNF :</b> communes situées dans le périmètre du Parc national de forêts."
    )
    story.append(Paragraph(methodo_text, styles["BodyText"]))
    story.append(Spacer(1, 6 * mm))

    story.append(Paragraph("Glossaire", styles["Heading2"]))
    glossaire = [
        ["Abréviation", "Signification"],
        ["DC", "Dossier de contrôle"],
        ["NATINF", "Nature d'infraction (nomenclature nationale)"],
        ["OSCEAN", "Outil pour la surveillance et le contrôle eau et nature (application nationale)"],
        ["PEJ", "Procédure d'enquête judiciaire"],
        ["PNF", "Parc national de forêts"],
        ["PVe", "Procès-verbal électronique"],
        ["TUB", "Zone délimitée par l'arrêté préfectoral « Tuberculose bovine »"],
    ]
    story.append(ofb_table(glossaire, col_widths=[avail_w * 0.25, avail_w * 0.75], col_aligns=["LEFT", "LEFT"]))

    doc.build(story)


def _parse_args() -> argparse.Namespace:
    """Paramètres de ligne de commande pour le bilan agrainage."""
    parser = argparse.ArgumentParser(
        description=(
            "Génère le bilan agrainage (OSCEAN, PVe, PEJ) pour un département "
            "et une période donnés à partir des sources OSCEAN/PVe."
        )
    )
    parser.add_argument(
        "--date-deb",
        type=str,
        default=None,
        help="Date de début de la période (YYYY-MM-DD). Si absent, saisie interactive ou erreur en batch.",
    )
    parser.add_argument(
        "--date-fin",
        type=str,
        default=None,
        help="Date de fin de la période (YYYY-MM-DD). Si absent, saisie interactive ou erreur en batch.",
    )
    parser.add_argument(
        "--dept-code",
        type=str,
        default=None,
        help="Code du département (ex: 21). Si absent, saisie interactive ou défaut 21.",
    )
    return parser.parse_args()


def run_bilan(date_deb: str, date_fin: str, dept_code: str) -> int:
    """Entry point callable by the orchestrator (no argparse, no subprocess)."""
    global DATE_DEB, DATE_FIN, DEPT_CODE
    try:
        DATE_DEB = pd.to_datetime(date_deb)
        DATE_FIN = pd.to_datetime(date_fin)
    except Exception:
        print("Paramètres date invalides.", file=sys.stderr)
        return 1
    DEPT_CODE = str(dept_code)

    root = _ROOT
    out_dir = get_out_dir("bilan_agrainage")

    print(
        f"Période analysée : du {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y} "
        f"(département {DEPT_CODE})."
    )

    print("Étape 1/7 : chargement des données...")
    with Spinner():
        point = load_point_ctrl(root, dept_code=DEPT_CODE, date_deb=DATE_DEB, date_fin=DATE_FIN)
        tub = load_tub(root)
        pnf = load_pnf(root)

    print("Étape 1bis : génération des couches SIG de points de contrôle (sources/sig)...")
    with Spinner():
        export_point_ctrl_sig(root)

    print("Étape 2/7 : analyse des contrôles agrainage (OSCEAN)...")
    with Spinner():
        dc_info = analyse_controles_agrainage(point, tub, pnf, out_dir)

    print("Étape 3/7 : analyse PVe agrainage...")
    with Spinner():
        analyse_pve_agrainage(root, tub, pnf, out_dir)

    print("Étape 4/7 : analyse PA agrainage...")
    with Spinner():
        analyse_pa_agrainage(root, dc_info, out_dir)

    print("Étape 5/7 : analyse PEJ agrainage...")
    with Spinner():
        analyse_pej_agrainage(root, dc_info, tub, pnf, out_dir)

    print("Étape 6/7 : génération de la synthèse...")
    with Spinner():
        generer_synthese(out_dir)

    print("Étape 7/7 : génération du bilan PDF...")
    with Spinner():
        generate_pdf_report(root, out_dir)

    print(f"Analyse terminée. Bilan généré dans '{out_dir}'.")
    return 0


def main() -> None:
    global DATE_DEB, DATE_FIN, DEPT_CODE

    args = _parse_args()
    if args.date_deb is None or args.date_fin is None or args.dept_code is None:
        date_deb_str, date_fin_str, dept_str = ask_periode_dept(
            date_deb_default=args.date_deb or str(DATE_DEB.date()),
            date_fin_default=args.date_fin or str(DATE_FIN.date()),
            dept_default=args.dept_code or DEPT_CODE,
        )
        args.date_deb = date_deb_str
        args.date_fin = date_fin_str
        args.dept_code = dept_str
    try:
        DATE_DEB = pd.to_datetime(args.date_deb)
        DATE_FIN = pd.to_datetime(args.date_fin)
    except Exception:
        raise SystemExit("Paramètres date invalides : utiliser le format YYYY-MM-DD pour --date-deb et --date-fin.")

    DEPT_CODE = str(args.dept_code)

    root = _ROOT
    out_dir = get_out_dir("bilan_agrainage")

    print(
        f"Période analysée : du {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y} "
        f"(département {DEPT_CODE})."
    )

    print("Étape 1/7 : chargement des données...")
    with Spinner():
        point = load_point_ctrl(root, dept_code=DEPT_CODE, date_deb=DATE_DEB, date_fin=DATE_FIN)
        tub = load_tub(root)
        pnf = load_pnf(root)

    print("Étape 1bis : génération des couches SIG de points de contrôle (sources/sig)...")
    with Spinner():
        export_point_ctrl_sig(root)

    print("Étape 2/7 : analyse des contrôles agrainage (OSCEAN)...")
    with Spinner():
        dc_info = analyse_controles_agrainage(point, tub, pnf, out_dir)

    print("Étape 3/7 : analyse PVe agrainage...")
    with Spinner():
        analyse_pve_agrainage(root, tub, pnf, out_dir)

    print("Étape 4/7 : analyse PA agrainage...")
    with Spinner():
        analyse_pa_agrainage(root, dc_info, out_dir)

    print("Étape 5/7 : analyse PEJ agrainage...")
    with Spinner():
        analyse_pej_agrainage(root, dc_info, tub, pnf, out_dir)

    print("Étape 6/7 : génération de la synthèse...")
    with Spinner():
        generer_synthese(out_dir)

    print("Étape 7/7 : génération du bilan PDF...")
    with Spinner():
        generate_pdf_report(root, out_dir)

    print(f"Analyse terminée. Bilan généré dans '{out_dir}'.")


if __name__ == "__main__":
    main()
