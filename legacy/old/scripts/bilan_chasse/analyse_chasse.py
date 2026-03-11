"""Bilan chasse — contrôles OSCEAN, PEJ, PA, PVe."""
import argparse
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import List

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
    Table,
    TableStyle,
)

# Bootstrap : exécution indépendante
_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from scripts.paths import get_cartes_dir, get_out_dir, get_sig_dir, PROJECT_ROOT
from scripts.common.loaders import (
    load_pa,
    load_pej,
    load_point_ctrl,
    load_pnf,
    load_pve,
    load_tub,
    load_natinf_ref,
    load_communes_noms,
)
from scripts.common.utils import (
    est_chasse_point,
    contient_natinf,
    get_dept_name,
    _load_csv_opt,
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
# Période d'analyse
# ---------------------------------------------------------------------------
DATE_DEB = pd.Timestamp("2025-09-01")
DATE_FIN = pd.Timestamp("2026-03-01")

DEPT_CODE = "21"

# NATINF pour le bilan chasse (PEJ et PVe) — liste fournie par le plan
NATINF_PVE_CHASSE: List[str] = [
    "26274", "26301", "27742", "27745",
    "20148", "20155", "20165", "20166",
    "2002",
]
NATINF_PEJ_CHASSE: List[str] = [
    "29821", "29844", "29845",  # Délit
    "26307", "323", "11975", "26305", "2172", "5984", "11889", "25001", "28969",  # C5
    "27742", "26303",  # C4
    "27741",  # C3
]

# Couches de référence SIG
COMMUNES_SHP = get_sig_dir() / "communes_21" / "communes.shp"
POCHOIR_GPKG = get_sig_dir() / "Pochoir_SD21.gpkg"


def analyse_controles_chasse(
    point: pd.DataFrame, pnf: pd.DataFrame, out_dir: Path
):
    """
    Contrôles chasse (point déjà filtré par le loader sur département et période).
    1 contrôle = 1 ligne (1 localisation) ; un même dc_id peut avoir plusieurs lignes.
    """
    point_periode = point.copy()

    # Filtre chasse
    point_periode["is_chasse"] = point_periode.apply(est_chasse_point, axis=1)
    point_chasse = point_periode[point_periode["is_chasse"]].copy()

    # Marquage PNF / hors PNF
    point_chasse["insee_comm"] = point_chasse["insee_comm"].astype(str).str.zfill(5)
    point_chasse = point_chasse.merge(
        pnf[["CODE_INSEE"]],
        left_on="insee_comm",
        right_on="CODE_INSEE",
        how="left",
    )
    point_chasse["PNF"] = point_chasse["CODE_INSEE"].notna().map(
        {True: "PNF", False: "Hors_PNF"}
    )
    point_chasse.drop(columns=["CODE_INSEE"], inplace=True)

    # 2.1 Nombre total de contrôles chasse (1 contrôle = 1 ligne / localisation)
    nb_controles_chasse = len(point_chasse)

    # 2.2 Résultats des contrôles (par ligne)
    tab_resultats = (
        point_chasse["resultat"]
        .value_counts()
        .rename_axis("resultat")
        .to_frame("nb")
        .reset_index()
    )
    tab_resultats["taux"] = tab_resultats["nb"] / float(nb_controles_chasse or 1)
    tab_resultats.to_csv(
        out_dir / "controles_chasse_resultats_dc.csv", sep=";", index=False
    )

    # 2.4 Indicateurs par commune (1 ligne = 1 contrôle)
    agg_commune = (
        point_chasse.groupby("insee_comm")
        .agg(
            nb_controles=("dc_id", "count"),
            nb_infractions=("resultat", lambda s: (s == "Infraction").sum()),
        )
        .reset_index()
    )
    agg_commune["taux_infraction"] = agg_commune["nb_infractions"] / agg_commune[
        "nb_controles"
    ].replace(0, pd.NA)
    agg_commune.to_csv(
        out_dir / "indicateurs_chasse_par_commune.csv", sep=";", index=False
    )

    # 2.6 Comparaison PNF / hors PNF (par ligne = 1 contrôle)
    agg_pnf = (
        point_chasse.groupby("PNF")
        .agg(
            nb_controles=("dc_id", "count"),
            nb_inf=("resultat", lambda s: (s == "Infraction").sum()),
        )
        .reset_index()
    )
    agg_pnf["taux_inf"] = agg_pnf["nb_inf"] / agg_pnf["nb_controles"].replace(0, pd.NA)
    agg_pnf.to_csv(out_dir / "indicateurs_chasse_par_pnf.csv", sep=";", index=False)

    # Export des points de contrôle (pour cartes densité)
    cols_points = [
        c
        for c in [
            "fid",
            "dc_id",
            "nom_dossie",
            "date_ctrl",
            "num_depart",
            "nom_commun",
            "insee_comm",
            "dr",
            "entit_ctrl",
            "type_usage",
            "domaine",
            "theme",
            "type_actio",
            "fc_type",
            "resultat",
            "code_pej",
            "code_pa",
            "x",
            "y",
            "PNF",
        ]
        if c in point_chasse.columns
    ]
    point_chasse[cols_points].to_csv(
        out_dir / "points_controles_chasse.csv", sep=";", index=False
    )

    return point_chasse, agg_commune, agg_pnf, tab_resultats


def analyse_pej_pa(
    root: Path, pa: pd.DataFrame, point_chasse: pd.DataFrame, out_dir: Path
) -> None:
    """
    PEJ chasse : même logique de gestion des sources que le bilan agrainage.
    Source = ODS PEJ (load_pej) ; périmètre = département (DC_ID dans contrôles chasse
    ou ENTITE_ORIGINE_PROCEDURE == SDxx) ; déduplication par DC_ID (1 ligne = 1 dossier).
    """
    pej = load_pej(root, date_deb=DATE_DEB, date_fin=DATE_FIN)

    # Restreint PEJ au département : DC_ID dans contrôles chasse OU entité SD21 (comme agrainage)
    dc_ids_dept = (
        set(point_chasse["dc_id"].dropna().unique())
        if not point_chasse.empty and "dc_id" in point_chasse.columns
        else set()
    )
    entity_sd = "SD" + DEPT_CODE
    if "ENTITE_ORIGINE_PROCEDURE" in pej.columns:
        pej_dept = pej[pej["ENTITE_ORIGINE_PROCEDURE"] == entity_sd].copy()
    else:
        pej_dept = pej.copy()

    # Déduplication par DC_ID (1 PEJ = 1 dossier), même logique qu'agrainage
    if "DATE_REF" in pej_dept.columns:
        pej_dept = (
            pej_dept.sort_values("DATE_REF", ascending=False)
            .drop_duplicates(subset="DC_ID", keep="first")
            .copy()
        )
    else:
        pej_dept = pej_dept.drop_duplicates(subset="DC_ID", keep="first").copy()

    # Filtre PEJ chasse par liste NATINF (plan : filtre NATINF seul)
    _natinf_pattern = "|".join(rf"(?:^|_){re.escape(c)}(?:_|$)" for c in NATINF_PEJ_CHASSE)
    pej_chasse = pej_dept[
        pej_dept["NATINF_PEJ"].fillna("").astype(str).str.contains(_natinf_pattern, regex=True)
    ].copy()

    pej_par_theme = (
        pej_chasse.groupby(["DOMAINE", "THEME"])
        .size()
        .rename("nb_pej")
        .reset_index()
    )
    pej_par_theme.to_csv(
        out_dir / "pej_chasse_par_theme.csv", sep=";", index=False
    )

    # Volume global PEJ chasse : 1 ligne = 1 dossier (déjà dédupliqué)
    nb_pej_chasse = len(pej_chasse)
    pd.DataFrame(
        [{"nb_pej_chasse": nb_pej_chasse}]
    ).to_csv(out_dir / "pej_chasse_resume.csv", sep=";", index=False)

    # Export liste des PEJ chasse (traçabilité, comme pej_agrainage.csv)
    cols_pej = [
        c for c in [
            "DC_ID", "DATE_REF", "NATINF_PEJ", "DOMAINE", "THEME",
            "TYPE_ACTION", "DUREE_PEJ", "CLOTUR_PEJ", "SUITE",
            "ENTITE_ORIGINE_PROCEDURE",
        ] if c in pej_chasse.columns
    ]
    if cols_pej:
        pej_chasse[cols_pej].to_csv(
            out_dir / "pej_chasse.csv", sep=";", index=False
        )

    # PA : restreint au département (même logique que PEJ)
    pa_mask = pa["DC_ID"].isin(dc_ids_dept)
    if "ENTITE_ORIGINE_PROCEDURE" in pa.columns:
        pa_mask = pa_mask | (pa["ENTITE_ORIGINE_PROCEDURE"] == entity_sd)
    pa_dept = pa[pa_mask].copy()

    theme_pa = pa_dept.get("THEME", pd.Series(dtype=object)).fillna("").astype(str).str.lower()
    type_act_pa = pa_dept.get("TYPE_ACTION", pd.Series(dtype=object)).fillna("").astype(str).str.lower()
    mask_pa_chasse = (
        theme_pa.str.contains("chasse", regex=False)
        | type_act_pa.str.contains("chasse", regex=False)
        | theme_pa.str.contains("police de la chasse", regex=False)
    )
    pa_chasse = pa_dept[mask_pa_chasse].copy()

    pa_par_theme = (
        pa_chasse.groupby(["DOMAINE", "THEME"])
        .size()
        .rename("nb_pa")
        .reset_index()
    )
    pa_par_theme.to_csv(
        out_dir / "pa_chasse_par_theme.csv", sep=";", index=False
    )

    # Volume global de PA chasse
    id_col = "DC_ID" if "DC_ID" in pa_chasse.columns else None
    if id_col:
        nb_pa_chasse = pa_chasse[id_col].nunique()
    else:
        nb_pa_chasse = len(pa_chasse)

    pd.DataFrame(
        [{"nb_pa_chasse": nb_pa_chasse}]
    ).to_csv(out_dir / "pa_chasse_resume.csv", sep=";", index=False)


def analyse_pve_chasse(root: Path, tub: pd.DataFrame, pnf: pd.DataFrame, out_dir: Path) -> None:
    """PVe chasse : filtrage par NATINF_PVE_CHASSE, agrégation par zone (département, TUB, PNF)."""
    tub_codes = set(tub["INSEE_COM"].astype(str).str.zfill(5).unique())
    pnf_codes = set(pnf["CODE_INSEE"].astype(str).str.zfill(5).unique())

    pve = load_pve(root, dept_code=DEPT_CODE, date_deb=DATE_DEB, date_fin=DATE_FIN)
    if "INF-NATINF" not in pve.columns:
        pve_chasse = pve.iloc[0:0].copy()
    else:
        pve_p = pve.copy()
        pve_p["is_chasse"] = pve_p["INF-NATINF"].apply(
            lambda x: contient_natinf(x, NATINF_PVE_CHASSE)
        )
        pve_chasse = pve_p[pve_p["is_chasse"]].copy()

    nb_dept = len(pve_chasse)
    if "INF-INSEE" in pve_chasse.columns:
        pve_chasse["INF-INSEE"] = pve_chasse["INF-INSEE"].astype(str).str.zfill(5)
        nb_tub = int(pve_chasse["INF-INSEE"].isin(tub_codes).sum())
        nb_pnf = int(pve_chasse["INF-INSEE"].isin(pnf_codes).sum())
    else:
        nb_tub = 0
        nb_pnf = 0

    zone_pve = pd.DataFrame([
        {"zone": "Département", "nb": nb_dept},
        {"zone": "Zone TUB", "nb": nb_tub},
        {"zone": "PNF", "nb": nb_pnf},
    ])
    zone_pve.to_csv(out_dir / "pve_chasse_par_zone.csv", sep=";", index=False)
    pd.DataFrame([{"nb_pve_chasse": nb_dept}]).to_csv(out_dir / "pve_chasse_resume.csv", sep=";", index=False)

    cols_pev = [
        c for c in [
            "INF-ID", "INF-DATE-INTG", "INF-NATINF", "INF-TYP-INF-STAT-LIB",
            "INF-DEPARTEMENT", "INF-INSEE", "INF-CP",
        ] if c in pve_chasse.columns
    ]
    if cols_pev:
        pve_chasse[cols_pev].to_csv(out_dir / "pve_chasse.csv", sep=";", index=False)


# ---------------------------------------------------------------------------
# Génération du PDF
# ---------------------------------------------------------------------------

def generate_pdf_report(root: Path, out_dir: Path) -> None:
    """Génère le bilan PDF complet avec reportlab + matplotlib."""

    styles = _get_styles()
    dept_name = get_dept_name(DEPT_CODE)
    tmp_dir = Path(tempfile.mkdtemp(prefix="ofb_charts_"))
    try:
        _generate_pdf_content(root, out_dir, tmp_dir, styles)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _generate_pdf_content(root: Path, out_dir: Path, tmp_dir: Path, styles) -> None:
    """Corps de la génération PDF, séparé pour garantir le nettoyage de tmp_dir."""

    # ── Chargement des données ─────────────────────────────────────────
    tab_resultats = pd.read_csv(out_dir / "controles_chasse_resultats_dc.csv", sep=";")
    agg_commune = pd.read_csv(out_dir / "indicateurs_chasse_par_commune.csv", sep=";")
    agg_pnf = pd.read_csv(out_dir / "indicateurs_chasse_par_pnf.csv", sep=";")
    points_controles = pd.read_csv(out_dir / "points_controles_chasse.csv", sep=";")

    pej_par_theme = _load_csv_opt(out_dir, "pej_chasse_par_theme.csv")
    pej_chasse_resume = _load_csv_opt(out_dir, "pej_chasse_resume.csv")
    pa_par_theme = _load_csv_opt(out_dir, "pa_chasse_par_theme.csv")
    pa_chasse_resume = _load_csv_opt(out_dir, "pa_chasse_resume.csv")
    pve_chasse_resume = _load_csv_opt(out_dir, "pve_chasse_resume.csv")
    pve_chasse_par_zone = _load_csv_opt(out_dir, "pve_chasse_par_zone.csv")

    pdf_path = out_dir / "bilan_chasse_Cote_dOr.pdf"

    # Libellés NATINF (chasse) pour les textes descriptifs.
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
                # On tronque légèrement les libellés pour garder un texte lisible.
                parts.append(f"{c} – {lib.strip()[:80]}")
            else:
                parts.append(str(c))
        return "; ".join(parts)

    natinf_pve_txt = _format_natinf_list(NATINF_PVE_CHASSE)
    natinf_pej_txt = _format_natinf_list(NATINF_PEJ_CHASSE)

    # ── Compteurs clés ─────────────────────────────────────────────────
    nb_controles = len(points_controles)
    nb_inf = int(tab_resultats.loc[tab_resultats["resultat"] == "Infraction", "nb"].sum()) if "Infraction" in tab_resultats["resultat"].values else 0
    taux_inf = nb_inf / nb_controles if nb_controles else 0

    nb_pej_total = (
        int(pej_chasse_resume["nb_pej_chasse"].iloc[0])
        if pej_chasse_resume is not None and not pej_chasse_resume.empty
        else (int(pej_par_theme["nb_pej"].sum()) if pej_par_theme is not None and not pej_par_theme.empty else 0)
    )
    nb_pa_total = int(pa_chasse_resume["nb_pa_chasse"].iloc[0]) if pa_chasse_resume is not None and not pa_chasse_resume.empty else 0

    # ── Sections du sommaire (ancre, titre) ────────────────────────────
    sections = [
        ("sec1", "I. Localisations de contrôle chasse"),
        ("sec_pve", "II. PVe chasse"),
        ("sec2", "III. Procédures judiciaires et administratives"),
        ("sec3", "IV. PNF / Hors PNF"),
        ("sec4", "V. Cartographie"),
        ("sec5", "Annexes"),
    ]

    # ── Construction du document ─────────────────────────────────────
    content_frame = Frame(
        MARGIN_LEFT, MARGIN_BOTTOM,
        PAGE_W - MARGIN_LEFT - MARGIN_RIGHT,
        PAGE_H - MARGIN_TOP - MARGIN_BOTTOM,
        id="content",
    )

    def _header_footer(canvas, doc):
        canvas.saveState()
        # Image décorative bas-droit (dessinée en premier, sous le texte)
        if IMG_FOOTER_DECO.exists():
            canvas.drawImage(
                str(IMG_FOOTER_DECO),
                PAGE_W - 60 * mm, 0,
                width=60 * mm, height=7 * mm,
                preserveAspectRatio=True, mask="auto",
            )
        # En-tête : filet bleu
        canvas.setStrokeColor(rl_colors.HexColor(COLOR_PRIMARY))
        canvas.setLineWidth(2)
        y_header = PAGE_H - 16 * mm
        canvas.line(MARGIN_LEFT, y_header, PAGE_W - MARGIN_RIGHT, y_header)
        canvas.setFont(f"{FONT_FAMILY}-Bold", 8)
        canvas.setFillColor(rl_colors.HexColor(COLOR_PRIMARY))
        canvas.drawString(MARGIN_LEFT, y_header + 3, f"Bilan chasse \u2013 {dept_name}")

        # Pied de page
        y_foot = 8 * mm
        canvas.setFont(f"{FONT_FAMILY}-Bold", 7)
        canvas.setFillColor(rl_colors.HexColor(COLOR_SECONDARY))
        canvas.drawString(MARGIN_LEFT, y_foot + 12, "Office fran\u00e7ais de la biodiversit\u00e9")
        canvas.setFont(FONT_FAMILY, 7)
        canvas.drawString(MARGIN_LEFT, y_foot + 3, f"SD {dept_name} \u2013 www.ofb.gouv.fr")
        canvas.drawRightString(PAGE_W - MARGIN_RIGHT, y_foot + 3, f"{doc.page}")
        canvas.restoreState()

    def _title_page_template(canvas, doc):
        """Page de garde : fond, bandeau logo, titre."""
        canvas.saveState()
        if IMG_BACKGROUND.exists():
            canvas.drawImage(
                str(IMG_BACKGROUND), 0, 0,
                width=PAGE_W, height=PAGE_H * 0.86,
                preserveAspectRatio=False, mask="auto",
            )
        if IMG_LOGO_BANNER.exists():
            canvas.drawImage(
                str(IMG_LOGO_BANNER), 0, PAGE_H * 0.86,
                width=PAGE_W, height=PAGE_H * 0.14,
                preserveAspectRatio=False, mask="auto",
            )
        cx = PAGE_W / 2
        canvas.setFont(f"{FONT_FAMILY}-Bold", 26)
        canvas.setFillColor(rl_colors.HexColor(COLOR_PRIMARY))
        canvas.drawCentredString(cx, PAGE_H * 0.62 + 14, "Bilan de l\u2019activit\u00e9 de contr\u00f4le")
        canvas.drawCentredString(cx, PAGE_H * 0.62 - 20, "Chasse")
        canvas.setFont(f"{FONT_FAMILY}-Bold", 22)
        canvas.drawCentredString(cx, PAGE_H * 0.50, dept_name)
        canvas.setFont(FONT_FAMILY, 14)
        canvas.setFillColor(rl_colors.HexColor(COLOR_GREY))
        canvas.drawCentredString(
            cx, PAGE_H * 0.42,
            f"P\u00e9riode : {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y}",
        )
        canvas.setFont(FONT_FAMILY, 11)
        canvas.drawCentredString(cx, PAGE_H * 0.36, "Donn\u00e9es OSCEAN")
        # Footer sur page de garde aussi
        canvas.setFont(f"{FONT_FAMILY}-Bold", 7)
        canvas.setFillColor(rl_colors.HexColor(COLOR_SECONDARY))
        canvas.drawString(MARGIN_LEFT, 30, "Office fran\u00e7ais de la biodiversit\u00e9")
        canvas.setFont(FONT_FAMILY, 7)
        canvas.drawString(MARGIN_LEFT, 20, f"SD {dept_name} \u2013 www.ofb.gouv.fr")
        canvas.restoreState()

    title_template = PageTemplate(
        id="TitlePage",
        frames=[Frame(0, 0, PAGE_W, PAGE_H, id="full")],
        onPage=_title_page_template,
    )
    normal_template = PageTemplate(
        id="Normal",
        frames=[content_frame],
        onPage=_header_footer,
    )

    doc = BaseDocTemplate(
        str(pdf_path), pagesize=A4,
        title=f"Bilan chasse \u2013 {dept_name}",
        author="Office fran\u00e7ais de la biodiversit\u00e9",
    )
    doc.addPageTemplates([title_template, normal_template])

    story = []
    avail_w = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT

    # ═══════════ PAGE DE GARDE ═══════════
    story.append(NextPageTemplate("Normal"))
    story.append(PageBreak())

    # ═══════════ SOMMAIRE ═══════════
    story.append(Paragraph("Sommaire", styles["Title"]))
    story.append(Spacer(1, 6 * mm))
    for anchor, sec_title in sections:
        story.append(Paragraph(
            f'<a href="#{anchor}" color="{COLOR_PRIMARY}">{sec_title}</a>',
            styles["TOCEntry"],
        ))
    story.append(PageBreak())

    # ═══════════ SECTION 1 : CONTRÔLES CHASSE ═══════════
    story.append(Paragraph(f'<a name="sec1"/>I. Localisations de contr\u00f4le chasse', styles["Heading1"]))

    story.append(key_figures_table([
        (str(nb_controles), "Localisations de contr\u00f4le"),
        (str(nb_inf), "Infractions"),
        (f"{taux_inf:.1%}", "Taux d\u2019infraction"),
    ], styles))
    story.append(Spacer(1, 5 * mm))

    # Tableau résultats
    tbl_data = [["R\u00e9sultat", "Nombre", "Taux"]]
    for _, row in tab_resultats.iterrows():
        taux_str = f"{row['taux']:.1%}" if pd.notna(row.get("taux")) else "n.d."
        tbl_data.append([str(row["resultat"]), str(int(row["nb"])), taux_str])
    _desc_parts = []
    for _, _r in tab_resultats.iterrows():
        _t = f"{_r['taux']:.1%}" if pd.notna(_r.get("taux")) else "n.d."
        _desc_parts.append(f"{int(_r['nb'])} {_r['resultat'].lower()} ({_t})")
    story.append(KeepTogether([
        Paragraph("Tableau 1 : R\u00e9sultats des contr\u00f4les chasse", styles["TableCaption"]),
        ofb_table(tbl_data, col_widths=[avail_w * 0.50, avail_w * 0.25, avail_w * 0.25],
                   col_aligns=["LEFT", "RIGHT", "RIGHT"]),
        Paragraph(
            f"Sur la p\u00e9riode, {nb_controles} contr\u00f4le(s) chasse ont \u00e9t\u00e9 "
            f"enregistr\u00e9s. Les r\u00e9sultats se r\u00e9partissent comme suit : "
            + ", ".join(_desc_parts) + ".",
            styles["BodyText"],
        ),
    ]))
    story.append(Spacer(1, 3 * mm))

    # Camembert résultats
    pie_data = {}
    for _, row in tab_resultats.iterrows():
        pie_data[str(row["resultat"])] = int(row["nb"])
    if pie_data:
        pie_path = chart_pie(pie_data, "R\u00e9partition des r\u00e9sultats", tmp_dir, "pie_resultats.png")
        _pimg = PILImage.open(pie_path)
        _pw, _ph = _pimg.size
        _pimg.close()
        _target_w = avail_w * 0.75
        _target_h = _target_w * (_ph / _pw)
        img_pie = RLImage(pie_path, width=_target_w, height=_target_h)
        img_pie.hAlign = "CENTER"
        story.append(img_pie)
    story.append(Spacer(1, 3 * mm))

    # Correspondance code INSEE → nom de commune pour le tableau PDF
    insee_to_nom = load_communes_noms(PROJECT_ROOT)

    def _nom_commune(code):
        return insee_to_nom.get(str(code).strip().zfill(5), str(code))

    # Top communes
    top_communes = agg_commune.sort_values("nb_controles", ascending=False).head(10)
    tbl_comm = [["Commune", "Nb contr\u00f4les", "Nb infractions", "Taux infraction"]]
    for _, row in top_communes.iterrows():
        taux_c = f"{row['taux_infraction']:.1%}" if pd.notna(row.get("taux_infraction")) else "n.d."
        tbl_comm.append([
            _nom_commune(row["insee_comm"]),
            str(int(row["nb_controles"])),
            str(int(row["nb_infractions"])),
            taux_c,
        ])
    _top1 = top_communes.iloc[0] if len(top_communes) > 0 else None
    _nb_comm = agg_commune["insee_comm"].nunique()
    story.append(KeepTogether([
        Paragraph("Tableau 2 : Communes avec le plus de contr\u00f4les", styles["TableCaption"]),
        ofb_table(tbl_comm,
                   col_widths=[avail_w * 0.25, avail_w * 0.25, avail_w * 0.25, avail_w * 0.25],
                   col_aligns=["LEFT", "RIGHT", "RIGHT", "RIGHT"]),
        Paragraph(
            f"Les contr\u00f4les chasse ont concern\u00e9 {_nb_comm} communes du d\u00e9partement. "
            + (f"La commune la plus contr\u00f4l\u00e9e ({_nom_commune(_top1['insee_comm'])}) "
               f"comptabilise {int(_top1['nb_controles'])} contr\u00f4le(s)." if _top1 is not None else ""),
            styles["BodyText"],
        ),
        Paragraph(
            f"Source : OSCEAN, contr\u00f4les chasse du {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y}, d\u00e9partement {DEPT_CODE}.",
            styles["BodySmall"],
        ),
    ]))
    story.append(PageBreak())

    # ═══════════ SECTION II : PVe CHASSE ═══════════
    nb_pve_chasse = (
        int(pve_chasse_resume["nb_pve_chasse"].iloc[0])
        if pve_chasse_resume is not None and not pve_chasse_resume.empty
        else 0
    )
    story.append(Paragraph(f'<a name="sec_pve"/>II. PVe chasse', styles["Heading1"]))
    story.append(Paragraph(
        "Procès-verbaux d\u2019infraction (PVe) chasse pour les NATINF suivants : "
        f"{natinf_pve_txt}.",
        styles["BodyText"],
    ))
    story.append(Spacer(1, 3 * mm))
    story.append(key_figures_table([(str(nb_pve_chasse), "PVe chasse")], styles))
    story.append(Spacer(1, 3 * mm))
    if pve_chasse_par_zone is not None and not pve_chasse_par_zone.empty:
        tbl_pve = [["Zone", "Nombre"]]
        for _, row in pve_chasse_par_zone.iterrows():
            tbl_pve.append([str(row["zone"]), str(int(row["nb"]))])
        story.append(KeepTogether([
            Paragraph("Tableau 3 : PVe chasse par zone", styles["TableCaption"]),
            ofb_table(tbl_pve, col_widths=[avail_w * 0.60, avail_w * 0.40], col_aligns=["LEFT", "RIGHT"]),
        ]))
        story.append(Spacer(1, 3 * mm))
    story.append(PageBreak())

    # ═══════════ SECTION 2 : PEJ / PA ═══════════
    story.append(Paragraph(f'<a name="sec2"/>III. Proc\u00e9dures judiciaires et administratives', styles["Heading1"]))

    story.append(key_figures_table([
        (str(nb_pej_total), "PEJ chasse"),
        (str(nb_pa_total), "PA chasse"),
    ], styles))
    story.append(Spacer(1, 5 * mm))

        if pej_par_theme is not None and not pej_par_theme.empty:
        tbl_pej = [["Domaine", "Th\u00e8me", "Nombre"]]
        for _, row in pej_par_theme.iterrows():
            tbl_pej.append([str(row.get("DOMAINE", "")), str(row.get("THEME", "")), str(int(row["nb_pej"]))])
        tbl_pej.append(["", "Total", str(nb_pej_total)])
            story.append(KeepTogether([
            Paragraph("Tableau 4 : PEJ chasse par th\u00e8me", styles["TableCaption"]),
            ofb_table(tbl_pej,
                       col_widths=[avail_w * 0.35, avail_w * 0.45, avail_w * 0.20],
                       col_aligns=["LEFT", "LEFT", "RIGHT"]),
            Paragraph(
                f"Au total, {nb_pej_total} proc\u00e9dure(s) d\u2019enqu\u00eate judiciaire "
                f"(PEJ) ont \u00e9t\u00e9 enregistr\u00e9es sur la th\u00e9matique chasse "
                f"pour les NATINF suivants : {natinf_pej_txt}.",
                styles["BodyText"],
            ),
        ]))
        story.append(Spacer(1, 3 * mm))

    if pa_par_theme is not None and not pa_par_theme.empty:
        tbl_pa = [["Domaine", "Th\u00e8me", "Nombre"]]
        for _, row in pa_par_theme.iterrows():
            tbl_pa.append([str(row.get("DOMAINE", "")), str(row.get("THEME", "")), str(int(row["nb_pa"]))])
        tbl_pa.append(["", "Total", str(nb_pa_total)])
        story.append(KeepTogether([
            Paragraph("Tableau 5 : PA chasse par th\u00e8me", styles["TableCaption"]),
            ofb_table(tbl_pa,
                       col_widths=[avail_w * 0.35, avail_w * 0.45, avail_w * 0.20],
                       col_aligns=["LEFT", "LEFT", "RIGHT"]),
            Paragraph(
                f"{nb_pa_total} proc\u00e9dure(s) administrative(s) chasse "
                f"ont \u00e9t\u00e9 engag\u00e9es sur la p\u00e9riode.",
                styles["BodyText"],
            ),
        ]))
        story.append(Spacer(1, 3 * mm))

    # Histogramme PEJ vs PA
    if nb_pej_total or nb_pa_total:
        bar_path = chart_bar(
            ["PEJ", "PA"], [nb_pej_total, nb_pa_total],
            "Proc\u00e9dures chasse : PEJ vs PA", "Nombre",
            tmp_dir, "bar_pej_pa.png",
        )
        img_bar = RLImage(bar_path, width=avail_w * 0.6, height=avail_w * 0.36)
        img_bar.hAlign = "CENTER"
        story.append(img_bar)
    story.append(PageBreak())

    # ═══════════ SECTION 3 : PNF / HORS PNF ═══════════
    story.append(Paragraph(f'<a name="sec3"/>IV. PNF / Hors PNF', styles["Heading1"]))
    story.append(Paragraph(
        "Comparaison des contr\u00f4les chasse entre les communes situ\u00e9es dans le Parc national "
        "de for\u00eats (PNF) et celles hors PNF.",
        styles["BodyText"],
    ))
    story.append(Spacer(1, 3 * mm))

    # Tableau PNF chasse
    if agg_pnf is not None and not agg_pnf.empty:
        tbl_pnf = [["Zone", "Nb contr\u00f4les", "Nb infractions", "Taux infraction"]]
        grp_labels = []
        series_ctrl = {"Localisations de contr\u00f4le": [], "Infractions": []}
        for _, row in agg_pnf.iterrows():
            taux_str = f"{row['taux_inf']:.1%}" if pd.notna(row.get("taux_inf")) else "n.d."
            tbl_pnf.append([
                str(row["PNF"]), str(int(row["nb_controles"])), str(int(row["nb_inf"])), taux_str,
            ])
            grp_labels.append(str(row["PNF"]))
            series_ctrl["Localisations de contr\u00f4le"].append(int(row["nb_controles"]))
            series_ctrl["Infractions"].append(int(row["nb_inf"]))
        _pnf_rows = agg_pnf.to_dict("records")
        _pnf_desc_parts = []
        for _pr in _pnf_rows:
            _t = f"{_pr['taux_inf']:.1%}" if pd.notna(_pr.get("taux_inf")) else "n.d."
            _pnf_desc_parts.append(
                f"{_pr['PNF']} : {int(_pr['nb_controles'])} contr\u00f4les, "
                f"{int(_pr['nb_inf'])} infraction(s) (taux {_t})"
            )
        story.append(KeepTogether([
            Paragraph("Tableau 6 : Contr\u00f4les chasse \u2013 PNF vs Hors PNF", styles["TableCaption"]),
            ofb_table(tbl_pnf,
                       col_widths=[avail_w * 0.30, avail_w * 0.23, avail_w * 0.23, avail_w * 0.24],
                       col_aligns=["LEFT", "RIGHT", "RIGHT", "RIGHT"]),
            Paragraph(
                "R\u00e9partition des contr\u00f4les chasse selon la localisation : "
                + " ; ".join(_pnf_desc_parts) + ".",
                styles["BodyText"],
            ),
        ]))
        story.append(Spacer(1, 3 * mm))

        if grp_labels:
            grp_path = chart_bar_grouped(
                grp_labels, series_ctrl,
                "Localisations de contr\u00f4le chasse : PNF vs Hors PNF", "Nombre",
                tmp_dir, "bar_pnf_chasse.png",
            )
            img_grp = RLImage(grp_path, width=avail_w * 0.6, height=avail_w * 0.42)
            img_grp.hAlign = "CENTER"
            story.append(img_grp)
    story.append(PageBreak())

    # ═══════════ SECTION 4 : CARTOGRAPHIE ═══════════
    story.append(Paragraph(f'<a name="sec4"/>V. Cartographie', styles["Heading1"]))
    story.append(Paragraph(
        f"Carte du d\u00e9partement ({dept_name}) repr\u00e9sentant la r\u00e9partition spatiale des contr\u00f4les "
        "chasse par commune.",
        styles["BodyText"],
    ))
    story.append(Spacer(1, 4 * mm))

    carte_path = get_cartes_dir() / "carte_chasse.png"
    if carte_path.exists():
        story.append(Paragraph(
            f"Contr\u00f4les chasse \u2013 {dept_name}",
            styles["Heading2"],
        ))
        story.append(Spacer(1, 3 * mm))
        _cimg = PILImage.open(str(carte_path))
        _cw, _ch = _cimg.size
        _cimg.close()
        _map_w = avail_w
        _map_h = _map_w * (_ch / _cw)
        img_carte = RLImage(str(carte_path), width=_map_w, height=_map_h)
        img_carte.hAlign = "CENTER"
        story.append(img_carte)
        story.append(Paragraph(
            "Sources : OFB/OSCEAN \u2013 IGN/BD TOPO \u2013 ESRI World Topographic \u2013 MNHN/Espaces prot\u00e9g\u00e9s.",
            styles["BodySmall"],
        ))
    else:
        story.append(Paragraph(
            "<i>Carte non disponible. D\u00e9posez le fichier "
            "<b>carte_chasse.png</b> dans le dossier des cartes pour l\u2019int\u00e9grer au bilan.</i>",
            styles["BodyText"],
        ))
    story.append(PageBreak())

    # ═══════════ ANNEXES ═══════════
    story.append(Paragraph(f'<a name="sec5"/>Annexes', styles["Heading1"]))

    story.append(Paragraph("M\u00e9thodologie", styles["Heading2"]))
    methodo_text = (
        f"<b>P\u00e9riode d\u2019analyse :</b> du {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y}.<br/>"
        f"<b>P\u00e9rim\u00e8tre :</b> d\u00e9partement {dept_name} ({DEPT_CODE}).<br/>"
        "<b>Sources :</b> OSCEAN (points de contr\u00f4le, PEJ, PA), Stats PVe OFB.<br/>"
        "<b>Chasse :</b> contr\u00f4les dont le th\u00e8me ou le type d\u2019action contient \u00ab\u00a0chasse\u00a0\u00bb ; "
        "PEJ chasse filtr\u00e9s par liste NATINF (C1/C3/C4/C5/D\u00e9lit) ; "
        "PVe chasse filtr\u00e9s par liste NATINF (26274, 26301, 27742, 27745, 20148, 20155, 20165, 20166, 2002).<br/>"
        "<b>PNF :</b> communes situ\u00e9es dans le p\u00e9rim\u00e8tre du Parc national de for\u00eats."
    )
    story.append(Paragraph(methodo_text, styles["BodyText"]))
    story.append(Spacer(1, 6 * mm))

    story.append(Paragraph("Glossaire", styles["Heading2"]))
    glossaire = [
        ["Abr\u00e9viation", "Signification"],
        ["DC", "Dossier de contr\u00f4le"],
        ["NATINF", "Nature d\u2019infraction (nomenclature nationale)"],
        ["OSCEAN", "Outil de suivi des contr\u00f4les en environnement (application nationale)"],
        ["PA", "Proc\u00e9dure administrative"],
        ["PEJ", "Proc\u00e9dure d\u2019enqu\u00eate judiciaire"],
        ["PNF", "Parc national de for\u00eats"],
        ["PVe", "Proc\u00e8s-verbal \u00e9lectronique"],
    ]
    story.append(ofb_table(glossaire,
                            col_widths=[avail_w * 0.25, avail_w * 0.75],
                            col_aligns=["LEFT", "LEFT"]))

    # ── Build du PDF ──────────────────────────────────────────────────
    doc.build(story)


def _parse_args() -> argparse.Namespace:
    """Paramètres de ligne de commande pour le bilan chasse."""
    parser = argparse.ArgumentParser(
        description=(
            "Génère le bilan chasse pour un département et une période donnés "
            "à partir des données OSCEAN (points de contrôle) et PEJ/PA."
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
    out_dir = get_out_dir("bilan_chasse")

    print(
        f"Période analysée : du {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y} "
        f"(département {DEPT_CODE})."
    )

    print("Étape 1/5 : chargement des données...")
    with Spinner():
        point = load_point_ctrl(root, dept_code=DEPT_CODE, date_deb=DATE_DEB, date_fin=DATE_FIN)
        pa = load_pa(root, date_deb=DATE_DEB, date_fin=DATE_FIN)
        pnf = load_pnf(root)
        tub = load_tub(root)

    print("Étape 2/5 : analyse des contrôles chasse...")
    with Spinner():
        point_chasse, agg_commune, agg_pnf, tab_resultats = analyse_controles_chasse(
            point, pnf, out_dir
        )

    print("Étape 3/5 : analyse des procédures PEJ/PA chasse...")
    with Spinner():
        analyse_pej_pa(root, pa, point_chasse, out_dir)

    print("Étape 4/5 : analyse PVe chasse...")
    with Spinner():
        analyse_pve_chasse(root, tub, pnf, out_dir)

    print("Étape 5/5 : génération du bilan PDF...")
    with Spinner():
        generate_pdf_report(root, out_dir)

    print("Analyse terminée. Bilan généré dans out/bilan_chasse.")
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
    out_dir = get_out_dir("bilan_chasse")

    print(
        f"Période analysée : du {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y} "
        f"(département {DEPT_CODE})."
    )

    print("Étape 1/5 : chargement des données...")
    with Spinner():
        point = load_point_ctrl(root, dept_code=DEPT_CODE, date_deb=DATE_DEB, date_fin=DATE_FIN)
        pa = load_pa(root, date_deb=DATE_DEB, date_fin=DATE_FIN)
        pnf = load_pnf(root)
        tub = load_tub(root)

    print("Étape 2/5 : analyse des contrôles chasse...")
    with Spinner():
        point_chasse, agg_commune, agg_pnf, tab_resultats = analyse_controles_chasse(
            point, pnf, out_dir
        )

    print("Étape 3/5 : analyse des procédures PEJ/PA chasse...")
    with Spinner():
        analyse_pej_pa(root, pa, point_chasse, out_dir)

    print("Étape 4/5 : analyse PVe chasse...")
    with Spinner():
        analyse_pve_chasse(root, tub, pnf, out_dir)

    print("Étape 5/5 : génération du bilan PDF...")
    with Spinner():
        generate_pdf_report(root, out_dir)

    print("Analyse terminée. Bilan généré dans out/bilan_chasse.")


if __name__ == "__main__":
    main()

