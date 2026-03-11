"""Bilan global — activité du service départemental (tous domaines/thèmes, PA, PEJ, PVe)."""
import argparse
import shutil
import sys
import tempfile
from pathlib import Path

import pandas as pd
from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
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

from scripts.paths import get_cartes_dir, get_out_dir
from scripts.common.loaders import load_pa, load_pej, load_point_ctrl, load_pnf, load_pve
from scripts.common.prompt_periode import ask_periode_dept
from scripts.common.utils import serie_type_usager
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
from scripts.common.pdf_utils import _key_figures_table, _ofb_table
from scripts.common.charts import _chart_bar, _chart_pie

# ---------------------------------------------------------------------------
# Période et paramètres
# ---------------------------------------------------------------------------
DATE_DEB = pd.Timestamp("2025-01-01")
DATE_FIN = pd.Timestamp("2026-02-05")
DEPT_CODE = "21"
ENTITY_SD = "SD21"


def _load_natinf_ref(root: Path) -> pd.DataFrame:
    """Charge le référentiel NATINF (ref/liste_natinf.csv) pour libeller les exports."""
    for base in ("ref", "sources"):
        for name in ("liste_natinf.csv", "liste-natinf-avril2023.csv"):
            path = _ROOT / base / name
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
                lib_col = None
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
                return df[["numero_natinf", "libelle_natinf"]].drop_duplicates()
            except Exception:
                continue
    return pd.DataFrame(columns=["numero_natinf", "libelle_natinf"])


def _load_csv_opt(out_dir: Path, name: str) -> pd.DataFrame | None:
    """Charge un CSV optionnel ; retourne None si absent ou vide."""
    p = out_dir / name
    if not p.exists():
        return None
    try:
        df = pd.read_csv(p, sep=";")
        return df if not df.empty else None
    except Exception:
        return None


def analyse_controles_global(point: pd.DataFrame, out_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Contrôles tous domaines/thèmes (point déjà filtré par le loader sur département et période).
    Produit : effectifs par domaine, par thème, résultats (Conforme/Infraction/Manquement).
    """
    pt = point.copy()
    pt["insee_comm"] = pt["insee_comm"].astype(str).str.zfill(5)

    nb_total = len(pt)

    # Résultats
    col_resultat = "resultat" if "resultat" in pt.columns else None
    if col_resultat:
        tab_resultats = (
            pt[col_resultat]
            .value_counts()
            .rename_axis("resultat")
            .to_frame("nb")
            .reset_index()
        )
        tab_resultats["taux"] = tab_resultats["nb"] / float(nb_total or 1)
        tab_resultats.to_csv(out_dir / "controles_global_resultats.csv", sep=";", index=False)
    else:
        tab_resultats = pd.DataFrame(columns=["resultat", "nb", "taux"])
        tab_resultats.to_csv(out_dir / "controles_global_resultats.csv", sep=";", index=False)

    # Par domaine
    col_domaine = "domaine" if "domaine" in pt.columns else None
    if col_domaine:
        agg_domaine = (
            pt[col_domaine]
            .fillna("Hors domaine")
            .value_counts()
            .rename_axis("domaine")
            .to_frame("nb")
            .reset_index()
        )
        agg_domaine["taux"] = agg_domaine["nb"] / float(nb_total or 1)
        agg_domaine.to_csv(out_dir / "controles_global_par_domaine.csv", sep=";", index=False)
    else:
        agg_domaine = pd.DataFrame(columns=["domaine", "nb", "taux"])
        agg_domaine.to_csv(out_dir / "controles_global_par_domaine.csv", sep=";", index=False)

    # Par thème
    col_theme = "theme" if "theme" in pt.columns else "type_actio"
    if col_theme in pt.columns:
        agg_theme = (
            pt[col_theme]
            .fillna("Hors thème")
            .value_counts()
            .rename_axis("theme")
            .to_frame("nb")
            .reset_index()
        )
        agg_theme["taux"] = agg_theme["nb"] / float(nb_total or 1)
        agg_theme.to_csv(out_dir / "controles_global_par_theme.csv", sep=";", index=False)
    else:
        agg_theme = pd.DataFrame(columns=["theme", "nb", "taux"])
        agg_theme.to_csv(out_dir / "controles_global_par_theme.csv", sep=";", index=False)

    # Par type d'usagers (si disponible)
    if "type_usager" in pt.columns:
        pt["type_usager_dominant"] = serie_type_usager(pt, "point_ctrl", "type_usager")
        agg_usager = (
            pt["type_usager_dominant"]
            .fillna("Autre")
            .value_counts()
            .rename_axis("type_usager")
            .to_frame("nb")
            .reset_index()
        )
        agg_usager["taux"] = agg_usager["nb"] / float(nb_total or 1)
        agg_usager.to_csv(out_dir / "controles_global_par_usager.csv", sep=";", index=False)

        # Tableau croisé Usagers × Domaine
        if col_domaine:
            dom = pt[col_domaine].fillna("Hors domaine").astype(str)
        else:
            dom = pd.Series(["Hors domaine"] * len(pt), index=pt.index, dtype=object)
        cross = pd.crosstab(pt["type_usager_dominant"], dom)
        cross.index.name = "type_usager"
        cross.reset_index().to_csv(out_dir / "controles_global_usager_par_domaine.csv", sep=";", index=False)

        # Indicateur : contrôles multi-usagers (valeur source contient une virgule)
        nb_multi = int(pt["type_usager"].fillna("").astype(str).str.contains(",", regex=False).sum())
        pd.DataFrame([{"nb_controles_multi_usagers": nb_multi}]).to_csv(
            out_dir / "controles_global_usagers_resume.csv", sep=";", index=False
        )
    else:
        pd.DataFrame(columns=["type_usager", "nb", "taux"]).to_csv(
            out_dir / "controles_global_par_usager.csv", sep=";", index=False
        )
        pd.DataFrame(columns=["type_usager"]).to_csv(
            out_dir / "controles_global_usager_par_domaine.csv", sep=";", index=False
        )
        pd.DataFrame([{"nb_controles_multi_usagers": 0}]).to_csv(
            out_dir / "controles_global_usagers_resume.csv", sep=";", index=False
        )

    return tab_resultats, agg_domaine, agg_theme


def analyse_pej_pa_global(
    root: Path,
    point: pd.DataFrame,
    pa: pd.DataFrame,
    pej: pd.DataFrame,
    out_dir: Path,
) -> None:
    """PEJ et PA du département (DC_ID dans contrôles ou ENTITE_ORIGINE_PROCEDURE == SD21), tous domaines/thèmes."""
    natinf_ref = _load_natinf_ref(root)
    dc_ids = set(point["dc_id"].dropna().unique()) if not point.empty and "dc_id" in point.columns else set()

    # PEJ département
    pej_mask = pej["DC_ID"].isin(dc_ids)
    if "ENTITE_ORIGINE_PROCEDURE" in pej.columns:
        pej_mask = pej_mask | (pej["ENTITE_ORIGINE_PROCEDURE"].astype(str).str.strip() == ENTITY_SD)
    pej_dept = pej[pej_mask].copy()
    if "DATE_REF" in pej_dept.columns:
        pej_dept = pej_dept.sort_values("DATE_REF", ascending=False).drop_duplicates(subset="DC_ID", keep="first").copy()
    else:
        pej_dept = pej_dept.drop_duplicates(subset="DC_ID", keep="first").copy()

    pej_par_domaine = (
        pej_dept.groupby(pej_dept.get("DOMAINE", pd.Series(dtype=object)).fillna("Hors domaine"))
        .size()
        .rename("nb_pej")
        .reset_index()
    )
    pej_par_domaine.columns = ["domaine", "nb_pej"]
    pej_par_domaine.to_csv(out_dir / "pej_global_par_domaine.csv", sep=";", index=False)

    pej_par_theme = (
        pej_dept.groupby(pej_dept.get("THEME", pd.Series(dtype=object)).fillna("Hors thème"))
        .size()
        .rename("nb_pej")
        .reset_index()
    )
    pej_par_theme.columns = ["theme", "nb_pej"]
    pej_par_theme.to_csv(out_dir / "pej_global_par_theme.csv", sep=";", index=False)

    # PEJ par NATINF (libellé depuis ref/liste_natinf.csv)
    if "NATINF_PEJ" in pej_dept.columns and not pej_dept.empty:
        codes = (
            pej_dept["NATINF_PEJ"]
            .fillna("")
            .astype(str)
            .str.split("_")
            .explode()
            .str.extract(r"(\d+)", expand=False)
            .dropna()
            .astype(str)
            .str.strip()
        )
        vc = codes.value_counts().rename_axis("numero_natinf").reset_index(name="nb_pej")
        if not natinf_ref.empty:
            vc = vc.merge(natinf_ref, on="numero_natinf", how="left")
        vc.to_csv(out_dir / "pej_global_par_natinf.csv", sep=";", index=False)

    pd.DataFrame([{"nb_pej_global": len(pej_dept)}]).to_csv(out_dir / "pej_global_resume.csv", sep=";", index=False)

    # PA département
    pa_mask = pa["DC_ID"].isin(dc_ids)
    if "ENTITE_ORIGINE_PROCEDURE" in pa.columns:
        pa_mask = pa_mask | (pa["ENTITE_ORIGINE_PROCEDURE"].astype(str).str.strip() == ENTITY_SD)
    pa_dept = pa[pa_mask].copy()

    pa_par_domaine = (
        pa_dept.groupby(pa_dept.get("DOMAINE", pd.Series(dtype=object)).fillna("Hors domaine"))
        .size()
        .rename("nb_pa")
        .reset_index()
    )
    pa_par_domaine.columns = ["domaine", "nb_pa"]
    pa_par_domaine.to_csv(out_dir / "pa_global_par_domaine.csv", sep=";", index=False)

    pa_par_theme = (
        pa_dept.groupby(pa_dept.get("THEME", pd.Series(dtype=object)).fillna("Hors thème"))
        .size()
        .rename("nb_pa")
        .reset_index()
    )
    pa_par_theme.columns = ["theme", "nb_pa"]
    pa_par_theme.to_csv(out_dir / "pa_global_par_theme.csv", sep=";", index=False)

    nb_pa = pa_dept["DC_ID"].nunique() if "DC_ID" in pa_dept.columns else len(pa_dept)
    pd.DataFrame([{"nb_pa_global": nb_pa}]).to_csv(out_dir / "pa_global_resume.csv", sep=";", index=False)


def analyse_pve_global(pve: pd.DataFrame, out_dir: Path) -> None:
    """PVe du département, tous NATINF."""
    nb_pve = len(pve)
    pd.DataFrame([{"nb_pve_global": nb_pve}]).to_csv(out_dir / "pve_global_resume.csv", sep=";", index=False)
    if "INF-NATINF" in pve.columns:
        pve_par_natinf = (
            pve["INF-NATINF"]
            .astype(str)
            .value_counts()
            .rename_axis("natinf")
            .to_frame("nb")
            .reset_index()
        )
        natinf_ref = _load_natinf_ref(_ROOT)
        if not natinf_ref.empty:
            pve_par_natinf["numero_natinf"] = pve_par_natinf["natinf"].astype(str).str.extract(r"(\d+)", expand=False)
            pve_par_natinf = pve_par_natinf.merge(natinf_ref, on="numero_natinf", how="left")
        pve_par_natinf.to_csv(out_dir / "pve_global_par_natinf.csv", sep=";", index=False)


def generate_pdf_report(out_dir: Path) -> None:
    """Génère le PDF du bilan global (page de garde, sommaire, chiffres clés, contrôles par domaine/thème, résultats, PEJ/PA/PVe)."""
    styles = _get_styles()
    tmp_dir = Path(tempfile.mkdtemp(prefix="ofb_global_"))
    avail_w = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT

    tab_resultats = _load_csv_opt(out_dir, "controles_global_resultats.csv")
    agg_domaine = _load_csv_opt(out_dir, "controles_global_par_domaine.csv")
    agg_theme = _load_csv_opt(out_dir, "controles_global_par_theme.csv")
    agg_usager = _load_csv_opt(out_dir, "controles_global_par_usager.csv")
    cross_usager_dom = _load_csv_opt(out_dir, "controles_global_usager_par_domaine.csv")
    usagers_resume = _load_csv_opt(out_dir, "controles_global_usagers_resume.csv")
    pej_resume = _load_csv_opt(out_dir, "pej_global_resume.csv")
    pa_resume = _load_csv_opt(out_dir, "pa_global_resume.csv")
    pve_resume = _load_csv_opt(out_dir, "pve_global_resume.csv")

    nb_ctrl = 0
    if agg_domaine is not None and not agg_domaine.empty:
        nb_ctrl = int(agg_domaine["nb"].sum())
    nb_pej = int(pej_resume["nb_pej_global"].iloc[0]) if pej_resume is not None and not pej_resume.empty else 0
    nb_pa = int(pa_resume["nb_pa_global"].iloc[0]) if pa_resume is not None and not pa_resume.empty else 0
    nb_pve = int(pve_resume["nb_pve_global"].iloc[0]) if pve_resume is not None and not pve_resume.empty else 0

    pdf_path = out_dir / "bilan_global_Cote_dOr.pdf"

    sections = [
        ("sec1", "I. Chiffres clés"),
        ("sec2", "II. Contrôles par domaine"),
        ("sec3", "III. Contrôles par thème"),
        ("sec4", "IV. Résultats des contrôles"),
        ("sec5", "V. Procédures (PEJ, PA, PVe)"),
        ("sec6", "VI. Types d’usagers"),
        ("sec7", "Annexes"),
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
        canvas.drawString(MARGIN_LEFT, y_header + 3, "Bilan activité SD – Côte-d'Or")
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
        canvas.drawCentredString(cx, PAGE_H * 0.62 + 14, "Bilan de l'activité du service départemental")
        canvas.drawCentredString(cx, PAGE_H * 0.62 - 20, "Côte-d'Or")
        canvas.setFont(FONT_FAMILY, 14)
        canvas.setFillColor(rl_colors.HexColor(COLOR_GREY))
        canvas.drawCentredString(cx, PAGE_H * 0.50, f"Période : {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y}")
        canvas.setFont(FONT_FAMILY, 11)
        canvas.drawCentredString(cx, PAGE_H * 0.42, "Tous domaines et thèmes – PA, PEJ, PVe")
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
        title="Bilan activité SD – Côte-d'Or",
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

    # I. Chiffres clés
    story.append(Paragraph('<a name="sec1"/>I. Chiffres clés', styles["Heading1"]))
    story.append(_key_figures_table([
        (str(nb_ctrl), "Contrôles"),
        (str(nb_pej), "PEJ"),
        (str(nb_pa), "PA"),
        (str(nb_pve), "PVe"),
    ], styles))
    story.append(Spacer(1, 8 * mm))

    # II. Contrôles par domaine
    story.append(Paragraph('<a name="sec2"/>II. Contrôles par domaine', styles["Heading1"]))
    if agg_domaine is not None and not agg_domaine.empty:
        story.append(Paragraph("Tableau 1 : Contrôles par domaine", styles["TableCaption"]))
        tbl = [["Domaine", "Nombre", "Taux"]]
        for _, row in agg_domaine.head(25).iterrows():
            taux_str = f"{row['taux']:.1%}" if pd.notna(row.get("taux")) else "n.d."
            tbl.append([str(row["domaine"])[:50], str(int(row["nb"])), taux_str])
        story.append(_ofb_table(tbl, col_widths=[avail_w * 0.55, avail_w * 0.22, avail_w * 0.23], col_aligns=["LEFT", "RIGHT", "RIGHT"]))
        if len(agg_domaine) > 25:
            story.append(Paragraph(f"... et {len(agg_domaine) - 25} autres domaines.", styles["BodySmall"]))
        # Graphique
        top_dom = agg_domaine.head(12)
        if not top_dom.empty:
            pie_data = {str(row["domaine"])[:30]: int(row["nb"]) for _, row in top_dom.iterrows()}
            if pie_data:
                pie_path = _chart_pie(pie_data, "Contrôles par domaine (top 12)", tmp_dir, "pie_domaine.png")
                from PIL import Image as PILImage
                from reportlab.platypus import Image as RLImage
                _pimg = PILImage.open(pie_path)
                _target_w = avail_w * 0.75
                _target_h = _target_w * (_pimg.height / _pimg.width)
                _pimg.close()
                story.append(RLImage(pie_path, width=_target_w, height=_target_h))
    else:
        story.append(Paragraph("Aucune donnée domaine disponible.", styles["BodyText"]))
    story.append(Spacer(1, 8 * mm))

    # III. Contrôles par thème
    story.append(Paragraph('<a name="sec3"/>III. Contrôles par thème', styles["Heading1"]))
    if agg_theme is not None and not agg_theme.empty:
        story.append(Paragraph("Tableau 2 : Contrôles par thème (extrait)", styles["TableCaption"]))
        tbl = [["Thème", "Nombre", "Taux"]]
        for _, row in agg_theme.head(20).iterrows():
            taux_str = f"{row['taux']:.1%}" if pd.notna(row.get("taux")) else "n.d."
            tbl.append([str(row["theme"])[:45], str(int(row["nb"])), taux_str])
        story.append(_ofb_table(tbl, col_widths=[avail_w * 0.55, avail_w * 0.22, avail_w * 0.23], col_aligns=["LEFT", "RIGHT", "RIGHT"]))
    else:
        story.append(Paragraph("Aucune donnée thème disponible.", styles["BodyText"]))
    story.append(Spacer(1, 8 * mm))

    # IV. Résultats des contrôles
    story.append(Paragraph('<a name="sec4"/>IV. Résultats des contrôles', styles["Heading1"]))
    if tab_resultats is not None and not tab_resultats.empty:
        story.append(Paragraph("Tableau 3 : Résultats (Conforme / Infraction / Manquement)", styles["TableCaption"]))
        tbl = [["Résultat", "Nombre", "Taux"]]
        for _, row in tab_resultats.iterrows():
            taux_str = f"{row['taux']:.1%}" if pd.notna(row.get("taux")) else "n.d."
            tbl.append([str(row["resultat"]), str(int(row["nb"])), taux_str])
        story.append(_ofb_table(tbl, col_widths=[avail_w * 0.50, avail_w * 0.25, avail_w * 0.25], col_aligns=["LEFT", "RIGHT", "RIGHT"]))
    else:
        story.append(Paragraph("Aucune donnée de résultat disponible.", styles["BodyText"]))
    story.append(Spacer(1, 8 * mm))

    # V. Procédures
    story.append(Paragraph('<a name="sec5"/>V. Procédures (PEJ, PA, PVe)', styles["Heading1"]))
    story.append(Paragraph(
        f"Sur la période : {nb_pej} procédure(s) d'enquête judiciaire (PEJ), "
        f"{nb_pa} procédure(s) administrative(s) (PA), {nb_pve} procès-verbal(aux) électronique(s) (PVe).",
        styles["BodyText"],
    ))
    pej_dom = _load_csv_opt(out_dir, "pej_global_par_domaine.csv")
    if pej_dom is not None and not pej_dom.empty:
        story.append(Paragraph("Tableau 4 : PEJ par domaine", styles["TableCaption"]))
        tbl = [["Domaine", "Nombre PEJ"]]
        for _, row in pej_dom.head(15).iterrows():
            tbl.append([str(row["domaine"])[:50], str(int(row["nb_pej"]))])
        story.append(_ofb_table(tbl, col_widths=[avail_w * 0.60, avail_w * 0.40], col_aligns=["LEFT", "RIGHT"]))

    story.append(PageBreak())

    # VI. Types d'usagers
    story.append(Paragraph('<a name="sec6"/>VI. Types d’usagers', styles["Heading1"]))
    if agg_usager is None or agg_usager.empty:
        story.append(Paragraph(
            "Aucune donnée « type d’usagers » n’est disponible dans les points de contrôle OSCEAN pour la période.",
            styles["BodyText"],
        ))
    else:
        nb_multi = (
            int(usagers_resume["nb_controles_multi_usagers"].iloc[0])
            if usagers_resume is not None and not usagers_resume.empty and "nb_controles_multi_usagers" in usagers_resume.columns
            else 0
        )
        story.append(Paragraph(
            "Répartition des contrôles par type d’usagers (catégorie dominante par contrôle).",
            styles["BodyText"],
        ))
        story.append(_key_figures_table([
            (str(nb_multi), "Contrôles multi-usagers"),
        ], styles))
        story.append(Spacer(1, 5 * mm))

        # Tableau distribution
        story.append(Paragraph("Tableau : Contrôles par type d’usagers", styles["TableCaption"]))
        tbl_u = [["Type d’usagers", "Nombre", "Taux"]]
        for _, row in agg_usager.iterrows():
            taux_str = f"{float(row['taux']):.1%}" if pd.notna(row.get("taux")) else "n.d."
            tbl_u.append([str(row["type_usager"]), str(int(row["nb"])), taux_str])
        story.append(_ofb_table(
            tbl_u,
            col_widths=[avail_w * 0.58, avail_w * 0.21, avail_w * 0.21],
            col_aligns=["LEFT", "RIGHT", "RIGHT"],
        ))
        story.append(Spacer(1, 5 * mm))

        # Graphique
        pie_data = {str(r["type_usager"])[:40]: int(r["nb"]) for _, r in agg_usager.iterrows()}
        if pie_data:
            pie_path = _chart_pie(pie_data, "Contrôles par type d’usagers", tmp_dir, "pie_usagers.png")
            from PIL import Image as PILImage
            from reportlab.platypus import Image as RLImage
            _pimg = PILImage.open(pie_path)
            _target_w = avail_w * 0.75
            _target_h = _target_w * (_pimg.height / _pimg.width)
            _pimg.close()
            story.append(RLImage(pie_path, width=_target_w, height=_target_h))
        story.append(Spacer(1, 4 * mm))

        # Carte (si générée par le générateur cartographique)
        carte_usagers = get_cartes_dir() / "carte_global_usagers.png"
        if carte_usagers.exists():
            story.append(Paragraph("Carte : Contrôles par types d’usagers", styles["Heading2"]))
            img = RLImage(str(carte_usagers), width=avail_w, height=avail_w * 0.65)
            img.hAlign = "CENTER"
            story.append(img)
            story.append(Spacer(1, 3 * mm))

        # Tableau croisé Usagers × Domaine
        if cross_usager_dom is not None and not cross_usager_dom.empty:
            story.append(Paragraph("Tableau : Usagers × Domaine (contrôles)", styles["TableCaption"]))
            domain_cols = [c for c in cross_usager_dom.columns if c != "type_usager"]
            header = ["Type d’usagers"] + [str(c)[:22] for c in domain_cols]
            tbl_cross = [header]
            for _, row in cross_usager_dom.iterrows():
                tbl_cross.append([str(row["type_usager"])] + [str(int(row[c])) for c in domain_cols])
            # Largeur colonnes : 28% pour le libellé, le reste réparti
            other_w = (avail_w * 0.72) / max(1, len(domain_cols))
            col_widths = [avail_w * 0.28] + [other_w] * len(domain_cols)
            col_aligns = ["LEFT"] + ["RIGHT"] * len(domain_cols)
            story.append(_ofb_table(tbl_cross, col_widths=col_widths, col_aligns=col_aligns))

    story.append(PageBreak())

    # Annexes
    story.append(Paragraph('<a name="sec7"/>Annexes', styles["Heading1"]))
    story.append(Paragraph("Méthodologie", styles["Heading2"]))
    methodo = (
        f"<b>Période :</b> du {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y}. "
        f"<b>Périmètre :</b> département de la Côte-d'Or (21). "
        "<b>Sources :</b> OSCEAN (points de contrôle, PEJ, PA) et PVe OFB. "
        "Aucun filtre sur domaine ou thème ; tous NATINF pour PEJ et PVe. "
        "<b>Types d’usagers :</b> issus du champ OSCEAN <i>type_usager</i> des points de contrôle ; "
        "catégorie « dominante » par contrôle via le mapping ref/types_usagers.csv."
    )
    story.append(Paragraph(methodo, styles["BodyText"]))
    story.append(Paragraph("Glossaire : PA = Procédure administrative ; PEJ = Procédure d'enquête judiciaire ; PVe = Procès-verbal électronique ; NATINF = Nature d'infraction.", styles["BodySmall"]))

    doc.build(story)
    shutil.rmtree(tmp_dir, ignore_errors=True)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Génère le bilan global du service départemental (tous domaines/thèmes, PA, PEJ, PVe)."
    )
    parser.add_argument("--date-deb", type=str, default=None, help="Date de début (YYYY-MM-DD).")
    parser.add_argument("--date-fin", type=str, default=None, help="Date de fin (YYYY-MM-DD).")
    parser.add_argument("--dept-code", type=str, default=None, help="Code département (ex: 21).")
    return parser.parse_args()


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
        raise SystemExit("Dates invalides : utiliser YYYY-MM-DD.")
    DEPT_CODE = str(args.dept_code)

    root = _ROOT
    out_dir = get_out_dir("bilan_global")

    print(f"Période : {DATE_DEB.date():%d/%m/%Y} au {DATE_FIN.date():%d/%m/%Y} – Département {DEPT_CODE}.")

    print("Étape 1/4 : chargement des données...")
    with Spinner():
        point = load_point_ctrl(root, dept_code=DEPT_CODE, date_deb=DATE_DEB, date_fin=DATE_FIN)
        pa = load_pa(root, date_deb=DATE_DEB, date_fin=DATE_FIN)
        pej = load_pej(root, date_deb=DATE_DEB, date_fin=DATE_FIN)
        pve = load_pve(root, dept_code=DEPT_CODE, date_deb=DATE_DEB, date_fin=DATE_FIN)

    print("Étape 2/4 : analyse des contrôles...")
    with Spinner():
        analyse_controles_global(point, out_dir)

    print("Étape 3/4 : analyse PEJ / PA / PVe...")
    with Spinner():
        analyse_pej_pa_global(root, point, pa, pej, out_dir)
        analyse_pve_global(pve, out_dir)

    print("Étape 4/4 : génération du PDF...")
    with Spinner():
        generate_pdf_report(out_dir)

    print("Bilan global généré dans out/bilan_global.")


if __name__ == "__main__":
    main()
