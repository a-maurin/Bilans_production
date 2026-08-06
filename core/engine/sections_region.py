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
MODULE : GENERATEUR DES SECTIONS REGIONALES PDF (`sections_region.py`)
========================================================================================
Ce module est dédié à la génération des tableaux et graphiques d'analyse interdépartementale
pour les bilans à l'échelle régionale (DR) ou des Brigades de Mission Interdépartementale (BMI).

Rôles :
  1. Génération des histogrammes comparatifs entre départements (localisations, opérations).
  2. Construction des tableaux d'agrégation interdépartementale par domaine et thématique.
  3. Rendu des graphiques camembert de répartition des procédures (PA, PEJ, PVe) entre départements.
  4. Création des annexes régionales détaillées.
========================================================================================
"""
import logging
import re
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image as PILImage
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from core.common.pdf_utils import ofb_table
from core.engine.pdf_context import PdfContext
from core.common.utilitaires_metier import get_dept_name
from core.common.pdf_report_builder import PDFReportBuilder
from core.common.rendus_graphiques import chart_interdept_stacked_bar, chart_pie, chart_pie_legend_right

logger = logging.getLogger(__name__)

try:
    import geopandas as gpd
except ImportError:
    gpd = None


def _create_proportional_rl_image(img_path: Path, target_w: float, fallback_text: str, styles: dict, max_h: float | None = 450.0) -> RLImage | Paragraph:
    """Instancie une RLImage ReportLab en conservant mathématiquement le ratio largeur/hauteur réel."""
    body_style = styles.get("BodyText", styles.get("Normal"))
    if not img_path.exists():
        return Paragraph(f"<i>{fallback_text}</i>", body_style)
    try:
        with PILImage.open(img_path) as im:
            w_px, h_px = im.size
        aspect = (h_px / float(w_px)) if w_px > 0 else 0.7
        target_h = target_w * aspect
        if max_h is not None and target_h > max_h:
            target_h = max_h
            target_w = target_h / aspect
        return RLImage(str(img_path), width=target_w, height=target_h)
    except Exception:
        return Paragraph(f"<i>{fallback_text}</i>", body_style)


def _shapely_to_pathpatch(geom, transform):
    """Convertit un Polygon ou MultiPolygon Shapely en PathPatch Matplotlib pour le clipping."""
    from matplotlib.patches import PathPatch
    from matplotlib.path import Path
    import numpy as np

    codes = []
    verts = []

    def _add_poly(poly):
        ext = np.asarray(poly.exterior.coords)
        verts.extend(ext)
        codes.extend([Path.MOVETO] + [Path.LINETO] * (len(ext) - 2) + [Path.CLOSEPOLY])
        for interior in poly.interiors:
            int_coords = np.asarray(interior.coords)
            verts.extend(int_coords)
            codes.extend([Path.MOVETO] + [Path.LINETO] * (len(int_coords) - 2) + [Path.CLOSEPOLY])

    if geom.geom_type == 'Polygon':
        _add_poly(geom)
    elif geom.geom_type == 'MultiPolygon':
        for poly in geom.geoms:
            _add_poly(poly)

    path = Path(verts, codes)
    return PathPatch(path, transform=transform, facecolor='none', edgecolor='none')


def _generate_dept_vignette(
    dept_code: str,
    out_dir: Path,
    tmp_dir: Path,
    img_name: str,
    figure_scale: float = 1.0,
    profile_id: str = "global",
) -> Path:
    """
    Génère la vignette cartographique épurée du département avec conservation STRICTE du ratio d'aspect
    (1:1 carré, sans aucune déformation) et carte de chaleur de la pression de contrôle (Hexbin / Heatmap).
    """
    out_path = tmp_dir / img_name
    fig, ax = plt.subplots(figsize=(2.5 * figure_scale, 2.5 * figure_scale), dpi=150)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#ffffff')
    
    ax.set_aspect('equal')
    
    try:
        from core.cartographie.pochoir_helper import load_department_gdf
        from core.common.chargeurs_donnees import load_pnf_aoa_gdf, load_pnf_coeur_gdf
        from core.chemins_projet import PROJECT_ROOT
        
        is_pnf = str(profile_id).strip().lower() in ("pnf_v2", "pnf") or "pnf" in str(out_dir.name).lower() or str(dept_code).strip() in ("21", "52") and "pnf" in str(out_dir).lower()
        if is_pnf:
            gdf_aoa = load_pnf_aoa_gdf(PROJECT_ROOT)
            gdf_coeur = load_pnf_coeur_gdf(PROJECT_ROOT)
            gdf_dept = gdf_aoa if gdf_aoa is not None and not gdf_aoa.empty else load_department_gdf(dept_code, project_root=PROJECT_ROOT)
        else:
            gdf_coeur = None
            gdf_dept = load_department_gdf(dept_code, project_root=PROJECT_ROOT)
        
        if gdf_dept is not None and not gdf_dept.empty:
            if is_pnf and gdf_aoa is not None and not gdf_aoa.empty:
                target_dept = str(dept_code).strip().zfill(2)
                col_dep = next((c for c in gdf_aoa.columns if c.lower() in ("insee_dep", "num_depart", "code_dept", "insee_com")), None)
                if col_dep:
                    mask_dept = gdf_aoa[col_dep].astype(str).str.strip().str[:2] == target_dept
                    gdf_aoa[~mask_dept].plot(ax=ax, color='#f1f5f9', edgecolor='#003366', linewidth=0.7)
                    gdf_aoa[mask_dept].plot(ax=ax, color='#bfdbfe', edgecolor='#003366', linewidth=1.1)
                else:
                    gdf_aoa.plot(ax=ax, color='#f1f5f9', edgecolor='#003366', linewidth=0.8)
                
                if gdf_coeur is not None and not gdf_coeur.empty:
                    gdf_coeur.plot(ax=ax, facecolor='none', edgecolor='#16a34a', linewidth=1.8, label='Cœur de parc')
            else:
                gdf_dept.plot(ax=ax, color='#f1f5f9', edgecolor='#003366', linewidth=1.2, aspect='equal')
            
            carto_dir_default = PROJECT_ROOT / "data" / "sources" / "sig" / "CARTO"
            gpkg_files = list(out_dir.rglob("controles_*.gpkg")) + list(out_dir.glob("controles_*.gpkg"))
            gpkg_files += list(carto_dir_default.glob("controles_*.gpkg"))
            
            pts_dept = None
            if gpd is not None and gpkg_files:
                try:
                    import re
                    code_clean = str(dept_code).strip().lower()
                    out_dir_name = str(out_dir.name).strip().lower()
                    reg_match = re.search(r"r\d+", out_dir_name)
                    reg_tag = reg_match.group(0) if reg_match else ""

                    matching = []
                    for f in gpkg_files:
                        fname = f.name.lower()
                        if reg_tag and reg_tag in fname:
                            matching.append(f)
                        elif code_clean in fname:
                            matching.append(f)

                    if matching:
                        matching.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                        target_gpkg = matching[0]
                    else:
                        gpkg_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                        target_gpkg = gpkg_files[0]
                    
                    gdf_pts = gpd.read_file(target_gpkg)
                    if gdf_pts.crs is None or gdf_pts.crs.to_epsg() != 2154:
                        if gdf_pts.crs is not None:
                            gdf_pts = gdf_pts.to_crs("EPSG:2154")
                    
                    target_dept = str(dept_code).strip().split('.')[0].zfill(2)
                    if "num_depart" in gdf_pts.columns:
                        dept_clean = gdf_pts["num_depart"].astype(str).str.strip().str.split('.').str[0].str.zfill(2)
                        pts_dept = gdf_pts[dept_clean == target_dept]
                    
                    if pts_dept is None or pts_dept.empty:
                        try:
                            pts_dept = gpd.sjoin(gdf_pts, gdf_dept, predicate="within")
                        except Exception:
                            pts_dept = None
                except Exception:
                    pts_dept = None
            
            if pts_dept is not None and not pts_dept.empty:
                x_coords = pts_dept.geometry.x
                y_coords = pts_dept.geometry.y
                hb = ax.hexbin(
                    x_coords, y_coords,
                    gridsize=18,
                    cmap='YlOrRd',
                    mincnt=1,
                    alpha=0.85,
                    edgecolors='none'
                )

                try:
                    poly_geom = gdf_dept.geometry.union_all() if hasattr(gdf_dept.geometry, "union_all") else gdf_dept.geometry.unary_union
                    patch = _shapely_to_pathpatch(poly_geom, ax.transData)
                    ax.add_patch(patch)
                    hb.set_clip_path(patch)
                except Exception:
                    pass

                cbar = fig.colorbar(hb, ax=ax, orientation='horizontal', pad=0.02, shrink=0.25, aspect=12)
                cbar.set_label('Pression de contrôle (Faible -> Forte)', fontsize=6.5, color='#003366')
                cbar.ax.tick_params(labelsize=5.5)

            minx, miny, maxx, maxy = gdf_dept.total_bounds
            cx = (minx + maxx) / 2.0
            cy = (miny + maxy) / 2.0
            max_span = max(maxx - minx, maxy - miny) * 0.56
            ax.set_xlim(cx - max_span, cx + max_span)
            ax.set_ylim(cy - max_span, cy + max_span)
        else:
            ax.text(0.5, 0.5, f"Département {dept_code}", ha='center', va='center', fontsize=9, color='#003366')
    except Exception:
        ax.text(0.5, 0.5, f"Département {dept_code}", ha='center', va='center', fontsize=9, color='#003366')
        
    ax.axis('off')
    fig.tight_layout(pad=0.05)
    fig.savefig(out_path, dpi=150, facecolor='white')
    plt.close(fig)
    return out_path


def _generate_region_choropleth(
    dept_counts: dict[str, int],
    dept_code: str,
    tmp_dir: Path,
    img_name: str,
    figure_scale: float = 1.0,
    out_dir: Path | None = None,
    profile_id: str = "global",
) -> Path:
    """Génère la carte choroplèthe régionale (aplat par département selon le volume d'opérations)."""
    out_path = tmp_dir / img_name
    fig, ax = plt.subplots(figsize=(4.2 * figure_scale, 2.1 * figure_scale), dpi=150)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#ffffff')
    ax.set_aspect('equal')
    
    try:
        from core.cartographie.pochoir_helper import load_department_gdf
        from core.common.chargeurs_donnees import load_pnf_aoa_gdf, load_pnf_coeur_gdf
        from core.chemins_projet import PROJECT_ROOT
        
        is_pnf = (
            str(profile_id).strip().lower() in ("pnf_v2", "pnf")
            or (out_dir and "pnf" in str(out_dir.name).lower())
            or "pnf" in str(dept_code).lower()
        )
        if is_pnf:
            gdf_region = load_pnf_aoa_gdf(PROJECT_ROOT)
            gdf_coeur = load_pnf_coeur_gdf(PROJECT_ROOT)
        else:
            gdf_coeur = None
            gdf_region = load_department_gdf(dept_code, project_root=PROJECT_ROOT, dissolve=False)

        if gdf_region is not None and not gdf_region.empty:
            col_dep = next((c for c in gdf_region.columns if c.lower() in ("insee_dep", "num_depart", "code_dept", "insee_com")), gdf_region.columns[0])
            gdf_region["dep_code_clean"] = gdf_region[col_dep].astype(str).str.strip().str.upper().str[:2]
            gdf_region["nb_ops"] = gdf_region["dep_code_clean"].map(
                lambda d: dept_counts.get(d, dept_counts.get(d.zfill(2), dept_counts.get(str(int(d)) if d.isdigit() else d, 0)))
            ).fillna(0)
            
            gdf_region.plot(
                column="nb_ops",
                cmap="YlGnBu",
                linewidth=0.8,
                ax=ax,
                edgecolor="#003366",
                legend=True,
                legend_kwds={
                    "orientation": "horizontal",
                    "shrink": 0.25,
                    "aspect": 12,
                    "pad": 0.02,
                    "label": "Volume d'opérations de contrôle"
                }
            )

            if is_pnf and gdf_coeur is not None and not gdf_coeur.empty:
                gdf_coeur.plot(ax=ax, facecolor='none', edgecolor='#16a34a', linewidth=1.8, label='Cœur de parc')
            
            # Affinage de la typographie de la colorbar (6pt pour le titre, 5.5pt pour les ticks)
            if len(fig.axes) > 1:
                cbar_ax = fig.axes[-1]
                cbar_ax.tick_params(labelsize=5.5)
                cbar_ax.xaxis.label.set_size(6.0)
                cbar_ax.xaxis.label.set_color('#003366')

            for _, row in gdf_region.iterrows():
                centroid = row.geometry.centroid
                dep = row["dep_code_clean"]
                ops = int(row["nb_ops"])
                ax.annotate(
                    f"{dep}\n({ops})",
                    xy=(centroid.x, centroid.y),
                    ha="center",
                    va="center",
                    fontsize=6.0,
                    fontweight="bold",
                    color="#002b49",
                    bbox=dict(boxstyle="round,pad=0.12", fc="white", ec="none", alpha=0.65)
                )
            
            ax.set_axis_off()
            fig.tight_layout(pad=0.05)
            fig.savefig(str(out_path), dpi=150, bbox_inches="tight", facecolor="white")
            plt.close(fig)
            return out_path
    except Exception as e:
        plt.close(fig)
        logger.warning(f"Impossible de générer la carte choroplèthe régionale : {e}")
    return out_path


def render_sec_region_dashboard(ctx: PdfContext) -> None:
    """
    Rendu de la Partie 1 : Synthèse Régionale structurée de manière équilibrée sur 2 pages A4.
    
    PAGE 1 :
      - En-tête de section "1. Synthèse régionale"
      - Pavés KPI Cards Régionaux (Ops, Localisations, Procédures)
      - Carte choroplèthe régionale centrée (58 % de largeur)
      - Diagramme Donut par domaine centré (Légende 3 colonnes sous l'anneau, 0 % filtrés)
      - Saut de page strict (PageBreak)
      
    PAGE 2 :
      - Tableau comparatif interdépartemental (Ops, Locs, PEJ, PA, PVe)
      - Graphique en barres empilées de l'activité procédurale par département
    """
    title = ctx.section_title.get("sec_region_dashboard", "1. Synthèse régionale")
    ctx.builder.add_section("sec_region_dashboard", title)

    # 1. Pavés KPI Cards Régionaux
    kf: list[tuple[str, str]] = []
    if ctx.nb_ops:
        kf.append((str(ctx.nb_ops), "Opérations de contrôle"))
    if ctx.nb_localisations:
        kf.append((str(ctx.nb_localisations), "Localisations de contrôle"))
    if ctx.nb_pej or ctx.nb_pa or ctx.nb_pve:
        kf.append((f"{ctx.nb_pej or 0} / {ctx.nb_pa or 0} / {ctx.nb_pve or 0}", "Procédures PEJ / PA / PVe"))
    ctx.builder.add_key_figures(kf)
    ctx.builder.add_spacer(6)

    csv_path = ctx.out_dir / "region_detail_par_dept.csv"
    df = pd.DataFrame()
    if csv_path.exists():
        try:
            df = pd.read_csv(csv_path, sep=";", encoding="utf-8")
        except Exception:
            df = pd.DataFrame()

    # --- PAGE 1 VISUELS SUPERPOSÉS : CARTE CHOROPLÈTHE & DONUT PAR DOMAINE ---
    dept_counts: dict[str, int] = {}
    domain_counts: dict[str, int] = {}
    if not df.empty:
        df["departement"] = df["departement"].astype(str)
        if "nb_operations" in df.columns:
            dept_counts = df.groupby("departement")["nb_operations"].sum().to_dict()
            domain_counts = df.groupby("domaine")["nb_operations"].sum().to_dict()

    body_style = ctx.builder.styles.get("BodyText", ctx.builder.styles.get("Normal"))

    # Visuel 1 : Carte choroplèthe régionale centrée (54 % de largeur)
    prof_id = str((ctx.profile or {}).get("id", "")) if hasattr(ctx, "profile") else "global"
    carto_path = _generate_region_choropleth(dept_counts, ctx.dept_code, ctx.tmp_dir, "region_choropleth.png", figure_scale=ctx.figure_scale, out_dir=ctx.out_dir, profile_id=prof_id)
    if carto_path.exists():
        img_carto = _create_proportional_rl_image(carto_path, ctx.avail_w * 0.54, "Carte régionale indisponible", ctx.builder.styles)
    else:
        img_carto = Paragraph("<para align='center'><i>Carte régionale non disponible</i></para>", body_style)

    tbl_carto = Table([[img_carto]], colWidths=[ctx.avail_w])
    tbl_carto.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    ctx.builder.story.append(tbl_carto)

    # Note explicative dynamique bidirectionnelle sous la carte régionale en cas d'écart de comptage
    somme_dept = sum(dept_counts.values()) if dept_counts else 0
    total_reg = int(ctx.nb_ops) if ctx.nb_ops else 0
    if somme_dept > 0 and total_reg > 0 and somme_dept != total_reg:
        diff = abs(somme_dept - total_reg)
        str_somme = f"{somme_dept:,}".replace(",", "\u00a0")
        str_total = f"{total_reg:,}".replace(",", "\u00a0")
        if somme_dept > total_reg:
            note_txt = (
                f"<i>* La somme par département ({str_somme} ops) excède le total régional ({str_total} ops) "
                f"car {diff} opération(s) inter-départementale(s) sont comptabilisées dans chaque territoire concerné.</i>"
            )
        else:
            note_txt = (
                f"<i>* La somme par département ({str_somme} ops) est inférieure au total régional ({str_total} ops) "
                f"car {diff} opération(s) ne sont pas rattachées à un département de la région.</i>"
            )
        style_note = ParagraphStyle(
            "RegMapDynamicNote",
            parent=ctx.builder.styles.get("BodySmall", ctx.builder.styles["BodyText"]),
            fontSize=7,
            leading=8.5,
            textColor=colors.HexColor("#555555"),
            alignment=TA_CENTER,
        )
        ctx.builder.story.append(Spacer(1, 1))
        ctx.builder.story.append(Paragraph(note_txt, style_note))

    ctx.builder.add_spacer(4)

    # Visuel 2 : Diagramme Donut par Domaine (82 % de largeur, Donut agrandi avec légende à droite sur 1 colonne)
    domain_counts_filtered = {k: int(v) for k, v in domain_counts.items() if int(v) > 0}
    if domain_counts_filtered and sum(domain_counts_filtered.values()) > 0:
        donut_path = chart_pie_legend_right(
            domain_counts_filtered,
            "Répartition par Domaine de contrôle",
            ctx.tmp_dir,
            "region_domaines_donut.png",
            figure_scale=ctx.figure_scale,
            donut=True,
            legend_fontsize=8.5
        )
        img_donut = _create_proportional_rl_image(Path(donut_path), ctx.avail_w * 0.82, "Graphique domaines indisponible", ctx.builder.styles)
    else:
        img_donut = Paragraph("<para align='center'><i>Répartition par domaine non disponible</i></para>", body_style)

    tbl_donut = Table([[img_donut]], colWidths=[ctx.avail_w])
    tbl_donut.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    ctx.builder.story.append(tbl_donut)

    # --- SAUT DE PAGE VERS LA PAGE 2 DE LA SYNTHÈSE RÉGIONALE ---
    ctx.builder.story.append(PageBreak())

    # --- PAGE 2 : TABLEAU COMPARATIF INTERDÉPARTEMENTAL & GRAPHIQUE DES PROCÉDURES ---
    if not df.empty:
        df["departement"] = df["departement"].astype(str)
        depts = sorted(df["departement"].unique().tolist())

        # 1. Construction du tableau comparatif interdépartemental
        table_data = [["Dépt", "Libellé département", "Opérations", "Localisations", "PEJ", "PA", "PVe"]]
        tot_ops, tot_locs, tot_pej, tot_pa, tot_pve = 0, 0, 0, 0, 0

        for d in depts:
            sub = df[df["departement"] == d]
            ops = int(sub["nb_operations"].sum()) if "nb_operations" in sub.columns else 0
            locs = int(sub["nb_localisations"].sum()) if "nb_localisations" in sub.columns else 0
            pej = int(sub["nb_pej"].sum()) if "nb_pej" in sub.columns else 0
            pa = int(sub["nb_pa"].sum()) if "nb_pa" in sub.columns else 0
            pve = int(sub["nb_pve"].sum()) if "nb_pve" in sub.columns else 0

            tot_ops += ops
            tot_locs += locs
            tot_pej += pej
            tot_pa += pa
            tot_pve += pve

            table_data.append([
                d,
                get_dept_name(d),
                f"{ops:,}".replace(",", "\u202f"),
                f"{locs:,}".replace(",", "\u202f"),
                f"{pej:,}".replace(",", "\u202f"),
                f"{pa:,}".replace(",", "\u202f"),
                f"{pve:,}".replace(",", "\u202f")
            ])

        # Ligne de totalisation régionale
        table_data.append([
            "Total",
            "Périmètre régional",
            f"{tot_ops:,}".replace(",", "\u202f"),
            f"{tot_locs:,}".replace(",", "\u202f"),
            f"{tot_pej:,}".replace(",", "\u202f"),
            f"{tot_pa:,}".replace(",", "\u202f"),
            f"{tot_pve:,}".replace(",", "\u202f")
        ])

        col_widths = [
            0.08 * ctx.avail_w,
            0.34 * ctx.avail_w,
            0.12 * ctx.avail_w,
            0.14 * ctx.avail_w,
            0.10 * ctx.avail_w,
            0.10 * ctx.avail_w,
            0.12 * ctx.avail_w
        ]
        col_aligns = ["C", "L", "R", "R", "R", "R", "R"]

        tbl_interdept = ofb_table(table_data, col_widths=col_widths, col_aligns=col_aligns)
        ctx.builder.story.append(tbl_interdept)
        ctx.builder.add_spacer(12)

        # 2. Graphique comparatif des procédures en barres empilées
        try:
            depts_labels = [f"{d} - {get_dept_name(d)}" for d in depts]
            categories = ["PEJ", "PA", "PVe"]
            data_by_cat = {
                "PEJ": [int(df[df["departement"] == d]["nb_pej"].sum()) if "nb_pej" in df.columns else 0 for d in depts],
                "PA": [int(df[df["departement"] == d]["nb_pa"].sum()) if "nb_pa" in df.columns else 0 for d in depts],
                "PVe": [int(df[df["departement"] == d]["nb_pve"].sum()) if "nb_pve" in df.columns else 0 for d in depts]
            }
            
            chart_path = chart_interdept_stacked_bar(
                depts_labels,
                categories,
                data_by_cat,
                ctx.tmp_dir,
                "interdept_procedures.png",
                title="Comparaison de l'activité procédurale par département",
                figure_scale=ctx.figure_scale
            )
            img_chart = _create_proportional_rl_image(Path(chart_path), ctx.avail_w * ctx.chart_bar_w, "Graphique comparatif indisponible", ctx.builder.styles)
            ctx.builder.story.append(img_chart)
            ctx.builder.add_spacer(5)
        except Exception as e:
            ctx.builder.add_paragraph(f"<i>Impossible d'afficher le graphique comparatif : {e}</i>")
    else:
        ctx.builder.add_paragraph("<i>Aucune donnée régionale détaillée disponible pour la synthèse interdépartementale.</i>")


def render_sec_region_fiches(ctx: PdfContext) -> None:
    """Rendu de la Partie 2 : Fiches Départementales (1 Page A4 portrait stricte dans le rapport consolidé)."""
    csv_path = ctx.out_dir / "region_detail_par_dept.csv"
    if not csv_path.exists():
        ctx.builder.add_paragraph("Aucune donnée régionale détaillée disponible.")
        return

    df = pd.read_csv(csv_path, sep=";", encoding="utf-8")
    if df.empty:
        ctx.builder.add_paragraph("Aucune donnée régionale détaillée disponible.")
        return

    df["departement"] = df["departement"].astype(str)
    depts = sorted(df["departement"].unique().tolist())

    dept_dir = ctx.out_dir / "departements"
    if dept_dir.exists():
        for old_f in dept_dir.glob("Fiche_Dept_*.pdf"):
            try:
                old_f.unlink()
            except Exception:
                pass
    dept_dir.mkdir(parents=True, exist_ok=True)

    title_sec2 = ctx.section_title.get("sec_region_fiches", "2. Fiches départementales")

    for idx, d in enumerate(depts):
        dept_str = str(d)
        dept_name = get_dept_name(dept_str)
        df_dept = df[df["departement"] == dept_str]

        # Élimination du titre orphelin en bas de page 4 : le titre du chapitre 2 est posé en haut de la Page 5
        ctx.builder.add_page_break()
        if idx == 0:
            ctx.builder.add_section("sec_region_fiches", title_sec2, level=1)

        sec_id = f"sec_dept_{dept_str}"
        sec_label = f"Département {dept_str} - {dept_name}"
        ctx.builder.add_section(sec_id, sec_label, level=2, toc_level=1)

        total_ops = int(df_dept["nb_operations"].sum()) if not df_dept.empty else 0
        total_locs = int(df_dept["nb_localisations"].sum()) if not df_dept.empty else 0
        total_pej = int(df_dept["nb_pej"].sum()) if not df_dept.empty else 0
        total_pa = int(df_dept["nb_pa"].sum()) if not df_dept.empty else 0
        total_pve = int(df_dept["nb_pve"].sum()) if not df_dept.empty else 0

        dept_kfis = [
            (str(total_ops), "Opérations"),
            (str(total_locs), "Localisations"),
            (f"{total_pej} / {total_pa} / {total_pve}", "PEJ / PA / PVe")
        ]

        # 1. Traitement pour le Rapport Consolidé Régional (1 Page A4 portrait)
        if df_dept.empty or (total_ops == 0 and total_locs == 0 and total_pej == 0):
            ctx.builder.add_callout_box(
                f"Aucun contrôle ou donnée répertorié pour le département {dept_str} - {dept_name} sur le périmètre sélectionné.",
                title="Département sans activité ciblée"
            )
        else:
            ctx.builder.add_key_figures(dept_kfis)
            ctx.builder.add_spacer(4)

            cols_dom = [c for c in ["nb_operations", "nb_localisations", "nb_pej", "nb_pa", "nb_pve"] if c in df_dept.columns]
            df_dom = df_dept.groupby("domaine")[cols_dom].sum().reset_index()
            for c in ["nb_operations", "nb_localisations", "nb_pej", "nb_pa", "nb_pve"]:
                if c not in df_dom.columns:
                    df_dom[c] = 0

            pie_dict = {str(r["domaine"]): int(r["nb_localisations"]) for _, r in df_dom.iterrows() if r["nb_localisations"] > 0}
            
            prof_id = str((ctx.profile or {}).get("id", "")) if hasattr(ctx, "profile") else "global"
            vignette_path = _generate_dept_vignette(dept_str, ctx.out_dir, ctx.tmp_dir, f"vignette_dept_{dept_str}.png", figure_scale=ctx.figure_scale, profile_id=prof_id)
            img_vignette = _create_proportional_rl_image(vignette_path, ctx.avail_w * 0.46, "Vignette cartographique", ctx.builder.styles)

            if pie_dict:
                img_name = f"pie_dept_{dept_str}.png"
                pie_path = chart_pie_legend_right(
                    pie_dict,
                    f"Répartition par domaine ({dept_name})",
                    ctx.tmp_dir,
                    img_name,
                    figure_scale=ctx.figure_scale,
                    legend_fontsize=7.5
                )
                img_pie = _create_proportional_rl_image(Path(pie_path), ctx.avail_w * 0.48, "Graphique indisponible", ctx.builder.styles)
            else:
                body_s = ctx.builder.styles.get("BodyText", ctx.builder.styles.get("Normal"))
                img_pie = Paragraph("<i>Aucune donnée thématique</i>", body_s)

            double_visuel_tbl = Table(
                [[img_vignette, img_pie]],
                colWidths=[ctx.avail_w * 0.48, ctx.avail_w * 0.48]
            )
            double_visuel_tbl.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))
            ctx.builder.story.append(double_visuel_tbl)
            ctx.builder.add_spacer(4)

            # Micro-tableau condensé (Top 5)
            df_dom_sorted = df_dom.sort_values(by="nb_localisations", ascending=False)
            top5_df = df_dom_sorted.head(5)
            other_df = df_dom_sorted.iloc[5:]

            tbl_dept = [["Domaine Métier", "Opérations", "Localisations", "PEJ / PA / PVe"]]
            for _, r in top5_df.iterrows():
                dom_label = re.sub(r"^\[\d+\]\s*", "", str(r["domaine"])).strip()
                tbl_dept.append([
                    dom_label,
                    str(int(r["nb_operations"])),
                    str(int(r["nb_localisations"])),
                    f"{int(r.get('nb_pej', 0))} / {int(r.get('nb_pa', 0))} / {int(r.get('nb_pve', 0))}"
                ])
            if not other_df.empty:
                tbl_dept.append([
                    f"Autres domaines ({len(other_df)})",
                    str(int(other_df["nb_operations"].sum())),
                    str(int(other_df["nb_localisations"].sum())),
                    f"{int(other_df['nb_pej'].sum())} / {int(other_df['nb_pa'].sum())} / {int(other_df['nb_pve'].sum())}"
                ])

            ctx.builder.add_table(
                tbl_dept,
                caption=f"Synthèse par domaine principal - {dept_name}",
                col_widths=[ctx.avail_w * 0.40, ctx.avail_w * 0.20, ctx.avail_w * 0.20, ctx.avail_w * 0.20],
                col_aligns=["LEFT", "CENTER", "CENTER", "CENTER"],
                keep_together=True
            )

        # 2. Export du PDF autonome individuel 2-PAGES (departements/Fiche_Dept_<Code>.pdf)
        try:
            dept_pdf_path = dept_dir / f"Fiche_Dept_{dept_str}.pdf"
            b_dept = PDFReportBuilder(
                dept_pdf_path,
                f"Bilan Départemental {dept_str} - {dept_name}",
                title=f"Bilan Départemental {dept_str} - {dept_name}",
                content_only=True
            )
            b_dept.begin_content_pages()
            
            # --- PAGE 1 DU PDF AUTONOME ---
            b_dept.add_section("sec_dept_solo", f"Département {dept_str} - {dept_name}", level=1)
            if df_dept.empty or (total_ops == 0 and total_locs == 0 and total_pej == 0):
                b_dept.add_callout_box(
                    f"Aucun contrôle ou donnée répertorié pour le département {dept_str} - {dept_name} sur le périmètre sélectionné.",
                    title="Département sans activité ciblée"
                )
            else:
                b_dept.add_key_figures(dept_kfis)
                b_dept.add_spacer(4)
                b_dept.story.append(double_visuel_tbl)
                b_dept.add_spacer(4)
                b_dept.add_table(
                    tbl_dept,
                    caption=f"Synthèse par domaine (Top 5) - {dept_name}",
                    col_widths=[b_dept.avail_w * 0.40, b_dept.avail_w * 0.20, b_dept.avail_w * 0.20, b_dept.avail_w * 0.20],
                    col_aligns=["LEFT", "CENTER", "CENTER", "CENTER"],
                    keep_together=True
                )
                
                # --- PAGE 2 DU PDF AUTONOME : DÉTAIL MATRICIEL EXHAUSTIF ---
                b_dept.add_page_break()
                b_dept.add_section("sec_dept_solo_detail", f"Détail par domaine et thème - {dept_name}", level=2)
                
                headers_full = ["Domaine Métier", "Thème", "Opérations", "Localisations", "PEJ / PA / PVe"]
                tbl_full = [headers_full]
                
                for _, r_full in df_dept.iterrows():
                    tbl_full.append([
                        str(r_full.get("domaine", "")),
                        str(r_full.get("theme", "")),
                        str(int(r_full.get("nb_operations", 0))),
                        str(int(r_full.get("nb_localisations", 0))),
                        f"{int(r_full.get('nb_pej', 0))} / {int(r_full.get('nb_pa', 0))} / {int(r_full.get('nb_pve', 0))}"
                    ])
                
                b_dept.add_table(
                    tbl_full,
                    caption=f"Détail exhaustif des contrôles et procédures - {dept_name}",
                    col_widths=[b_dept.avail_w * 0.30, b_dept.avail_w * 0.30, b_dept.avail_w * 0.13, b_dept.avail_w * 0.13, b_dept.avail_w * 0.14],
                    col_aligns=["LEFT", "LEFT", "CENTER", "CENTER", "CENTER"],
                    keep_together=False
                )
                
            b_dept.build()
        except Exception as e_dept_pdf:
            pass


def render_sec_region_detail(ctx: PdfContext) -> None:
    """Rendu de l'Annexe Technique Détaillée (Détail matriciel par domaine et par thème)."""
    title = ctx.section_title.get("secregion", "Détail par département")
    ctx.builder.add_section("secregion", title)

    csv_path = ctx.out_dir / "region_detail_par_dept.csv"
    if not csv_path.exists():
        ctx.builder.add_paragraph("Aucune donnée régionale détaillée disponible.")
        return

    df = pd.read_csv(csv_path, sep=";", encoding="utf-8")
    if df.empty:
        ctx.builder.add_paragraph("Aucune donnée régionale détaillée disponible.")
        return

    df["departement"] = df["departement"].astype(str)
    for c in ["nb_operations", "nb_localisations", "nb_pej", "nb_pa", "nb_pve"]:
        if c not in df.columns:
            df[c] = 0

    depts = sorted(df["departement"].unique().tolist())

    for domaine, group_dom in df.groupby("domaine"):
        if domaine == "Hors domaine":
            continue

        ctx.builder.add_section(f"secregion_{domaine}", f"Domaine : {domaine}", level=3)
        agg_theme = group_dom.groupby("theme")[["nb_operations", "nb_localisations", "nb_pej", "nb_pa", "nb_pve"]].sum().reset_index()

        headers = ["Thème", "Département", "Opérations", "Localisations", "PEJ / PA / PVe"]
        col_widths = [ctx.avail_w * 0.25, ctx.avail_w * 0.25, ctx.avail_w * 0.15, ctx.avail_w * 0.15, ctx.avail_w * 0.20]
        col_aligns = ["LEFT", "LEFT", "CENTER", "CENTER", "CENTER"]
        tbl = [headers]

        for _, row_theme in agg_theme.iterrows():
            theme = str(row_theme["theme"])
            df_theme = group_dom[group_dom["theme"] == theme]
            first = True
            for d in depts:
                sub = df_theme[df_theme["departement"] == d]
                if sub.empty:
                    v_ops, v_locs, v_pej, v_pa, v_pve = 0, 0, 0, 0, 0
                else:
                    v_ops = sub["nb_operations"].sum()
                    v_locs = sub["nb_localisations"].sum()
                    v_pej = sub["nb_pej"].sum()
                    v_pa = sub["nb_pa"].sum()
                    v_pve = sub["nb_pve"].sum()

                dept_label = f"{d} - {get_dept_name(d)}"
                theme_label = theme if first else ""
                first = False

                tbl.append([
                    theme_label,
                    dept_label,
                    str(int(v_ops)),
                    str(int(v_locs)),
                    f"{int(v_pej)} / {int(v_pa)} / {int(v_pve)}"
                ])

            tot_suites = f"{int(row_theme['nb_pej'])} / {int(row_theme['nb_pa'])} / {int(row_theme['nb_pve'])}"
            tbl.append([
                "",
                "Total Région",
                str(int(row_theme["nb_operations"])),
                str(int(row_theme["nb_localisations"])),
                tot_suites
            ])

        ctx.builder.add_table(
            tbl,
            caption=f"Détail par département pour le domaine {domaine}",
            col_widths=col_widths,
            col_aligns=col_aligns,
            keep_together=True
        )
        ctx.builder.add_spacer(5)