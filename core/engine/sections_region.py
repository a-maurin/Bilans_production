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

# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, mais SANS AUCUNE GARANTIE ;
# sans même la garantie implicite de QUALITÉ MARCHANDE ou D'ADÉQUATION À UN USAGE PARTICULIER.
# Voir la Licence Publique Générale GNU pour plus de détails.

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib import colors
from core.common.pdf_utils import ofb_table
from core.engine.pdf_context import PdfContext
from core.common.utilitaires_metier import get_dept_name
from core.common.pdf_report_builder import PDFReportBuilder
from core.common.rendus_graphiques import chart_interdept_stacked_bar, chart_pie

try:
    import geopandas as gpd
except ImportError:
    gpd = None


def _generate_dept_vignette(dept_code: str, out_dir: Path, tmp_dir: Path, img_name: str, figure_scale: float = 1.0) -> Path:
    """
    Génère la vignette cartographique épurée du département avec conservation STRICTE du ratio d'aspect
    (1:1, pas de déformation) et carte de chaleur de la pression de contrôle (Hexbin / Heatmap).
    """
    out_path = tmp_dir / img_name
    fig, ax = plt.subplots(figsize=(2.8 * figure_scale, 2.0 * figure_scale), dpi=150)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('#ffffff')
    
    # CRITIQUE : Conserver l'échelle géographique 1:1 pour ne jamais déformer/écraser le département
    ax.set_aspect('equal')
    
    try:
        from core.cartographie.pochoir_helper import load_department_gdf
        from core.chemins_projet import PROJECT_ROOT
        gdf_dept = load_department_gdf(dept_code, project_root=PROJECT_ROOT)
        
        if gdf_dept is not None and not gdf_dept.empty:
            gdf_dept.plot(ax=ax, color='#f1f5f9', edgecolor='#003366', linewidth=1.2, aspect='equal')
            
            # Recherche récursive de la couche géolocalisée des points de contrôles (GPKG)
            gpkg_files = list(out_dir.rglob("controles_*.gpkg"))
            pts_dept = None
            if gpd is not None and gpkg_files:
                try:
                    gdf_pts = gpd.read_file(gpkg_files[0])
                    if gdf_pts.crs is None or gdf_pts.crs.to_epsg() != 2154:
                        if gdf_pts.crs is not None:
                            gdf_pts = gdf_pts.to_crs("EPSG:2154")
                    
                    target_dept = str(dept_code).strip().split('.')[0].zfill(2)
                    if "num_depart" in gdf_pts.columns:
                        dept_clean = gdf_pts["num_depart"].astype(str).str.strip().str.split('.').str[0].str.zfill(2)
                        pts_dept = gdf_pts[dept_clean == target_dept]
                    else:
                        pts_dept = gpd.sjoin(gdf_pts, gdf_dept, predicate="within")
                except Exception:
                    pts_dept = None
            
            # Carte de chaleur / Hexbin de la densité de la pression de contrôle
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
                cbar = fig.colorbar(hb, ax=ax, orientation='horizontal', pad=0.02, shrink=0.7, aspect=18)
                cbar.set_label('Pression de contrôle (Faible ➔ Forte)', fontsize=6.5, color='#003366')
                cbar.ax.tick_params(labelsize=5.5)

            minx, miny, maxx, maxy = gdf_dept.total_bounds
            pad_x = (maxx - minx) * 0.05
            pad_y = (maxy - miny) * 0.05
            ax.set_xlim(minx - pad_x, maxx + pad_x)
            ax.set_ylim(miny - pad_y, maxy + pad_y)
        else:
            ax.text(0.5, 0.5, f"Département {dept_code}", ha='center', va='center', fontsize=9, color='#003366')
    except Exception:
        ax.text(0.5, 0.5, f"Département {dept_code}", ha='center', va='center', fontsize=9, color='#003366')
        
    ax.axis('off')
    plt.tight_layout(pad=0.1)
    fig.savefig(out_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return out_path


def render_sec_region_dashboard(ctx: PdfContext) -> None:
    """Rendu de la Partie 1 : Dashboard Régional Macro (Pavés KPI Cards + Visuels interdépartementaux)."""
    title = ctx.section_title.get("sec_region_dashboard", "1. Synthèse régionale")
    ctx.builder.add_section("sec_region_dashboard", title)

    # 1. Pavés KPI Cards Régionaux (Stylisés)
    kf: list[tuple[str, str]] = []
    if ctx.nb_ops:
        kf.append((str(ctx.nb_ops), "Opérations de contrôle"))
    if ctx.nb_localisations:
        kf.append((str(ctx.nb_localisations), "Localisations de contrôle"))
    if ctx.nb_pej or ctx.nb_pa or ctx.nb_pve:
        kf.append((f"{ctx.nb_pej or 0} / {ctx.nb_pa or 0} / {ctx.nb_pve or 0}", "Procédures PEJ / PA / PVe"))
    ctx.builder.add_key_figures(kf)
    ctx.builder.add_spacer(10)

    # 2. Graphique comparatif interdépartemental
    csv_path = ctx.out_dir / "region_detail_par_dept.csv"
    if csv_path.exists():
        try:
            df = pd.read_csv(csv_path, sep=";", encoding="utf-8")
            if not df.empty:
                df["departement"] = df["departement"].astype(str)
                depts = sorted(df["departement"].unique().tolist())
                depts_labels = [f"{d} - {get_dept_name(d)}" for d in depts]
                
                categories = ["PEJ", "PA", "PVe"]
                data_by_cat = {
                    "PEJ": [df[df["departement"] == d]["nb_pej"].sum() for d in depts],
                    "PA": [df[df["departement"] == d]["nb_pa"].sum() for d in depts],
                    "PVe": [df[df["departement"] == d]["nb_pve"].sum() for d in depts]
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
                ctx.builder.add_image(Path(chart_path), width_ratio=ctx.chart_bar_w)
                ctx.builder.add_spacer(5)
        except Exception as e:
            ctx.builder.add_paragraph(f"<i>Impossible d'afficher le graphique comparatif : {e}</i>")


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
            
            try:
                vignette_path = _generate_dept_vignette(dept_str, ctx.out_dir, ctx.tmp_dir, f"vignette_dept_{dept_str}.png", figure_scale=ctx.figure_scale)
                img_vignette = RLImage(str(vignette_path), width=ctx.avail_w * 0.46, height=ctx.avail_w * 0.32)
            except Exception:
                img_vignette = Paragraph("<i>Vignette cartographique</i>", ctx.builder.styles["Normal"])

            if pie_dict:
                try:
                    img_name = f"pie_dept_{dept_str}.png"
                    pie_path = chart_pie(
                        pie_dict,
                        f"Répartition par domaine ({dept_name})",
                        ctx.tmp_dir,
                        img_name,
                        figure_scale=ctx.figure_scale
                    )
                    img_pie = RLImage(str(pie_path), width=ctx.avail_w * 0.46, height=ctx.avail_w * 0.32)
                except Exception:
                    img_pie = Paragraph("<i>Graphique indisponible</i>", ctx.builder.styles["Normal"])
            else:
                img_pie = Paragraph("<i>Aucune donnée thématique</i>", ctx.builder.styles["Normal"])

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
                tbl_dept.append([
                    str(r["domaine"]),
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
                caption=f"Synthèse par domaine (Top 5) - {dept_name}",
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
                
                # Tableau matriciel complet sans restriction au Top 5
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