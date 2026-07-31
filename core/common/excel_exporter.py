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
# Conformément à la section 7(b) DE LA GPL v3, vous devez expressément conserver
# intactes et lisibles toutes les mentions d'auteur, notices de copyright et la présente
# clause dans chaque fichier source ou interface utilisateur redistribué. Toute version modifiée
# doit clairement indiquer qu'elle a été altérée et ne doit en aucun cas supprimer le nom
# de l'auteur original (Aguirre MAURIN).

"""
========================================================================================
MODULE : EXPORTATEUR DE DONNÉES EN CLASSEURS EXCEL (`excel_exporter.py`)
========================================================================================
Ce module s'occupe de la création et de la mise en forme automatique des fichiers Excel (.xlsx).

Fonctionnalités principales :
  1. Application de la charte graphique OFB (en-têtes bleus `#003366`, polices Arial, bordures grises).
  2. Ajustement automatique de la largeur des colonnes selon la longueur du contenu.
  3. Génération de classeurs multi-onglets (Synthèse Régionale + 1 onglet par Département).
========================================================================================
"""

import logging
from pathlib import Path
from typing import List, Optional
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)


# ========================================================================================
# MISE EN FORME DES FEUILLES DE CALCUL (STYLES OFB ET AUTO-FIT)
# ========================================================================================

def _format_worksheet(ws: Any) -> None:
    """Applique la charte graphique OFB et ajuste automatiquement la largeur des colonnes."""
    header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=9)
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )

    if hasattr(ws.views, "sheetView") and ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # 1. En-têtes (fond bleu roi, texte blanc gras, centré)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 2. Données et bordures (nombres à droite, textes à gauche)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # 3. Calcul dynamique de la largeur optimale des colonnes
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)


# ========================================================================================
# EXPORT DE LA SYNTHESE REGIONALE MULTI-ONGLETS
# ========================================================================================

def export_synthese_region_excel(
    out_dir: Path,
    df_detail: pd.DataFrame,
    depts: List[str],
    df_ratio: Optional[pd.DataFrame] = None,
    filename: str = "Synthese_Region.xlsx"
) -> Optional[Path]:
    """Exporte la synthèse régionale et les détails départementaux dans un classeur Excel.

    - Onglet 1 : Synthèse Régionale & Ratios globaux
    - Onglets 2+ : Résumé propre à chaque Département
    """
    if df_detail.empty:
        logger.warning("df_detail est vide, annulation de l'export Excel régional.")
        return None

    xlsx_path = out_dir / filename
    try:
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            # 1. Génération de l'onglet de Synthèse Régionale
            if df_ratio is not None and not df_ratio.empty:
                df_ratio.to_excel(writer, sheet_name="Synthese_Regionale", index=False)
            else:
                cols = [c for c in ["nb_operations", "nb_localisations", "nb_pej", "nb_pa", "nb_pve"] if c in df_detail.columns]
                agg_region = df_detail.groupby(["domaine", "theme"])[cols].sum().reset_index()
                agg_region.to_excel(writer, sheet_name="Synthese_Regionale", index=False)

            _format_worksheet(writer.sheets["Synthese_Regionale"])

            # 2. Génération des onglets individuels par département
            df_detail["departement"] = df_detail["departement"].astype(str)
            for dept in depts:
                dept_str = str(dept)
                df_dept = df_detail[df_detail["departement"] == dept_str]
                sheet_title = f"Dept_{dept_str}"[:31]
                if not df_dept.empty:
                    df_dept.to_excel(writer, sheet_name=sheet_title, index=False)
                else:
                    pd.DataFrame({"Message": ["Aucune donnée disponible pour ce département"]}).to_excel(
                        writer, sheet_name=sheet_title, index=False
                    )
                _format_worksheet(writer.sheets[sheet_title])

        logger.info(f"Classeur Excel régional formaté généré avec succès : {xlsx_path}")
        return xlsx_path

    except (OSError, ValueError, KeyError) as e:
        logger.error(f"Erreur lors de la génération du classeur Excel {xlsx_path} : {e}")
        return None

