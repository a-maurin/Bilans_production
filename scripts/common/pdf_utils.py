"""Helpers reportlab : tableaux OFB, chiffres clés."""
from reportlab.lib import colors as rl_colors
from reportlab.platypus import Paragraph, Table, TableStyle

from scripts.common.ofb_charte import (
    COLOR_PRIMARY,
    COLOR_TABLE_ALT_ROW,
    COLOR_TABLE_BORDER,
    COLOR_TABLE_HEADER_BG,
    MARGIN_LEFT,
    MARGIN_RIGHT,
    PAGE_W,
    _CELL_HEADER,
    _CELL_HEADER_RIGHT,
    _CELL_NORMAL,
    _CELL_RIGHT,
)


def ofb_table(data_rows, col_widths=None, col_aligns=None):
    """Crée un Table reportlab stylisé charte OFB (en-tête bleu, lignes alternées)."""
    wrapped = []
    for ri, row in enumerate(data_rows):
        new_row = []
        for ci, cell in enumerate(row):
            if isinstance(cell, str):
                is_right = (
                    col_aligns and ci < len(col_aligns) and col_aligns[ci] == "RIGHT"
                )
                if ri == 0:
                    style = _CELL_HEADER_RIGHT if is_right else _CELL_HEADER
                else:
                    style = _CELL_RIGHT if is_right else _CELL_NORMAL
                new_row.append(Paragraph(cell, style))
            else:
                new_row.append(cell)
        wrapped.append(new_row)

    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_TABLE_HEADER_BG),
        # Lignes d'en-tête : padding légèrement réduit pour éviter des hauteurs
        # excessives lorsque les libellés sont courts.
        ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
        ("TOPPADDING", (0, 0), (-1, 0), 5),
        # Lignes de données : padding plus serré pour compacter les tableaux.
        ("BOTTOMPADDING", (0, 1), (-1, -1), 3),
        ("TOPPADDING", (0, 1), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_TABLE_BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    for i in range(1, len(wrapped)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), COLOR_TABLE_ALT_ROW))

    tbl = Table(wrapped, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle(style_cmds))
    return tbl


def key_figures_table(figures: list[tuple[str, str]], styles):
    """Bloc de chiffres clés : liste de (valeur, libellé) affichés en ligne."""
    if not figures:
        return Spacer(1, 0)
    header = []
    labels = []
    for val, lbl in figures:
        header.append(Paragraph(f"<b>{val}</b>", styles["KeyFigure"]))
        labels.append(Paragraph(lbl, styles["KeyFigureLabel"]))
    col_w = (PAGE_W - MARGIN_LEFT - MARGIN_RIGHT) / len(figures)
    tbl = Table([header, labels], colWidths=[col_w] * len(figures))
    tbl.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOX", (0, 0), (-1, -1), 1, rl_colors.HexColor(COLOR_PRIMARY)),
                ("LINEBELOW", (0, 0), (-1, 0), 0.5, COLOR_TABLE_BORDER),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
            ]
        )
    )
    return tbl
