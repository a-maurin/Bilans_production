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
MODULE : UTILITAIRES DE RENDU REPORTLAB DE L'OFB (`pdf_utils.py`)
========================================================================================
Ce module regroupe les composants bas niveau pour construire des éléments graphiques dans les PDF.

Fonctions principales :
  1. Troncation dynamique et habillage de textes pour les cellules de tableaux ReportLab.
  2. Composant personnalisé `VerticalText` (rotation à -90°) pour les en-têtes de tableaux étroits.
  3. Composant `OFBSplitTable` gérant la sécabilité propre des tableaux multi-pages sans coupures isolées.
  4. Générateur de tableaux d'indicateurs et chiffres clés (`key_figures_table`).
========================================================================================
"""
import re
import textwrap
from html import escape

from reportlab.lib import colors as rl_colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Flowable, Paragraph, Spacer, Table, TableStyle, PageBreak

from core.common.ofb_charte import (
    COLOR_PRIMARY,
    COLOR_TABLE_ALT_ROW,
    COLOR_TABLE_BORDER,
    COLOR_TABLE_HEADER_BG,
    FONT_FAMILY,
    MARGIN_LEFT,
    MARGIN_RIGHT,
    PAGE_W,
    _CELL_HEADER,
    _CELL_HEADER_RIGHT,
    _CELL_NORMAL,
    _CELL_RIGHT,
)

# Largeur utile par défaut pour le calcul des colonnes des tableaux larges
_AVail_W = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT


# ========================================================================================
# TRONCATION ET HABILLAGE DES TEXTES DANS LES CELLULES PDF
# ========================================================================================

def truncate_text_to_width(
    value: str,
    max_width_pt: float,
    *,
    font_name: str | None = None,
    font_size: float = 9.0,
    padding_pt: float = 8.0,
    suffix: str = "…",
) -> str:
    """Tronque un texte au pixel près avec des points de suspension si sa largeur excède la cellule."""
    txt = " ".join(str(value or "").split())
    if not txt:
        return ""
    fn = font_name or FONT_FAMILY
    usable = max(float(max_width_pt) - float(padding_pt), 1.0)
    try:
        if pdfmetrics.stringWidth(txt, fn, font_size) <= usable:
            return txt
    except Exception:
        return txt if len(txt) <= 48 else txt[:47].rstrip() + suffix

    end = len(txt)
    while end > 0:
        probe = txt[:end].rstrip()
        if not probe:
            break
        candidate = probe + suffix
        try:
            fits = pdfmetrics.stringWidth(candidate, fn, font_size) <= usable
        except Exception:
            fits = len(candidate) <= 48
        if fits:
            return candidate
        end -= 1
    return suffix


def wrap_plain_text_for_pdf_paragraph(
    value: object,
    *,
    wrap_width: int = 36,
    max_lines: int = 16,
) -> str:
    """Prépare du texte pour un composant ReportLab `Paragraph` avec découpe contrôlée par saut de ligne `<br/>`."""
    raw = " ".join(str(value or "").split())
    if not raw:
        return ""
    lines = textwrap.wrap(
        raw,
        width=max(8, int(wrap_width)),
        break_long_words=True,
        break_on_hyphens=True,
    )
    if len(lines) > int(max_lines):
        lines = lines[: int(max_lines)]
        if lines:
            lines[-1] = lines[-1].rstrip() + "…"
    return "<br/>".join(escape(line) for line in lines)


def _header_wrap_chars_for_col(col_width_pt: float, font_size: float) -> int:
    """Estime le nombre maximal de caractères rentrant sur une ligne d'en-tête selon sa largeur en points."""
    try:
        char_w = pdfmetrics.stringWidth("n", FONT_FAMILY, font_size)
    except Exception:
        char_w = float(font_size) * 0.55
    return max(8, int((float(col_width_pt) - 12.0) / max(char_w, 1.0)))


# ========================================================================================
# COMPOSANT FLOWABLE DE TEXTE VERTICAL (-90 DEGRÉS)
# ========================================================================================

class VerticalText(Flowable):
    """Élément personnalisé ReportLab dessinant du texte orienté verticalement à -90°."""

    def __init__(
        self,
        text: str,
        style=None,
        *,
        pad_x_pt: float = 0.0,
        max_lines: int = 6,
    ):
        super().__init__()
        self.text = str(text)
        self.style = style or _CELL_HEADER
        self.pad_x_pt = float(pad_x_pt)
        self.max_lines = max(1, int(max_lines))
        self._lines: list[str] = [self.text]

    def _split_lines(self, avail_width: float) -> list[str]:
        """Découpe le texte en plusieurs colonnes verticales si le libellé est très long."""
        txt = " ".join(self.text.split())
        if not txt:
            return [""]
        leading = max(float(getattr(self.style, "leading", 0) or 0), float(self.style.fontSize) + 1.5)
        width_cap = int(max(1, (avail_width - 6) // leading))
        max_lines = min(self.max_lines, max(width_cap, 2 if len(txt) > 22 else 1))
        if max_lines <= 1:
            return [txt]
        try:
            char_w = max(
                pdfmetrics.stringWidth("n", self.style.fontName, self.style.fontSize),
                self.style.fontSize * 0.45,
            )
        except Exception:
            char_w = self.style.fontSize * 0.5
        wrap_w = max(6, int((avail_width - 4) / max(char_w, 1.0)))
        wrapped = textwrap.wrap(
            txt,
            width=wrap_w,
            break_long_words=True,
            break_on_hyphens=True,
        )
        if not wrapped:
            return [txt]
        if len(wrapped) <= max_lines:
            return wrapped
        lines = wrapped[: max_lines - 1]
        tail = " ".join(wrapped[max_lines - 1 :])
        if len(tail) > wrap_w:
            tail = tail[: max(1, wrap_w - 1)].rstrip() + "…"
        lines.append(tail)
        return lines

    def wrap(self, availWidth: float, availHeight: float):
        """Calcule les dimensions nécessaires pour accueillir le texte après rotation."""
        self._lines = self._split_lines(availWidth)
        try:
            font = pdfmetrics.getFont(self.style.fontName)
            text_width = max(font.stringWidth(line, self.style.fontSize) for line in self._lines)
        except Exception:
            text_width = max(len(line) for line in self._lines) * self.style.fontSize * 0.6
        leading = max(float(getattr(self.style, "leading", 0) or 0), float(self.style.fontSize) + 1.5)
        self._block_w = len(self._lines) * leading + 6
        stack_extra = max(0, len(self._lines) - 1) * 6.0
        need_h = text_width + 30.0 + stack_extra
        self._height = max(float(self.style.fontSize) * 5.0, need_h)
        return (availWidth, self._height)

    def draw(self):
        """Effectue le dessin de la chaîne tournée à -90° sur le canvas du PDF."""
        canv = self.canv
        leading = max(float(getattr(self.style, "leading", 0) or 0), float(self.style.fontSize) + 1.5)
        fs = float(self.style.fontSize)
        fn = self.style.fontName
        col_step = leading + 1.0
        n = len(self._lines)
        total_stack = (n - 1) * col_step if n > 1 else 0.0

        canv.saveState()
        canv.setFont(fn, fs)
        canv.setFillColor(self.style.textColor)
        cx = self.width / 2.0 + self.pad_x_pt
        cy = self.height / 2.0
        canv.translate(cx, cy)
        canv.rotate(-90)
        start_y = total_stack / 2.0
        for i, line in enumerate(self._lines):
            y = start_y - i * col_step
            canv.drawCentredString(0, y, line)
        canv.restoreState()


# ========================================================================================
# TABLEAU SURCHARGÉ AVEC SÉCABILITÉ CONTRÔLÉE
# ========================================================================================

class OFBSplitTable(Table):
    """Surcharge de la classe `Table` ReportLab garantissant une césure élégante des tableaux sur 2 pages."""

    def split(self, availWidth, availHeight):
        result = super().split(availWidth, availHeight)
        if not result or len(result) != 2:
            return result

        t1, t2 = result
        lignes_avant = len(t1._rowHeights) if hasattr(t1, '_rowHeights') else len(t1._cellvalues)
        lignes_apres = len(t2._rowHeights) if hasattr(t2, '_rowHeights') else len(t2._cellvalues)

        if lignes_avant >= 5 and lignes_apres >= lignes_avant:
            return result

        if lignes_avant >= 5:
            repeat = getattr(self, 'repeatRows', 0)
            total_unique = lignes_avant + lignes_apres - repeat
            max_n = (total_unique + repeat) // 2

            if max_n >= 5:
                target_height = availHeight * 0.5
                if hasattr(t1, '_rowHeights') and len(t1._rowHeights) >= max_n:
                    target_height = sum(t1._rowHeights[:max_n]) + 15

                new_result = super().split(availWidth, target_height)
                if new_result and len(new_result) == 2:
                    nt1, nt2 = new_result
                    nl_avant = len(nt1._rowHeights) if hasattr(nt1, '_rowHeights') else len(nt1._cellvalues)
                    nl_apres = len(nt2._rowHeights) if hasattr(nt2, '_rowHeights') else len(nt2._cellvalues)
                    if nl_avant >= 5 and nl_apres >= nl_avant:
                        return [nt1, PageBreak(), nt2]

        if availHeight > 450:
            return result

        return []


# ========================================================================================
# CONSTRUCTEURS DE TABLEAUX OFB (STANDARDS ET LARGES)
# ========================================================================================

def ofb_table_wide(
    data_rows: list,
    col_widths=None,
    col_aligns=None,
    avail_w: float = None,
    *,
    split_by_row: bool = False,
    vertical_header_pad_x_pt: float = 0.0,
    vertical_header_max_lines: int = 6,
    vertical_header_font_size: float = 7.0,
    vertical_header_row_padding_pt: float = 8.0,
    header_layout: str = "vertical",
):
    """Construit un tableau aux couleurs OFB avec en-têtes verticales ou multi-lignes pour les nombreuses colonnes."""
    if not data_rows:
        return OFBSplitTable([], colWidths=[])

    vh_fs = max(6.0, float(vertical_header_font_size))
    vh_leading = max(vh_fs + 1.5, vh_fs * 1.25)
    vh_pad = max(4.0, float(vertical_header_row_padding_pt))
    vh_max_lines = max(1, int(vertical_header_max_lines))
    cell_header_vert = ParagraphStyle(
        "CellHeaderVert",
        parent=_CELL_HEADER,
        fontSize=vh_fs,
        leading=vh_leading,
    )
    cell_header_wrap = ParagraphStyle(
        "CellHeaderWrap",
        parent=_CELL_HEADER,
        fontSize=vh_fs,
        leading=vh_leading,
        alignment=TA_CENTER,
    )
    use_horizontal_wrap = str(header_layout or "vertical").strip().lower() in (
        "horizontal",
        "horizontal_wrap",
        "wrap",
    )

    avail_w = avail_w or _AVail_W
    n_cols = max(len(r) for r in data_rows)
    if n_cols == 0:
        return OFBSplitTable([], colWidths=[])

    def _looks_numeric(text: str) -> bool:
        txt = str(text).strip().replace("\u202f", "").replace(" ", "")
        if not txt:
            return False
        if txt.endswith("%"):
            txt = txt[:-1]
        return bool(re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?", txt))

    if col_widths is None:
        first_w = avail_w * 0.28
        rest_w = (avail_w - first_w) / max(1, n_cols - 1)
        col_widths = [first_w] + [rest_w] * (n_cols - 1)
    if col_aligns is None:
        right_cols = set()
        for row in data_rows[1:]:
            for ci, cell in enumerate(row):
                if _looks_numeric(str(cell)):
                    right_cols.add(ci)
        col_aligns = [
            "RIGHT" if ci in right_cols else "LEFT"
            for ci in range(n_cols)
        ]

    wrapped = []
    for ri, row in enumerate(data_rows):
        new_row = []
        for ci, cell in enumerate(row):
            cell_str = str(cell) if cell is not None else ""
            if ri == 0:
                col_w = col_widths[ci] if col_widths and ci < len(col_widths) else None
                if use_horizontal_wrap:
                    wrap_chars = _header_wrap_chars_for_col(
                        float(col_w) if col_w is not None else avail_w / max(1, n_cols),
                        vh_fs,
                    )
                    html = wrap_plain_text_for_pdf_paragraph(
                        cell_str,
                        wrap_width=wrap_chars,
                        max_lines=vh_max_lines,
                    )
                    hdr_style = _CELL_HEADER if ci == 0 else cell_header_wrap
                    new_row.append(Paragraph(html, hdr_style))
                elif ci == 0:
                    new_row.append(Paragraph(cell_str, _CELL_HEADER))
                else:
                    new_row.append(
                        VerticalText(
                            cell_str,
                            cell_header_vert,
                            pad_x_pt=vertical_header_pad_x_pt,
                            max_lines=vh_max_lines,
                        )
                    )
            else:
                is_right = col_aligns[ci] == "RIGHT" if ci < len(col_aligns) else False
                style = _CELL_RIGHT if is_right else _CELL_NORMAL
                new_row.append(Paragraph(cell_str, style))
        while len(new_row) < n_cols:
            if ri == 0:
                if use_horizontal_wrap:
                    new_row.append(Paragraph("", cell_header_wrap))
                else:
                    new_row.append(
                        VerticalText(
                            "",
                            cell_header_vert,
                            pad_x_pt=vertical_header_pad_x_pt,
                            max_lines=vh_max_lines,
                        )
                    )
            else:
                new_row.append(Paragraph("", _CELL_NORMAL))
        wrapped.append(new_row)

    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_TABLE_HEADER_BG),
        ("BOTTOMPADDING", (0, 0), (-1, 0), vh_pad),
        ("TOPPADDING", (0, 0), (-1, 0), vh_pad),
        ("LEFTPADDING", (0, 0), (-1, 0), 6),
        ("RIGHTPADDING", (0, 0), (-1, 0), 6),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 3),
        ("TOPPADDING", (0, 1), (-1, -1), 3),
        ("LEFTPADDING", (0, 1), (-1, -1), 4),
        ("RIGHTPADDING", (0, 1), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_TABLE_BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
    ]
    if use_horizontal_wrap and n_cols > 1:
        style_cmds.append(("ALIGN", (1, 0), (-1, 0), "CENTER"))
    for i in range(1, len(wrapped)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), COLOR_TABLE_ALT_ROW))

    tbl = OFBSplitTable(
        wrapped,
        colWidths=col_widths,
        repeatRows=1,
        splitByRow=1,
    )
    tbl.setStyle(TableStyle(style_cmds))
    return tbl


def ofb_table(
    data_rows,
    col_widths=None,
    col_aligns=None,
    *,
    header_font_size: float | None = None,
    split_by_row: bool = False,
    show_grid: bool = True,
    zebra_rows: bool = True,
    header_row: bool = True,
):
    """Crée un tableau standard stylisé aux couleurs OFB (en-tête bleu, lignes alternées gris/blanc)."""

    def _looks_numeric(text: str) -> bool:
        txt = text.strip().replace("\u202f", "").replace(" ", "")
        if not txt:
            return False
        if txt.endswith("%"):
            txt_num = txt[:-1]
        else:
            txt_num = txt
        return bool(re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?", txt_num))

    inferred_aligns = None
    if not col_aligns and data_rows:
        right_cols = set()
        data_start = 1 if header_row else 0
        for row in data_rows[data_start:]:
            for ci, cell in enumerate(row):
                if isinstance(cell, str) and _looks_numeric(cell):
                    right_cols.add(ci)
        if right_cols:
            max_cols = max(len(r) for r in data_rows)
            inferred_aligns = [
                "RIGHT" if ci in right_cols else "LEFT" for ci in range(max_cols)
            ]

    hdr_left = _CELL_HEADER
    hdr_right = _CELL_HEADER_RIGHT
    if header_font_size is not None:
        fs = float(header_font_size)
        lead = max(fs + 2.0, fs * 1.25)
        hdr_left = ParagraphStyle(
            "CellHeaderCustom",
            parent=_CELL_HEADER,
            fontSize=fs,
            leading=lead,
        )
        hdr_right = ParagraphStyle(
            "CellHeaderRightCustom",
            parent=_CELL_HEADER_RIGHT,
            fontSize=fs,
            leading=lead,
        )

    wrapped = []
    for ri, row in enumerate(data_rows):
        new_row = []
        for ci, cell in enumerate(row):
            if isinstance(cell, str):
                is_right = False
                if col_aligns and ci < len(col_aligns):
                    is_right = col_aligns[ci] == "RIGHT"
                elif inferred_aligns and ci < len(inferred_aligns):
                    is_right = inferred_aligns[ci] == "RIGHT"

                is_header = header_row and ri == 0
                if is_header:
                    style = hdr_right if is_right else hdr_left
                else:
                    style = _CELL_RIGHT if is_right else _CELL_NORMAL
                new_row.append(Paragraph(cell, style))
            else:
                new_row.append(cell)
        wrapped.append(new_row)

    style_cmds = [
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    if header_row:
        style_cmds.extend(
            [
                ("BACKGROUND", (0, 0), (-1, 0), COLOR_TABLE_HEADER_BG),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
                ("TOPPADDING", (0, 0), (-1, 0), 5),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 3),
                ("TOPPADDING", (0, 1), (-1, -1), 3),
            ]
        )
    else:
        style_cmds.extend(
            [
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    if show_grid:
        style_cmds.append(("GRID", (0, 0), (-1, -1), 0.5, COLOR_TABLE_BORDER))
    if zebra_rows:
        start = 1 if header_row else 0
        for i in range(start, len(wrapped)):
            if (i - start) % 2 == 1:
                style_cmds.append(("BACKGROUND", (0, i), (-1, i), COLOR_TABLE_ALT_ROW))

    tbl = OFBSplitTable(
        wrapped,
        colWidths=col_widths,
        repeatRows=1 if header_row else 0,
        splitByRow=1,
    )
    tbl.setStyle(TableStyle(style_cmds))
    return tbl


# ========================================================================================
# GENERATEURS DE TABLEAUX DE CHIFFRES CLES
# ========================================================================================

def key_figures_table(
    figures: list[tuple[str, str]],
    styles,
    *,
    density: str = "auto",
    table_width: float | None = None,
):
    """Construit un bloc de mise en valeur des chiffres clés (cadre bleu, grands nombres en gras)."""
    if not figures:
        return Spacer(1, 0)
    import math

    n = len(figures)
    max_per_line = 4

    if n <= max_per_line:
        rows = 1
        cols_per_row = n
    else:
        rows = math.ceil(n / max_per_line)
        cols_per_row = math.ceil(n / rows)

    if cols_per_row <= 3:
        scale = 1.0
        val_pad, lbl_pad = 5, 5
    else:
        scale = 0.85
        val_pad, lbl_pad = 4, 4

    val_style = ParagraphStyle(
        "KFDynVal",
        parent=styles["KeyFigure"],
        fontSize=styles["KeyFigure"].fontSize * scale,
        leading=styles["KeyFigure"].leading * scale,
    )
    lbl_style = ParagraphStyle(
        "KFDynLbl",
        parent=styles["KeyFigureLabel"],
        fontSize=styles["KeyFigureLabel"].fontSize * scale,
        leading=styles["KeyFigureLabel"].leading * scale,
    )

    if rows == 1:
        header = []
        labels = []
        for val, lbl in figures:
            v_text = str(val)
            if len(v_text) > 18:
                v_scale = max(0.55, min(1.0, 18.0 / len(v_text)))
                v_style = ParagraphStyle(
                    f"KFDynValLong_{len(v_text)}",
                    parent=val_style,
                    fontSize=val_style.fontSize * v_scale,
                    leading=val_style.leading * v_scale,
                )
            else:
                v_style = val_style
            header.append(Paragraph(f"<nobr><b>{v_text}</b></nobr>", v_style))
            labels.append(Paragraph(lbl, lbl_style))
        total_w = float(table_width) if table_width is not None else (PAGE_W - MARGIN_LEFT - MARGIN_RIGHT)
        col_w = total_w / n
        tbl = OFBSplitTable([header, labels], colWidths=[col_w] * n)
        tbl.setStyle(
            TableStyle(
                [
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("BOX", (0, 0), (-1, -1), 1, rl_colors.HexColor(COLOR_PRIMARY)),
                    ("LINEBELOW", (0, 0), (-1, 0), 0.5, COLOR_TABLE_BORDER),
                    ("TOPPADDING", (0, 0), (-1, 0), val_pad),
                    ("BOTTOMPADDING", (0, -1), (-1, -1), lbl_pad),
                ]
            )
        )
        return tbl

    split_rows = []
    start = 0
    for i in range(rows):
        take = math.ceil((n - start) / (rows - i))
        split_rows.append(figures[start : start + take])
        start += take

    return key_figures_table_rows(
        split_rows,
        styles,
        table_width=table_width,
        val_style=val_style,
        lbl_style=lbl_style,
    )


def key_figures_table_rows(
    figures_rows: list[list[tuple[str, str]]],
    styles,
    *,
    table_width: float | None = None,
    val_style=None,
    lbl_style=None,
):
    """Génère la grille multi-lignes des chiffres clés pour les bilans denses."""
    if not figures_rows or not any(figures_rows):
        return Spacer(1, 0)
    n_cols = max(len(row) for row in figures_rows)
    val_style = val_style or styles["KeyFigure"]
    lbl_style = lbl_style or styles["KeyFigureLabel"]
    table_rows: list[list] = []
    for row_figs in figures_rows:
        val_cells: list = []
        lbl_cells: list = []
        for i in range(n_cols):
            if i < len(row_figs):
                val, lbl = row_figs[i]
                val_cells.append(Paragraph(f"<b>{val}</b>", val_style))
                lbl_cells.append(Paragraph(lbl, lbl_style))
            else:
                val_cells.append(Paragraph("", val_style))
                lbl_cells.append(Paragraph("", lbl_style))
        table_rows.append(val_cells)
        table_rows.append(lbl_cells)
    total_w = float(table_width) if table_width is not None else (PAGE_W - MARGIN_LEFT - MARGIN_RIGHT)
    col_w = total_w / n_cols
    tbl = OFBSplitTable(table_rows, colWidths=[col_w] * n_cols)
    style_cmds: list = [
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 1, rl_colors.HexColor(COLOR_PRIMARY)),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
    ]
    if len(table_rows) >= 4:
        style_cmds.append(("LINEBELOW", (0, 1), (-1, 1), 0.5, COLOR_TABLE_BORDER))
    tbl.setStyle(TableStyle(style_cmds))
    return tbl