"""Charte graphique OFB : couleurs, polices, styles PDF, Spinner."""
import itertools
import sys
import threading
import time
from pathlib import Path

from reportlab.lib import colors as rl_colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------------------------------------------------------------------------
# Charte graphique OFB
# ---------------------------------------------------------------------------
COLOR_PRIMARY = "#003A76"
COLOR_SECONDARY = "#1E4E85"
COLOR_GREY = "#333333"
COLOR_TABLE_HEADER_BG = rl_colors.HexColor("#003A76")
COLOR_TABLE_HEADER_FG = rl_colors.white
COLOR_TABLE_ALT_ROW = rl_colors.HexColor("#F0F4F8")
COLOR_TABLE_BORDER = rl_colors.HexColor("#CCCCCC")

# Couleurs pour les graphiques (camemberts, barres groupées) — aucune couleur en dur dans les scripts de bilan
COLOR_CHART_1 = COLOR_PRIMARY  # bleu OFB
COLOR_CHART_2 = "#53AB60"   # vert
COLOR_CHART_3 = "#F4A261"   # orangé
COLOR_CHART_4 = "#E76F51"   # rouge-orange
COLOR_CHART_5 = "#90BF83"   # vert clair
COLOR_CHART_6 = "#4296CE"   # bleu clair
CHART_PIE_COLORS = [COLOR_CHART_1, COLOR_CHART_2, COLOR_CHART_3, COLOR_CHART_4, COLOR_CHART_5, COLOR_CHART_6]
CHART_BAR_GROUPED_COLORS = [COLOR_CHART_1, COLOR_CHART_2, COLOR_CHART_3, COLOR_CHART_4]

PAGE_W, PAGE_H = A4
MARGIN_LEFT = 17 * mm
MARGIN_RIGHT = 17 * mm
MARGIN_TOP = 22 * mm
MARGIN_BOTTOM = 28 * mm


def _ref_img(name: str) -> Path:
    """Chemin vers une image dans ref/modele_ofb/word/media/."""
    ref_dir = Path(__file__).resolve().parent.parent.parent / "ref"
    return ref_dir / "modele_ofb" / "word" / "media" / name


IMG_LOGO_BANNER = _ref_img("image5.jpg")
IMG_BACKGROUND = _ref_img("image4.png")
IMG_FOOTER_DECO = _ref_img("image3.jpeg")


# ---------------------------------------------------------------------------
# Enregistrement polices
# ---------------------------------------------------------------------------
def _register_fonts() -> str:
    """Enregistre les polices dans reportlab. Retourne le nom de famille."""
    # Essayer d'abord la police Marianne (si disponible en TTF)
    marianne_dirs = [
        Path("/usr/share/fonts/truetype/marianne"),
        Path("/usr/local/share/fonts/marianne"),
        Path("~/.local/share/fonts/marianne").expanduser(),
    ]
    
    for marianne_dir in marianne_dirs:
        if marianne_dir.exists():
            # Look for TTF files only (ReportLab doesn't support OTF with PostScript outlines)
            regular_fonts = list(marianne_dir.glob("*Regular*.ttf"))
            bold_fonts = list(marianne_dir.glob("*Bold*.ttf"))
            italic_fonts = list(marianne_dir.glob("*Italic*.ttf"))
            bolditalic_fonts = list(marianne_dir.glob("*Bold*Italic*.ttf"))
            
            # Find the best matches
            regular = None
            bold = None
            italic = None
            bolditalic = None
            
            # Look for Regular font
            for font in regular_fonts:
                if "Regular" in font.name and not "Italic" in font.name:
                    regular = font
                    break
            
            # Look for Bold font (not italic)
            for font in bold_fonts:
                if "Bold" in font.name and "Italic" not in font.name and "ExtraBold" not in font.name:
                    bold = font
                    break
            
            # Look for Italic font (not bold)
            for font in italic_fonts:
                if "Italic" in font.name and "Bold" not in font.name and "Regular" in font.name:
                    italic = font
                    break
            
            # Look for BoldItalic font
            for font in bolditalic_fonts:
                if "Bold" in font.name and "Italic" in font.name and "ExtraBold" not in font.name:
                    bolditalic = font
                    break
            
            # If we found all required fonts, register them
            if regular and bold and italic and bolditalic:
                try:
                    pdfmetrics.registerFont(TTFont("Marianne", str(regular)))
                    pdfmetrics.registerFont(TTFont("Marianne-Bold", str(bold)))
                    pdfmetrics.registerFont(TTFont("Marianne-Italic", str(italic)))
                    pdfmetrics.registerFont(TTFont("Marianne-BoldItalic", str(bolditalic)))
                    pdfmetrics.registerFontFamily(
                        "Marianne",
                        normal="Marianne",
                        bold="Marianne-Bold",
                        italic="Marianne-Italic",
                        boldItalic="Marianne-BoldItalic",
                    )
                    return "Marianne"
                except Exception as e:
                    print(f"Erreur lors de l'enregistrement de la police Marianne: {e}")
                    # Continue to fallback fonts
    
    
    # Essayer Arial sur Windows
    if sys.platform == "win32":
        fonts_dir = Path(r"C:\Windows\Fonts")
        arial = fonts_dir / "arial.ttf"
        arial_bd = fonts_dir / "arialbd.ttf"
        arial_it = fonts_dir / "ariali.ttf"
        arial_bi = fonts_dir / "arialbi.ttf"
        if arial.exists():
            pdfmetrics.registerFont(TTFont("Arial", str(arial)))
            pdfmetrics.registerFont(TTFont("Arial-Bold", str(arial_bd)))
            pdfmetrics.registerFont(TTFont("Arial-Italic", str(arial_it)))
            pdfmetrics.registerFont(TTFont("Arial-BoldItalic", str(arial_bi)))
            pdfmetrics.registerFontFamily(
                "Arial",
                normal="Arial",
                bold="Arial-Bold",
                italic="Arial-Italic",
                boldItalic="Arial-BoldItalic",
            )
            return "Arial"
    
    # Utiliser Liberation Sans comme alternative (disponible sur la plupart des systèmes Linux)
    try:
        liberation_regular = Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf")
        liberation_bold = Path("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf")
        liberation_italic = Path("/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf")
        liberation_bolditalic = Path("/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf")
        
        if liberation_regular.exists() and liberation_bold.exists() and liberation_italic.exists() and liberation_bolditalic.exists():
            pdfmetrics.registerFont(TTFont("LiberationSans", str(liberation_regular)))
            pdfmetrics.registerFont(TTFont("LiberationSans-Bold", str(liberation_bold)))
            pdfmetrics.registerFont(TTFont("LiberationSans-Italic", str(liberation_italic)))
            pdfmetrics.registerFont(TTFont("LiberationSans-BoldItalic", str(liberation_bolditalic)))
            pdfmetrics.registerFontFamily(
                "LiberationSans",
                normal="LiberationSans",
                bold="LiberationSans-Bold",
                italic="LiberationSans-Italic",
                boldItalic="LiberationSans-BoldItalic",
            )
            return "LiberationSans"
    except Exception:
        pass
    
    # Fallback final : Helvetica (police par défaut de ReportLab)
    return "Helvetica"


FONT_FAMILY = _register_fonts()

_CELL_NORMAL = ParagraphStyle(
    "CellNormal",
    fontName=FONT_FAMILY,
    fontSize=9,
    leading=12,
    textColor=rl_colors.black,
    alignment=TA_LEFT,
)
_CELL_RIGHT = ParagraphStyle(
    "CellRight",
    fontName=FONT_FAMILY,
    fontSize=9,
    leading=12,
    textColor=rl_colors.black,
    alignment=TA_RIGHT,
)
_CELL_HEADER = ParagraphStyle(
    "CellHeader",
    fontName=f"{FONT_FAMILY}-Bold",
    fontSize=10,
    leading=13,
    textColor=rl_colors.white,
    alignment=TA_LEFT,
)
_CELL_HEADER_RIGHT = ParagraphStyle(
    "CellHeaderRight",
    fontName=f"{FONT_FAMILY}-Bold",
    fontSize=10,
    leading=13,
    textColor=rl_colors.white,
    alignment=TA_RIGHT,
)


def _get_styles():
    """Construit les ParagraphStyles conformes à la charte OFB."""
    ss = getSampleStyleSheet()
    styles = {
        "Title": ParagraphStyle(
            "OFBTitle",
            parent=ss["Title"],
            fontName=f"{FONT_FAMILY}-Bold",
            fontSize=26,
            leading=36,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_CENTER,
            spaceAfter=6 * mm,
        ),
        "Heading1": ParagraphStyle(
            "OFBH1",
            parent=ss["Heading1"],
            fontName=f"{FONT_FAMILY}-Bold",
            fontSize=18,
            leading=24,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_LEFT,
            spaceBefore=10 * mm,
            spaceAfter=4 * mm,
            keepWithNext=1,
        ),
        "Heading2": ParagraphStyle(
            "OFBH2",
            parent=ss["Heading2"],
            fontName=f"{FONT_FAMILY}-Bold",
            fontSize=14,
            leading=18,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_LEFT,
            spaceBefore=6 * mm,
            spaceAfter=3 * mm,
            keepWithNext=1,
        ),
        "Heading3": ParagraphStyle(
            "OFBH3",
            parent=ss["Heading3"],
            fontName=f"{FONT_FAMILY}-Italic",
            fontSize=12,
            leading=15,
            textColor=rl_colors.HexColor(COLOR_GREY),
            alignment=TA_LEFT,
            spaceBefore=4 * mm,
            spaceAfter=2 * mm,
            keepWithNext=1,
        ),
        "BodyText": ParagraphStyle(
            "OFBBody",
            parent=ss["BodyText"],
            fontName=FONT_FAMILY,
            fontSize=10,
            leading=15,
            textColor=rl_colors.black,
            alignment=TA_JUSTIFY,
            spaceBefore=1 * mm,
            spaceAfter=2 * mm,
        ),
        "BodySmall": ParagraphStyle(
            "OFBBodySmall",
            parent=ss["BodyText"],
            fontName=FONT_FAMILY,
            fontSize=8,
            leading=11,
            textColor=rl_colors.HexColor("#666666"),
            alignment=TA_LEFT,
            spaceBefore=1 * mm,
            spaceAfter=1 * mm,
        ),
        "TableCaption": ParagraphStyle(
            "OFBTableCaption",
            parent=ss["BodyText"],
            fontName=f"{FONT_FAMILY}-Bold",
            fontSize=10,
            leading=13,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_LEFT,
            spaceBefore=4 * mm,
            spaceAfter=2 * mm,
            keepWithNext=1,
        ),
        "KeyFigure": ParagraphStyle(
            "OFBKeyFigure",
            parent=ss["BodyText"],
            fontName=f"{FONT_FAMILY}-Bold",
            fontSize=22,
            leading=28,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_CENTER,
        ),
        "KeyFigureLabel": ParagraphStyle(
            "OFBKeyFigureLabel",
            parent=ss["BodyText"],
            fontName=FONT_FAMILY,
            fontSize=9,
            leading=12,
            textColor=rl_colors.HexColor(COLOR_GREY),
            alignment=TA_CENTER,
        ),
        "TOCEntry": ParagraphStyle(
            "OFBTOCEntry",
            parent=ss["BodyText"],
            fontName=FONT_FAMILY,
            fontSize=12,
            leading=20,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_LEFT,
            leftIndent=5 * mm,
        ),
    }
    return styles


class Spinner:
    """Spinner texte pour indiquer qu'un traitement est en cours."""

    def __init__(self) -> None:
        self._stop_event = threading.Event()
        self._thread = threading.Thread(target=self._spin, daemon=True)

    def _spin(self) -> None:
        frames = itertools.cycle("|/-\\")
        while not self._stop_event.is_set():
            frame = next(frames)
            sys.stdout.write(f"\r{frame}")
            sys.stdout.flush()
            time.sleep(0.1)

    def __enter__(self) -> "Spinner":
        sys.stdout.write("\n")
        sys.stdout.flush()
        self._thread.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self._stop_event.set()
        self._thread.join()
        sys.stdout.write("\r \r")
        sys.stdout.flush()
