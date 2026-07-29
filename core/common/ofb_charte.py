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
MODULE : CHARTE GRAPHIQUE OFFICIELLE DE L'OFB (`ofb_charte.py`)
========================================================================================
Ce module définit la charte visuelle officielle appliquée à tous les rapports PDF générés.

Contenu :
  1. Palette de couleurs institutionnelles (Bleu OFB, Vert, Orangé, Rouge d'alerte).
  2. Enregistrement des polices de caractères officielles (Marianne, Arial, LiberationSans).
  3. Définition des marges, espaces et styles de textes ReportLab (Titres, Corps, Tableaux).
  4. Classe d'animation visuelle 'Spinner' pour la console lors du traitement batch.
========================================================================================
"""
import os
import sys
import threading
import time
from pathlib import Path

# --- ACTIVATION DES CODES COULEURS ET ANIMATIONS ANSI (VT100) SUR CONSOLE WINDOWS ---
if sys.platform == "win32":
    import ctypes
    _STD_OUTPUT_HANDLE = -11
    _ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004

    def _enable_windows_vt100() -> bool:
        """Active le mode VT100/ANSI sur l'invite de commande Windows pour gérer le spinner."""
        try:
            os.system("")  # Déclenche l'activation ANSI de la console Windows
            handle = ctypes.windll.kernel32.GetStdHandle(_STD_OUTPUT_HANDLE)
            if handle is None or handle == -1:
                return False
            mode = ctypes.c_ulong()
            if ctypes.windll.kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                mode.value |= _ENABLE_VIRTUAL_TERMINAL_PROCESSING
                return ctypes.windll.kernel32.SetConsoleMode(handle, mode) != 0
            return ctypes.windll.kernel32.SetConsoleMode(handle, 7) != 0
        except Exception:
            return False
else:

    def _enable_windows_vt100() -> bool:
        return True  # Sur Linux/macOS, le terminal supporte nativement ANSI


# --- IMPORTS DE LA BIBLIOTHÈQUE PDF REPORTLAB ---
from reportlab.lib import colors as rl_colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ========================================================================================
# PALETTE DE COULEURS OFFICIELLES DE L'OFB
# ========================================================================================

COLOR_PRIMARY = "#003A76"  # Bleu roi officiel OFB (utilisé pour les titres et en-têtes)
COLOR_SECONDARY = "#1E4E85"  # Bleu secondaire pour les sous-titres
COLOR_GREY = "#333333"  # Gris foncé pour les textes généraux et valeurs neutres

# Couleurs pour les tableaux ReportLab
COLOR_CHART_AUTRE_RESULTAT = COLOR_GREY
COLOR_TABLE_HEADER_BG = rl_colors.HexColor("#003A76")  # Fond bleu des en-têtes de tableaux
COLOR_TABLE_HEADER_FG = rl_colors.white  # Texte blanc des en-têtes de tableaux
COLOR_TABLE_ALT_ROW = rl_colors.HexColor("#F0F4F8")  # Couleur de fond alternée des lignes de tableau
COLOR_TABLE_BORDER = rl_colors.HexColor("#CCCCCC")  # Bordures grises claires des tableaux

# Couleurs pour les encadrés d'avertissement et d'information
COLOR_NOTICE_BG = "#E8EEF4"
COLOR_CALLOUT_BG = "#EAF2F8"

# Palette de couleurs pour les camemberts et graphiques Matplotlib
COLOR_CHART_1 = COLOR_PRIMARY  # Bleu principal
COLOR_CHART_2 = "#53AB60"   # Vert écologie
COLOR_CHART_3 = "#F4A261"   # Orangé d'attention
COLOR_CHART_4 = "#D95C4A"   # Rouge doux d'infraction / non-conformité
COLOR_CHART_5 = "#90BF83"   # Vert clair
COLOR_CHART_6 = "#4296CE"   # Bleu ciel
CHART_PIE_COLORS = [COLOR_CHART_1, COLOR_CHART_2, COLOR_CHART_3, COLOR_CHART_4, COLOR_CHART_5, COLOR_CHART_6]
CHART_BAR_GROUPED_COLORS = [COLOR_CHART_1, COLOR_CHART_2, COLOR_CHART_3, COLOR_CHART_4]

# Couleurs thématiques par thématique/domaine d'intervention de l'OFB
COLOR_MAP_DOMAINE = {
    "Assurer la protection des espèces animales et végétales": "#E74C3C",  # Rouge
    "Espaces protégés et protection des milieux et du cadre de vie": "#1E8449",  # Vert
    "Préservation des milieux aquatiques": "#008080",  # Teal/Bleu vert
    "Gestion qualitative de la ressource en eau": "#2980B9",  # Bleu
    "Gestion quantitative de l'eau": "#00B4D8",  # Cyan
    "Sujets transversaux": "#D35400",  # Marron/Orangé
    "Sécurité publique et Prévention des inondations": "#6C5CE7",  # Violet
    "Hors domaine": "#7F8C8D",  # Gris
}

# ========================================================================================
# DIMENSIONS, MARGES ET ESPACEMENTS PAR DÉFAUT DU DOCUMENT PDF
# ========================================================================================

PAGE_W, PAGE_H = A4  # Dimensions A4 (210mm x 297mm)
MARGIN_LEFT = 7.0 * mm  # Marge gauche optimisée
MARGIN_RIGHT = 7.0 * mm  # Marge droite optimisée
MARGIN_BOTTOM = 22 * mm  # Marge basse réservée au pied de page
MARGIN_TOP = 14 * mm  # Marge haute réservée au bandeau de titre

# Grille d'espacements verticaux standards
SPACING_XXS = 0.5 * mm
SPACING_XS = 1.0 * mm
SPACING_S = 1.5 * mm
SPACING_M = 2.0 * mm
SPACING_L = 4.0 * mm

_HEADER_LINE_STEP = 3.2 * mm
_HEADER_GAP_RULE = 2.5 * mm
_HEADER_GAP_CONTENT = 4.5 * mm


def header_layout_metrics(n_header_lines: int) -> tuple[float, float]:
    """Calcule la hauteur occupée par l'en-tête de page selon le nombre de lignes affichées."""
    n = max(1, min(int(n_header_lines), 3))
    text_block_h = n * _HEADER_LINE_STEP
    rule_from_top = text_block_h + _HEADER_GAP_RULE
    margin_top = rule_from_top + _HEADER_GAP_CONTENT
    return rule_from_top, margin_top


# Logo texte ASCII affiché lors du lancement dans la console
ASCII_LOGO_OFB = r"""
  OOOOOOO   FFFFFFF   BBBBBBB 
  OOOOOOO   FFFFFFF   BBBBBBB 
  OO   OO   FF        BB   BB
  OO   OO   FFFFFF    BBBBBBB
  OO   OO   FFFFFF    BBBBBBB
  OO   OO   FF        BB   BB
  OOOOOOO   FF        BBBBBBB
  OOOOOOO   FF        BBBBBBB

   OFFICE FRANÇAIS
 DE LA BIODIVERSITÉ
"""


def print_ascii_logo_ofb() -> None:
    """Affiche le logo OFB en art ASCII dans la console au démarrage."""
    print(ASCII_LOGO_OFB)


def _ref_img(name: str) -> Path:
    """Construit le chemin absolu vers une image de référence de la charte."""
    ref_dir = Path(__file__).resolve().parents[2] / "ref" / "programme"
    return ref_dir / "modele_ofb" / "word" / "media" / name


# Images et filigranes officiels de la charte OFB
IMG_BANNER = _ref_img("image5.jpg")
IMG_TITLE_DECO = _ref_img("image6.jpeg")
IMG_FILIGRANE = _ref_img("image3.jpeg")
IMG_FILIGRANE_ALT = _ref_img("image4.png")

IMG_LOGO_BANNER = IMG_BANNER
IMG_TITLE_PAGE_DECO = IMG_TITLE_DECO
IMG_FOOTER_DECO = IMG_FILIGRANE
IMG_BACKGROUND = IMG_FILIGRANE_ALT

CHARTE_ASSET_DEFAULT_FILES: dict[str, str] = {
    "banner": "image5.jpg",
    "title_page_deco": "image6.jpeg",
    "watermark": "image3.jpeg",
    "footer_deco": "image4.jpeg",
}


def charte_asset_path(
    assets_cfg: dict | None,
    key: str,
    default_filename: str,
    *,
    fallback: Path | None = None,
) -> Path:
    """Résout le chemin d'une image de la charte à partir de la configuration YAML."""
    name = default_filename
    if isinstance(assets_cfg, dict):
        raw = assets_cfg.get(key)
        if raw is not None and str(raw).strip():
            name = str(raw).strip()
    path = _ref_img(name)
    if path.exists():
        return path
    if fallback is not None and fallback.exists():
        return fallback
    return path


# ========================================================================================
# ENREGISTREMENT DES POLICES DE CARACTÈRES
# ========================================================================================

def _register_fonts() -> str:
    """Enregistre les polices (Marianne / Arial / LiberationSans) dans ReportLab.

    Tente d'abord de charger la police officielle de l'État 'Marianne'.
    En cas d'absence, bascule sur Arial (Windows) ou LiberationSans (Linux).
    """
    # 1. Recherche de la police officielle de l'État 'Marianne'
    marianne_dirs = [
        Path("/usr/share/fonts/truetype/marianne"),
        Path("/usr/share/fonts/opentype/marianne"),
        Path("/usr/local/share/fonts/marianne"),
        Path("~/.local/share/fonts/marianne").expanduser(),
    ]

    for marianne_dir in marianne_dirs:
        if marianne_dir.exists():
            regular_fonts = list(marianne_dir.glob("*Regular*"))
            bold_fonts = list(marianne_dir.glob("*Bold*"))
            italic_fonts = list(marianne_dir.glob("*Italic*"))
            bolditalic_fonts = list(marianne_dir.glob("*Bold*Italic*"))

            regular = next((f for f in regular_fonts if "Regular" in f.name and "Italic" not in f.name), None)
            bold = next((f for f in bold_fonts if "Bold" in f.name and "Italic" not in f.name and "ExtraBold" not in f.name and "Medium" not in f.name), None)
            italic = next((f for f in italic_fonts if "Italic" in f.name and "Bold" not in f.name and "Regular" in f.name), None)
            bolditalic = next((f for f in bolditalic_fonts if "Bold" in f.name and "Italic" in f.name and "ExtraBold" not in f.name), None)

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

    # 2. Fallback sur Arial (sur systèmes Windows)
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

    # 3. Fallback sur LiberationSans (Linux)
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

    # 4. Fallback ultime : Helvetica (police de base ReportLab)
    return "Helvetica"


# Famille de police active
FONT_FAMILY = _register_fonts()

# Styles prédéfinis de cellules de tableaux
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


# ========================================================================================
# FEUILLE DE STYLES DU RAPPORT PDF (TITRES, CORPS, DELEGATION)
# ========================================================================================

def _get_styles(typography_config: dict | None = None):
    """Construit l'ensemble des styles de paragraphe (ParagraphStyle) aux normes de l'OFB."""
    ss = getSampleStyleSheet()

    sub_italic = True
    if typography_config is not None:
        sub_italic = bool(typography_config.get("subsections_italic", True))

    h_font = f"{FONT_FAMILY}-BoldItalic" if sub_italic else f"{FONT_FAMILY}-Bold"

    styles = {
        "Title": ParagraphStyle(
            "OFBTitle",
            parent=ss["Title"],
            fontName=f"{FONT_FAMILY}-Bold",
            fontSize=26,
            leading=36,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_CENTER,
            spaceAfter=3 * mm,
        ),
        "Heading1": ParagraphStyle(
            "OFBH1",
            parent=ss["Heading1"],
            fontName=f"{FONT_FAMILY}-Bold",
            fontSize=18,
            leading=24,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_LEFT,
            spaceBefore=1 * mm,
            spaceAfter=1 * mm,
            keepWithNext=1,
        ),
        "Heading2": ParagraphStyle(
            "OFBH2",
            parent=ss["Heading2"],
            fontName=h_font,
            fontSize=14,
            leading=18,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_LEFT,
            spaceBefore=1 * mm,
            spaceAfter=1 * mm,
            keepWithNext=1,
        ),
        "Heading3": ParagraphStyle(
            "OFBH3",
            parent=ss["Heading3"],
            fontName=h_font,
            fontSize=12,
            leading=15,
            textColor=rl_colors.HexColor(COLOR_GREY),
            alignment=TA_LEFT,
            spaceBefore=0.5 * mm,
            spaceAfter=0.5 * mm,
            keepWithNext=1,
        ),
        "BodyText": ParagraphStyle(
            "OFBBody",
            parent=ss["BodyText"],
            fontName=FONT_FAMILY,
            fontSize=10,
            leading=14,
            textColor=rl_colors.black,
            alignment=TA_JUSTIFY,
            spaceBefore=0.5 * mm,
            spaceAfter=1 * mm,
        ),
        "BodySmall": ParagraphStyle(
            "OFBBodySmall",
            parent=ss["BodyText"],
            fontName=FONT_FAMILY,
            fontSize=8,
            leading=11,
            textColor=rl_colors.HexColor("#666666"),
            alignment=TA_LEFT,
            spaceBefore=0.5 * mm,
            spaceAfter=0.5 * mm,
        ),
        "TableCaption": ParagraphStyle(
            "OFBTableCaption",
            parent=ss["BodyText"],
            fontName=f"{FONT_FAMILY}-Bold",
            fontSize=10,
            leading=13,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_LEFT,
            spaceBefore=2 * mm,
            spaceAfter=1 * mm,
            keepWithNext=0,
        ),
        "FigureCaption": ParagraphStyle(
            "OFBFigureCaption",
            parent=ss["BodyText"],
            fontName=f"{FONT_FAMILY}-Italic",
            fontSize=9,
            leading=13,
            textColor=rl_colors.HexColor(COLOR_PRIMARY),
            alignment=TA_CENTER,
            spaceBefore=0,
            spaceAfter=2 * mm,
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


# ========================================================================================
# ANIMATEUR DE CONSOLE (SPINNER DE PATIENCE POUR EXECUTION CLI)
# ========================================================================================

class Spinner:
    """Affiche une animation de chargement en console pendant l'exécution d'une tâche de calcul."""

    def __init__(self, message: str = "Traitement des données en cours. Patience...") -> None:
        self.message = message
        self._stop_event = threading.Event()
        self._thread = threading.Thread(target=self._spin, daemon=True)

    def _spin(self) -> None:
        """Boucle d'animation affichant le texte caractère par caractère."""
        message = self.message
        appear_delay = 0.06
        disappear_delay = 0.03
        pause_full = 1.5
        pause_empty = 0.3
        msg_len = len(message)

        while not self._stop_event.is_set():
            # Apparition progressive du texte
            for i in range(1, msg_len + 1):
                if self._stop_event.is_set():
                    break
                sys.stdout.write("\r" + message[:i])
                sys.stdout.flush()
                time.sleep(appear_delay)

            if self._stop_event.is_set():
                break

            # Pause d'affichage du message complet
            elapsed = 0.0
            while elapsed < pause_full and not self._stop_event.is_set():
                time.sleep(0.1)
                elapsed += 0.1

            if self._stop_event.is_set():
                break

            # Effacement progressif du texte
            for i in range(msg_len, -1, -1):
                if self._stop_event.is_set():
                    break
                sys.stdout.write("\r" + message[:i] + " " * (msg_len - i))
                sys.stdout.flush()
                time.sleep(disappear_delay)

            if self._stop_event.is_set():
                break

            # Pause sur ligne vide
            elapsed = 0.0
            while elapsed < pause_empty and not self._stop_event.is_set():
                time.sleep(0.1)
                elapsed += 0.1

    def __enter__(self) -> "Spinner":
        """Démarre l'animation lors de l'entrée dans un bloc `with Spinner():`."""
        if not sys.stdout.isatty():
            return self

        import logging
        is_debug = False
        for logger_name in ("ofbilan", "core"):
            logger = logging.getLogger(logger_name)
            for h in logger.handlers:
                if isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler):
                    if h.level <= logging.DEBUG:
                        is_debug = True
                        break
        if is_debug:
            return self

        _enable_windows_vt100()
        sys.stdout.write("\n")
        sys.stdout.flush()
        self._thread.start()
        return self

    def __exit__(self, _exc_type, _exc_val, _exc_tb) -> None:
        """Arrête l'animation et efface la ligne lors de la sortie du bloc `with`."""
        if not sys.stdout.isatty():
            return

        import logging
        is_debug = False
        for logger_name in ("ofbilan", "core"):
            logger = logging.getLogger(logger_name)
            for h in logger.handlers:
                if isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler):
                    if h.level <= logging.DEBUG:
                        is_debug = True
                        break
        if is_debug:
            return

        self._stop_event.set()
        self._thread.join()
        sys.stdout.write("\r" + " " * len(self.message) + "\r")
        sys.stdout.flush()