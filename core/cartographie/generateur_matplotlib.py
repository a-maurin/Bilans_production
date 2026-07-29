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
MODULE : GENERATEUR CARTOGRAPHIQUE MATPLOTLIB / GEOPANDAS (`generateur_matplotlib.py`)
========================================================================================
Ce module sert de moteur cartographique de secours ("fallback"). Lorsque QGIS n'est pas
installé ou qu'un interpréteur PyQGIS complet n'est pas disponible, ce composant prend
le relais pour tracer une carte statistique via Matplotlib et GeoPandas.

Fonctions principales :
  1. `charger_couche_pochoir()` : charge le contour du département depuis les SHP/GeoJSON.
  2. `tracer_carte()` : dessine le fond de carte et la géométrie départementale.
  3. `tracer_cartouche()` : compose le cartouche latéral (titre, dates, charte OFB).
  4. `exporter_carte_matplotlib()` : exporte l'image PNG au format A4 Paysage à 300 DPI.
========================================================================================
"""
import logging
from pathlib import Path
import matplotlib.pyplot as plt

try:
    import geopandas as gpd
except ImportError:
    gpd = None

logger = logging.getLogger(__name__)

COLOR_PRIMARY = "#003A76"
FONT_FAMILY = "sans-serif" # Fallback generique

def charger_couche_pochoir(dept_code: str, project_root: Path):
    if gpd is None:
        return None
    from core.cartographie.pochoir_helper import load_department_gdf
    try:
        return load_department_gdf(dept_code, project_root=project_root)
    except Exception as e:
        logger.warning(f"Impossible de charger le pochoir departement {dept_code}: {e}")
        return gpd.GeoDataFrame(columns=["geometry"], geometry="geometry")

def tracer_carte(ax, pochoir_gdf, layers_to_render: list):
    """Dessine le fond et les donnees sur l'axe principal."""
    # Fond general clair
    ax.set_facecolor("#f0f4f8")
    
    if pochoir_gdf is None or gpd is None:
        ax.axis("off")
        ax.text(0.5, 0.5, "Cartographie indisponible.\nVeuillez mettre à jour votre version de Qgis : https://qgis.org/download/ ,\nsélectionner  Get OSGeo4W Installer.", 
                ha='center', va='center', fontsize=12, color="#333333", transform=ax.transAxes, wrap=True)
        return

    # Dessin du pochoir (departement)
    if not pochoir_gdf.empty:
        pochoir_gdf.plot(ax=ax, color="white", edgecolor="#CCCCCC", linewidth=1.5)
        # Ajuster emprise (extent)
        minx, miny, maxx, maxy = pochoir_gdf.total_bounds
        pad_x, pad_y = (maxx - minx)*0.05, (maxy - miny)*0.05
        ax.set_xlim(minx - pad_x, maxx + pad_x)
        ax.set_ylim(miny - pad_y, maxy + pad_y)

    ax.axis("off")

def tracer_cartouche(ax, titre: str, dept_name: str, date_deb: str, date_fin: str):
    """Dessine le panneau lateral."""
    ax.set_facecolor("white")
    ax.axis("off")
    
    # Ligne de separation
    ax.axvline(0, color=COLOR_PRIMARY, linewidth=2)
    
    # Textes
    ax.text(0.1, 0.9, "Office Français\nde la Biodiversité", fontsize=12, fontweight='bold', color=COLOR_PRIMARY, transform=ax.transAxes)
    ax.text(0.1, 0.8, titre, fontsize=16, fontweight='bold', color=COLOR_PRIMARY, transform=ax.transAxes)
    ax.text(0.1, 0.75, dept_name, fontsize=12, color="#333333", transform=ax.transAxes)
    ax.text(0.1, 0.70, f"Du {date_deb}\nau {date_fin}", fontsize=10, color="#666666", transform=ax.transAxes)
    
    # Footer
    ax.text(0.1, 0.05, "Sources: OFB, IGN\nProjection: RGF93", fontsize=8, color="#999999", transform=ax.transAxes)

def exporter_carte_matplotlib(prof, output_path: Path, dept_code: str, layers_to_render: list, project_root: Path):
    """Point d'entree principal du generateur."""
    logger.info(f"Génération Matplotlib (Fallback) pour le profil {prof.id}")
    
    fig = plt.figure(figsize=(11.69, 8.27), dpi=300) # A4 Landscape
    gs = fig.add_gridspec(1, 2, width_ratios=[4, 1], wspace=0)
    
    ax_map = fig.add_subplot(gs[0])
    ax_side = fig.add_subplot(gs[1])
    
    pochoir = charger_couche_pochoir(dept_code, project_root)
    tracer_carte(ax_map, pochoir, layers_to_render)
    
    titre = getattr(prof, "title_main", "") or getattr(prof, "title", "Bilan")
    tracer_cartouche(ax_side, titre, f"Département {dept_code}", prof.date_deb, prof.date_fin)
    
    fig.savefig(output_path, bbox_inches="tight", dpi=300)
    plt.close(fig)
    return True