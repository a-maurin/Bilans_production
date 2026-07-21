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

#
"""Constantes et fichiers par défaut de la charte OFB."""

from __future__ import annotations

from core.common.ofb_charte import (
    CHARTE_ASSET_DEFAULT_FILES,
    IMG_BACKGROUND,
    IMG_BANNER,
    IMG_FILIGRANE,
    IMG_FILIGRANE_ALT,
    IMG_FOOTER_DECO,
    IMG_LOGO_BANNER,
    IMG_TITLE_DECO,
    IMG_TITLE_PAGE_DECO,
)


def test_charte_asset_default_files_match_yaml_keys() -> None:
    assert set(CHARTE_ASSET_DEFAULT_FILES) == {
        "banner",
        "title_page_deco",
        "watermark",
        "footer_deco",
    }
    assert CHARTE_ASSET_DEFAULT_FILES["banner"] == "image5.jpg"
    assert CHARTE_ASSET_DEFAULT_FILES["watermark"] == "image3.jpeg"


def test_charte_legacy_aliases_point_to_canonical_constants() -> None:
    assert IMG_LOGO_BANNER == IMG_BANNER
    assert IMG_TITLE_PAGE_DECO == IMG_TITLE_DECO
    assert IMG_FOOTER_DECO == IMG_FILIGRANE
    assert IMG_BACKGROUND == IMG_FILIGRANE_ALT