# Copyright (C) 2026 Aguirre MAURIN
"""Ré-exportation de transition vers engine_pdf_helpers.py pour rétro-compatibilité."""
from core.engine.engine_pdf_helpers import (
    truncate_with_dash,
    nb_non_conformes_brut,
    pct_table_cell,
    get_region_name_for_footer,
)

__all__ = [
    "truncate_with_dash",
    "nb_non_conformes_brut",
    "pct_table_cell",
    "get_region_name_for_footer",
]