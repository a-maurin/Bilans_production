# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

import pytest
import pandas as pd
from core.engine.commentaires_auto import get_comment_text, format_number_fr, format_pct_fr, pluralize


class DummyContext:
    def __init__(self):
        self.presentation_cfg = {}
        self.nb_localisations = 100
        self.act_theme = pd.DataFrame([
            {"theme": "Chasse", "nb_total": 60},
            {"theme": "Pêche", "nb_total": 40},
        ])
        self.tab_resultats = pd.DataFrame([
            {"resultat": "Conforme", "nb_localisations": 80},
            {"resultat": "Non-conforme - Infraction", "nb_localisations": 15},
            {"resultat": "Non-conforme - Manquement", "nb_localisations": 5},
        ])
        self.agg_usager = pd.DataFrame([
            {"type_usager": "Particuliers", "nb_total": 70},
            {"type_usager": "Agriculteurs", "nb_total": 30},
        ])


def test_formatters():
    assert format_number_fr(1250) == "1 250"
    assert format_pct_fr(0.314) == "31 %"
    assert pluralize(1, "localisation") == "1 localisation"
    assert pluralize(5, "localisation") == "5 localisations"


def test_get_comment_text_sec21():
    ctx = DummyContext()
    res = get_comment_text("sec21_themes", ctx)
    assert res is not None
    assert "Chasse" in res
    assert "60 %" in res


def test_get_comment_text_sec22():
    ctx = DummyContext()
    res = get_comment_text("sec22_conformite", ctx)
    assert res is not None
    assert "20 %" in res
    assert "15 infractions" in res
    assert "5 manquements" in res
