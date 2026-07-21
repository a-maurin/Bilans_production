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
import pandas as pd

from core.engine.generation_pdf_synthese import (
    _KEY_FIGURES_GRAIN_NOTE,
    _build_usager_theme_table_rows,
    _resultats_controles_pie_data,
    _rollup_small_categories,
    _wrap_table_label,
)
from core.engine.generation_pdf_synthese_brochure import (
    _brochure_methodology_html,
    _theme_pct_strings_brochure,
)


def test_key_figures_grain_note_explains_difference_briefly() -> None:
    note = _KEY_FIGURES_GRAIN_NOTE

    assert "points de contrôle" in note
    assert "fiche de contrôle" in note
    assert "peuvent donc être inférieurs ou supérieurs" in note


def test_wrap_table_label_inserts_line_breaks_without_truncating() -> None:
    wrapped = _wrap_table_label("Contrôles espaces protégés et protection des milieux")

    assert "<br/>" in wrapped
    assert "Contrôles espaces protégés" in wrapped
    assert "protection des milieux" in wrapped


def test_resultats_controles_pie_data_uses_four_expected_categories() -> None:
    df = pd.DataFrame(
        [
            {"resultat": "Conforme", "nb": 10},
            {"resultat": "Infraction", "nb": 3},
            {"resultat": "Manquement", "nb": 2},
            {"resultat": "En attente", "nb": 1},
            {"resultat": "Non-conforme", "nb": 5},
        ]
    )

    out = _resultats_controles_pie_data(df)

    assert out == {
        "Conforme": 10,
        "Infraction": 3,
        "Manquement": 2,
        "En attente": 1,
    }


def test_rollup_small_categories_adds_last_other_row() -> None:
    df = pd.DataFrame(
        [
            {"theme": "Milieux aquatiques", "nb_localisations": 500, "nb_pej_hors_controle": 0, "nb_total": 500},
            {"theme": "Chasse", "nb_localisations": 250, "nb_pej_hors_controle": 5, "nb_total": 255},
            {"theme": "Déchets", "nb_localisations": 70, "nb_pej_hors_controle": 0, "nb_total": 70},
            {"theme": "Bruit", "nb_localisations": 6, "nb_pej_hors_controle": 0, "nb_total": 6},
            {"theme": "Publicité", "nb_localisations": 4, "nb_pej_hors_controle": 1, "nb_total": 5},
        ]
    )

    out = _rollup_small_categories(
        df,
        label_col="theme",
        other_label="Autres thèmes de contrôle",
        value_col="nb_total",
        min_pct=0.01,
        sum_cols=["nb_localisations", "nb_pej_hors_controle", "nb_total"],
    )

    assert out is not None
    assert out["theme"].tolist() == [
        "Milieux aquatiques",
        "Chasse",
        "Déchets",
        "Autres thèmes de contrôle",
    ]
    assert int(out.iloc[-1]["nb_localisations"]) == 10
    assert int(out.iloc[-1]["nb_pej_hors_controle"]) == 1
    assert int(out.iloc[-1]["nb_total"]) == 11


def test_theme_pct_strings_brochure_use_global_total() -> None:
    values = [326, 162, 107, 61, 36]

    out = _theme_pct_strings_brochure(values, total_value=980)

    assert out == ["33 %", "17 %", "11 %", "6 %", "4 %"]


def test_build_usager_theme_table_rows_keeps_full_theme_label() -> None:
    df = pd.DataFrame(
        [
            {
                "theme": "Protection des milieux naturels et de la biodiversite remarquable",
                "nb_effectifs": 3,
                "nb_pej_suite_controle": 1,
                "nb_pej_hors_controle": 2,
                "nb_total": 6,
            }
        ]
    )

    rows = _build_usager_theme_table_rows(df)

    assert rows[1][0] == "Protection des milieux naturels et de la biodiversite remarquable"


def test_brochure_methodology_html_includes_realisation() -> None:
    html = _brochure_methodology_html(
        date_deb=pd.Timestamp("2025-01-01"),
        date_fin=pd.Timestamp("2025-12-31"),
        ventilation_mode="globale",
        diffusion="externe",
    )

    assert "Réalisation" in html
    assert "service départemental de la Côte d'Or" in html