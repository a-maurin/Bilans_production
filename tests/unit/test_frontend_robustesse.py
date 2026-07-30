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
"""
Tests de non-régression : Mission 4 (validation inputs) et Mission 5 (cohérence profils).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))


# ── Mission 5 : Typo département 07 ──────────────────────────────────────────

def test_no_typo_dept_07_explorer_js():
    """
    Vérifie que 'Ordèche' (typo) est absent d'explorer.js et que 'Ardèche' est présent.
    """
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "explorer.js").read_text(encoding="utf-8")
    assert "Ordèche" not in source, "Typo 'Ordèche' encore présente dans explorer.js"
    assert "Ardèche" in source, "Département 07 manquant dans explorer.js"


def test_region_default_code_consistent_app_js():
    """
    Vérifie que app.js utilise 'r27' (format CLI) et non '27' pour le code région par défaut.
    """
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "app.js").read_text(encoding="utf-8")
    # Le code '27' sans préfixe 'r' ne doit pas apparaître comme valeur par défaut pour region
    assert "inputCode.value = '27'" not in source, (
        "app.js utilise '27' sans préfixe 'r' pour le code région — divergence avec explorer.js et le CLI"
    )


# ── Mission 4 : Validation form présente dans app.js ─────────────────────────

def test_validate_form_function_exists():
    """validateForm() doit exister dans app.js pour bloquer les soumissions invalides."""
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "app.js").read_text(encoding="utf-8")
    assert "function validateForm()" in source, "validateForm() absente de app.js"


def test_validate_form_checks_dates():
    """validateForm() doit vérifier la cohérence des dates."""
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "app.js").read_text(encoding="utf-8")
    assert "dateDeb > dateFin" in source, "Validation date-deb ≤ date-fin absente de validateForm()"


def test_validate_form_checks_profil():
    """validateForm() doit vérifier que le profil est dans la liste API."""
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "app.js").read_text(encoding="utf-8")
    assert "profilesList.some" in source, "Validation profil contre liste API absente de validateForm()"


def test_validate_form_checks_code():
    """validateForm() doit exiger un code géographique pour les échelles non-nationales."""
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "app.js").read_text(encoding="utf-8")
    assert "echelle !== 'national'" in source, "Validation code géo absente de validateForm()"


def test_validate_form_called_before_generate():
    """validateForm() doit être appelée dans le listener du bouton Générer."""
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "app.js").read_text(encoding="utf-8")
    # Le pattern doit apparaître dans le listener click
    assert "if (!validateForm()) return;" in source, (
        "validateForm() n'est pas appelée dans le listener du bouton Générer"
    )


# ── Mission 5 : Cohérence source profils ─────────────────────────────────────

def test_both_js_use_same_profils_api_endpoint():
    """explorer.js et app.js consomment tous deux /api/profils — source unique."""
    explorer = (Path(__file__).resolve().parents[2] / "core" / "web" / "explorer.js").read_text(encoding="utf-8")
    app = (Path(__file__).resolve().parents[2] / "core" / "web" / "app.js").read_text(encoding="utf-8")
    assert "fetch('/api/profils')" in explorer, "/api/profils absent de explorer.js"
    assert "fetch('/api/profils')" in app, "/api/profils absent de app.js"


def test_no_debug_pve_in_serveur():
    """Régression A : debug_pve.txt ne doit plus être référencé dans serveur.py."""
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "serveur.py").read_text(encoding="utf-8")
    assert "debug_pve.txt" not in source


# ── Points restants A, B, C ──────────────────────────────────────────────────

def test_point_B_form_error_msg_in_index_html():
    """
    Point B : L'élément #form-error-msg doit exister dans index.html.
    Sans lui, validateForm() dans app.js tombait en fallback silencieux.
    """
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "index.html").read_text(encoding="utf-8")
    assert 'id="form-error-msg"' in source, "#form-error-msg absent de index.html"


def test_point_C_date_validation_in_explorer_js():
    """
    Point C : explorer.js doit valider date-deb ≤ date-fin avant l'appel /api/data.
    """
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "explorer.js").read_text(encoding="utf-8")
    assert "dateDebEl.value > dateFinEl.value" in source, (
        "Validation date-deb ≤ date-fin absente de explorer.js"
    )


def test_point_A_banner_reset_on_load_start():
    """
    Point A : La bannière d'erreur est réinitialisée au début de chaque loadData().
    Vérifie la présence du reset dans le code avant le fetch.
    """
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "explorer.js").read_text(encoding="utf-8")
    assert "_errBanner.style.display = 'none'" in source, (
        "Reset bannière au début de loadData() absent de explorer.js"
    )


def test_point_A_banner_reset_on_success():
    """
    Point A : La bannière d'erreur est réinitialisée après un succès (chainé dans .then()).
    """
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "explorer.js").read_text(encoding="utf-8")
    assert "reset bannière sur succès" in source.lower() or "Point A : reset" in source, (
        "Reset bannière post-succès absent de explorer.js"
    )


def test_compare_active_title_mention_in_explorer_js():
    """
    Vérifie la présence de la mention (Comparaison années N / N-1) selon l'état de compare-active.
    """
    source = (Path(__file__).resolve().parents[2] / "core" / "web" / "explorer.js").read_text(encoding="utf-8")
    assert "compare-active" in source, "Élément compare-active absent de explorer.js"
    assert "(Comparaison années N / N-1)" in source, "Mention de comparaison (Comparaison années N / N-1) absente de explorer.js"