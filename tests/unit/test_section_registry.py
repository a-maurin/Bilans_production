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
from __future__ import annotations


def test_section_registry_register_and_render() -> None:
    from core.engine.registre_sections_pdf import SectionRegistry

    seen: list[str] = []

    def _render(ctx: dict) -> None:
        seen.append(str(ctx.get("x", "")))

    reg = SectionRegistry()
    reg.register("sec_test", _render)
    reg.render("sec_test", {"x": "ok"})
    assert seen == ["ok"]


def test_section_registry_missing_raises() -> None:
    from core.engine.registre_sections_pdf import SectionRegistry

    reg = SectionRegistry()
    try:
        reg.render("missing", {})
    except KeyError as e:
        assert "missing" in str(e).lower() or "aucun" in str(e).lower()
    else:
        raise AssertionError("expected KeyError")


def test_section_registry_render_many_skips_unknown_by_default() -> None:
    from core.engine.registre_sections_pdf import SectionRegistry

    seen: list[str] = []

    def _render_a(ctx: dict) -> None:
        seen.append(f"a:{ctx.get('v', '')}")

    reg = SectionRegistry()
    reg.register("a", _render_a)
    reg.render_many(["a", "missing"], {"v": "ok"})
    assert seen == ["a:ok"]


def test_section_registry_render_many_raises_when_requested() -> None:
    from core.engine.registre_sections_pdf import SectionRegistry

    reg = SectionRegistry()
    try:
        reg.render_many(["missing"], {}, skip_unknown=False)
    except KeyError:
        pass
    else:
        raise AssertionError("expected KeyError")