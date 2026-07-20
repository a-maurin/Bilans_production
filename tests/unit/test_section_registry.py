#
# Copyright (C) 2026 Aguirre MAURIN
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

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
