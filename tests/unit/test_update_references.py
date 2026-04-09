"""Tests for the xls_update_references formula rewriting logic."""

from __future__ import annotations

from excel_agent.tools.cells.xls_update_references import (
    _apply_updates_to_formula,
    _normalize_ref,
)


class TestNormalizeRef:
    def test_strips_dollars(self) -> None:
        assert _normalize_ref("$A$1") == "A1"

    def test_uppercases(self) -> None:
        assert _normalize_ref("sheet1!a1") == "SHEET1!A1"

    def test_passthrough(self) -> None:
        assert _normalize_ref("Sheet1!B5") == "SHEET1!B5"


class TestApplyUpdates:
    def test_simple_replacement(self) -> None:
        update_map = {"SHEET1!A1": "Sheet1!B5"}
        result = _apply_updates_to_formula("=A1*2", update_map, "Sheet1")
        assert "B5" in result

    def test_cross_sheet_replacement(self) -> None:
        update_map = {"SHEET1!C1": "Sheet2!D1"}
        result = _apply_updates_to_formula("=Sheet1!C1+10", update_map, "Sheet2")
        assert "Sheet2!D1" in result

    def test_no_match_unchanged(self) -> None:
        update_map = {"SHEET1!Z99": "Sheet1!Z100"}
        result = _apply_updates_to_formula("=A1+B1", update_map, "Sheet1")
        assert result == "=A1+B1"

    def test_multiple_replacements(self) -> None:
        update_map = {
            "SHEET1!A1": "Sheet1!A5",
            "SHEET1!B1": "Sheet1!B5",
        }
        result = _apply_updates_to_formula("=A1+B1", update_map, "Sheet1")
        assert "A5" in result
        assert "B5" in result

    def test_formula_with_functions(self) -> None:
        update_map = {"SHEET1!A1": "Sheet1!A10"}
        result = _apply_updates_to_formula("=SUM(A1,B1)", update_map, "Sheet1")
        assert "A10" in result
        # B1 should be unchanged
        assert "B1" in result

    def test_preserves_local_ref_style(self) -> None:
        """If original ref had no sheet prefix, the replacement shouldn't add one
        when the reference stays on the same sheet."""
        update_map = {"SHEET1!A1": "Sheet1!C3"}
        result = _apply_updates_to_formula("=A1*2", update_map, "Sheet1")
        # Should be =C3*2 (not =Sheet1!C3*2) because original was local
        assert result == "=C3*2"
