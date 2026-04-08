"""Tests for unified range parsing (RangeSerializer)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_agent.core.serializers import (
    CellCoordinate,
    RangeCoordinate,
    RangeSerializer,
    col_letter_to_number,
    col_number_to_letter,
)
from excel_agent.utils.exceptions import ValidationError


class TestColConversions:
    """Tests for column letter ↔ number conversions."""

    @pytest.mark.parametrize(
        ("letter", "expected"),
        [("A", 1), ("Z", 26), ("AA", 27), ("AZ", 52), ("XFD", 16384)],
    )
    def test_letter_to_number(self, letter: str, expected: int) -> None:
        assert col_letter_to_number(letter) == expected

    @pytest.mark.parametrize(
        ("number", "expected"),
        [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (16384, "XFD")],
    )
    def test_number_to_letter(self, number: int, expected: str) -> None:
        assert col_number_to_letter(number) == expected

    @pytest.mark.parametrize("n", [1, 26, 27, 52, 256, 702, 16384])
    def test_roundtrip_number(self, n: int) -> None:
        assert col_letter_to_number(col_number_to_letter(n)) == n

    @pytest.mark.parametrize("letter", ["A", "Z", "AA", "AZ", "XFD"])
    def test_roundtrip_letter(self, letter: str) -> None:
        assert col_number_to_letter(col_letter_to_number(letter)) == letter


class TestCellCoordinate:
    """Tests for CellCoordinate dataclass."""

    def test_valid(self) -> None:
        c = CellCoordinate(row=1, col=1)
        assert c.row == 1 and c.col == 1

    def test_invalid_row(self) -> None:
        with pytest.raises(ValidationError):
            CellCoordinate(row=0, col=1)

    def test_invalid_col(self) -> None:
        with pytest.raises(ValidationError):
            CellCoordinate(row=1, col=0)


class TestRangeSerializerA1:
    """Tests for A1 notation parsing."""

    def setup_method(self) -> None:
        self.s = RangeSerializer()

    def test_single_cell(self) -> None:
        r = self.s.parse("A1")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1)
        assert r.is_single_cell

    def test_single_cell_absolute(self) -> None:
        r = self.s.parse("$A$1")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1)

    def test_range(self) -> None:
        r = self.s.parse("A1:C10")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)

    def test_range_absolute(self) -> None:
        r = self.s.parse("$A$1:$C$10")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)

    def test_with_sheet_unquoted(self) -> None:
        r = self.s.parse("Sheet1!A1:C10")
        assert r == RangeCoordinate(sheet="Sheet1", min_row=1, min_col=1, max_row=10, max_col=3)

    def test_with_sheet_quoted(self) -> None:
        r = self.s.parse("'Sheet Name'!A1")
        assert r == RangeCoordinate(sheet="Sheet Name", min_row=1, min_col=1)

    def test_multi_letter_column(self) -> None:
        r = self.s.parse("AA100")
        assert r == RangeCoordinate(sheet=None, min_row=100, min_col=27)

    def test_default_sheet(self) -> None:
        r = self.s.parse("A1", default_sheet="Data")
        assert r.sheet == "Data"


class TestRangeSerializerR1C1:
    """Tests for R1C1 notation parsing."""

    def setup_method(self) -> None:
        self.s = RangeSerializer()

    def test_single_cell(self) -> None:
        r = self.s.parse("R1C1")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1)

    def test_range(self) -> None:
        r = self.s.parse("R1C1:R10C3")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)

    def test_case_insensitive(self) -> None:
        r = self.s.parse("r5c10")
        assert r == RangeCoordinate(sheet=None, min_row=5, min_col=10)


class TestRangeSerializerFullRowCol:
    """Tests for full row and column references."""

    def setup_method(self) -> None:
        self.s = RangeSerializer()

    def test_full_column(self) -> None:
        r = self.s.parse("A:C")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=None, max_col=3)

    def test_full_row(self) -> None:
        r = self.s.parse("1:10")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=None)


class TestRangeSerializerNamedRange:
    """Tests for named range resolution (requires workbook)."""

    def test_named_range_resolution(self, sample_workbook: Path) -> None:
        from openpyxl import load_workbook

        wb = load_workbook(str(sample_workbook))
        s = RangeSerializer(workbook=wb)
        r = s.parse("SalesData")
        assert r.sheet == "Sheet3"
        assert r.min_row == 1
        assert r.min_col == 1


class TestRangeSerializerInvalid:
    """Tests for error handling on invalid inputs."""

    def setup_method(self) -> None:
        self.s = RangeSerializer()

    def test_empty_string(self) -> None:
        with pytest.raises(ValidationError, match="Empty range"):
            self.s.parse("")

    def test_garbage(self) -> None:
        with pytest.raises(ValidationError, match="Cannot parse"):
            self.s.parse("not_a_valid_range!!!")


class TestToA1:
    """Tests for RangeCoordinate → A1 conversion."""

    def test_single_cell(self) -> None:
        coord = RangeCoordinate(sheet=None, min_row=1, min_col=1)
        assert RangeSerializer.to_a1(coord) == "A1"

    def test_range(self) -> None:
        coord = RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)
        assert RangeSerializer.to_a1(coord) == "A1:C10"

    def test_with_sheet(self) -> None:
        coord = RangeCoordinate(sheet="Sheet1", min_row=1, min_col=1, max_row=10, max_col=3)
        assert RangeSerializer.to_a1(coord) == "Sheet1!A1:C10"

    def test_with_quoted_sheet(self) -> None:
        coord = RangeCoordinate(sheet="My Sheet", min_row=1, min_col=1)
        assert RangeSerializer.to_a1(coord) == "'My Sheet'!A1"


class TestToR1C1:
    """Tests for RangeCoordinate → R1C1 conversion."""

    def test_single_cell(self) -> None:
        coord = RangeCoordinate(sheet=None, min_row=5, min_col=10)
        assert RangeSerializer.to_r1c1(coord) == "R5C10"

    def test_range(self) -> None:
        coord = RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)
        assert RangeSerializer.to_r1c1(coord) == "R1C1:R10C3"


class TestRoundtrip:
    """Tests for parse → to_a1 roundtrip fidelity."""

    @pytest.mark.parametrize(
        "input_str",
        ["A1", "A1:C10", "Z26", "AA1:AZ100", "A1:A1"],
    )
    def test_a1_roundtrip(self, input_str: str) -> None:
        s = RangeSerializer()
        coord = s.parse(input_str)
        output = s.to_a1(coord)
        # Parse the output again — should produce same coord
        coord2 = s.parse(output)
        assert coord == coord2
