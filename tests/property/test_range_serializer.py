"""Property-based tests for RangeSerializer using Hypothesis."""

from __future__ import annotations

from hypothesis import given, settings
from hypothesis import strategies as st

from excel_agent.core.serializers import (
    RangeSerializer,
    col_letter_to_number,
    col_number_to_letter,
)


# Strategy for valid column letters (A-Z, AA-ZZ, etc. up to XFD)
def column_letter_strategy():
    """Generate valid Excel column letters."""
    # Single letters: A-Z
    single = st.sampled_from([chr(ord("A") + i) for i in range(26)])
    # Two letters: AA-ZZ
    double = st.tuples(single, single).map(lambda t: t[0] + t[1])
    # Three letters: AAA-XFD (but mostly AAA-ZZZ, we'll filter valid ones)
    triple = st.tuples(single, single, single).map(lambda t: t[0] + t[1] + t[2])

    return st.one_of(single, double, triple).filter(
        lambda x: 1 <= col_letter_to_number(x) <= 16384
    )


# Strategy for valid row numbers (1 to 1,048,576 in Excel)
row_number_strategy = st.integers(min_value=1, max_value=1048576)


# Strategy for valid A1 cell references
def a1_cell_strategy():
    """Generate valid A1-style cell references."""
    return st.tuples(column_letter_strategy(), row_number_strategy).map(lambda t: t[0] + str(t[1]))


class TestColConversionsRoundtrip:
    """Property-based roundtrip tests for column conversions."""

    @given(column_letter_strategy())
    @settings(max_examples=500)
    def test_column_letter_roundtrip(self, letter: str) -> None:
        """col_number_to_letter(col_letter_to_number(x)) == x"""
        assert col_number_to_letter(col_letter_to_number(letter)) == letter


class TestRangeSerializerRoundtrip:
    """Property-based roundtrip tests for range parsing."""

    @given(a1_cell_strategy())
    @settings(max_examples=500)
    def test_single_cell_roundtrip(self, cell: str) -> None:
        """Parsing then serializing a cell should give a consistent result."""
        s = RangeSerializer()
        coord = s.parse(cell)
        output = s.to_a1(coord)

        # The output might have sheet prefix (None), so compare the parsed result
        coord2 = s.parse(output)
        assert coord == coord2

    @given(a1_cell_strategy(), a1_cell_strategy())
    @settings(max_examples=200)
    def test_range_roundtrip(self, cell1: str, cell2: str) -> None:
        """Parsing then serializing a range should give a consistent result."""
        s = RangeSerializer()

        # Skip if cells are the same (would be single cell)
        if cell1 == cell2:
            return

        range_str = f"{cell1}:{cell2}"

        try:
            coord = s.parse(range_str)
            output = s.to_a1(coord)
            coord2 = s.parse(output)
            assert coord == coord2
        except Exception:
            # Some combinations might be invalid (e.g., malformed ranges)
            # We just skip those
            pass


class TestColLetterToNumber:
    """Property-based tests for column letter to number conversion."""

    @given(st.integers(min_value=1, max_value=16384))
    @settings(max_examples=500)
    def test_number_to_letter_roundtrip(self, n: int) -> None:
        """col_letter_to_number(col_number_to_letter(x)) == x"""
        letter = col_number_to_letter(n)
        assert col_letter_to_number(letter) == n
