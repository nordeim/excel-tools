"""Unit tests for CLI helper functions.

Tests for validate_output_path() enhancements and other CLI utilities.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_agent.utils.cli_helpers import (
    check_macro_contract,
    validate_input_path,
    validate_output_path,
    validate_output_suffix,
)
from excel_agent.utils.exit_codes import ExitCode


class TestValidateOutputPath:
    """Tests for enhanced validate_output_path()."""

    def test_valid_extension_xlsx(self, tmp_path: Path) -> None:
        """Accepts .xlsx extension."""
        output_path = tmp_path / "output.xlsx"

        result = validate_output_path(str(output_path))

        assert result == output_path

    def test_valid_extension_xlsm(self, tmp_path: Path) -> None:
        """Accepts .xlsm extension."""
        output_path = tmp_path / "output.xlsm"

        result = validate_output_path(str(output_path))

        assert result == output_path

    def test_valid_extension_xltx(self, tmp_path: Path) -> None:
        """Accepts .xltx extension."""
        output_path = tmp_path / "output.xltx"

        result = validate_output_path(str(output_path))

        assert result == output_path

    def test_valid_extension_xltm(self, tmp_path: Path) -> None:
        """Accepts .xltm extension."""
        output_path = tmp_path / "output.xltm"

        result = validate_output_path(str(output_path))

        assert result == output_path

    def test_invalid_extension_txt(self, tmp_path: Path) -> None:
        """Rejects .txt extension."""
        output_path = tmp_path / "output.txt"

        with pytest.raises(SystemExit) as exc_info:
            validate_output_path(str(output_path))

        assert exc_info.value.code == ExitCode.VALIDATION_ERROR

    def test_invalid_extension_csv(self, tmp_path: Path) -> None:
        """Rejects .csv extension."""
        output_path = tmp_path / "output.csv"

        with pytest.raises(SystemExit) as exc_info:
            validate_output_path(str(output_path))

        assert exc_info.value.code == ExitCode.VALIDATION_ERROR

    def test_custom_allowed_extensions(self, tmp_path: Path) -> None:
        """Accepts custom allowed extensions."""
        output_path = tmp_path / "output.csv"

        result = validate_output_path(str(output_path), allowed_suffixes={".csv", ".xlsx"})

        assert result == output_path

    def test_rejects_existing_file_no_overwrite(self, tmp_path: Path) -> None:
        """Rejects existing file with overwrite=False."""
        output_path = tmp_path / "existing.xlsx"
        output_path.write_text("exists")  # Create file

        with pytest.raises(SystemExit) as exc_info:
            validate_output_path(str(output_path), overwrite=False)

        assert exc_info.value.code == ExitCode.VALIDATION_ERROR

    def test_accepts_existing_file_with_overwrite(self, tmp_path: Path) -> None:
        """Accepts existing file with overwrite=True."""
        output_path = tmp_path / "existing.xlsx"
        output_path.write_text("exists")  # Create file

        result = validate_output_path(str(output_path), overwrite=True)

        assert result == output_path

    def test_creates_parent_directories(self, tmp_path: Path) -> None:
        """Creates parent directories when create_parents=True."""
        output_path = tmp_path / "nested" / "deep" / "output.xlsx"

        result = validate_output_path(str(output_path), create_parents=True)

        assert result == output_path
        assert output_path.parent.exists()

    def test_rejects_missing_parent(self, tmp_path: Path) -> None:
        """Rejects missing parent when create_parents=False."""
        output_path = tmp_path / "nested" / "deep" / "output.xlsx"

        with pytest.raises(SystemExit) as exc_info:
            validate_output_path(str(output_path), create_parents=False)

        assert exc_info.value.code == ExitCode.VALIDATION_ERROR

    def test_case_insensitive_extensions(self, tmp_path: Path) -> None:
        """Extensions are case-insensitive."""
        output_path = tmp_path / "output.XLSX"

        result = validate_output_path(str(output_path))

        assert result == output_path


class TestValidateOutputSuffix:
    """Tests for validate_output_suffix()."""

    def test_valid_extension(self, tmp_path: Path) -> None:
        """Accepts valid extension."""
        path = tmp_path / "test.xlsx"

        # Should not raise
        validate_output_suffix(path, {".xlsx"})

    def test_invalid_extension(self, tmp_path: Path) -> None:
        """Rejects invalid extension."""
        path = tmp_path / "test.txt"

        with pytest.raises(SystemExit) as exc_info:
            validate_output_suffix(path, {".xlsx", ".xlsm"})

        assert exc_info.value.code == ExitCode.VALIDATION_ERROR


class TestCheckMacroContract:
    """Tests for check_macro_contract()."""

    def test_xlsm_to_xlsx_loses_macros(self, tmp_path: Path) -> None:
        """Warns when xlsm → xlsx conversion strips macros."""
        input_path = tmp_path / "input.xlsm"
        output_path = tmp_path / "output.xlsx"

        warning = check_macro_contract(input_path, output_path)

        assert warning is not None
        assert "strip VBA macros" in warning
        assert ".xlsm" in warning

    def test_xlsm_to_xlsm_preserves_macros(self, tmp_path: Path) -> None:
        """No warning when xlsm → xlsm preserves macros."""
        input_path = tmp_path / "input.xlsm"
        output_path = tmp_path / "output.xlsm"

        warning = check_macro_contract(input_path, output_path)

        assert warning is None

    def test_xlsx_to_xlsx_no_macros(self, tmp_path: Path) -> None:
        """No warning when xlsx → xlsx (no macros to lose)."""
        input_path = tmp_path / "input.xlsx"
        output_path = tmp_path / "output.xlsx"

        warning = check_macro_contract(input_path, output_path)

        assert warning is None

    def test_xlsx_to_xlsm_no_warning(self, tmp_path: Path) -> None:
        """No warning when xlsx → xlsm (no macros to lose)."""
        input_path = tmp_path / "input.xlsx"
        output_path = tmp_path / "output.xlsm"

        warning = check_macro_contract(input_path, output_path)

        assert warning is None

    def test_xltm_to_xlsx_loses_macros(self, tmp_path: Path) -> None:
        """Warns when xltm → xlsx conversion strips macros."""
        input_path = tmp_path / "input.xltm"
        output_path = tmp_path / "output.xlsx"

        warning = check_macro_contract(input_path, output_path)

        assert warning is not None
        assert "strip VBA macros" in warning


class TestValidateInputPath:
    """Tests for validate_input_path()."""

    def test_valid_input_file(self, tmp_path: Path) -> None:
        """Accepts valid Excel file."""
        from openpyxl import Workbook

        path = tmp_path / "test.xlsx"
        wb = Workbook()
        wb.save(str(path))

        result = validate_input_path(str(path))

        assert result == path

    def test_rejects_missing_file(self, tmp_path: Path) -> None:
        """Rejects non-existent file."""
        path = tmp_path / "missing.xlsx"

        with pytest.raises(SystemExit) as exc_info:
            validate_input_path(str(path))

        assert exc_info.value.code == ExitCode.FILE_NOT_FOUND

    def test_rejects_directory(self, tmp_path: Path) -> None:
        """Rejects directory path."""
        with pytest.raises(SystemExit) as exc_info:
            validate_input_path(str(tmp_path))

        assert exc_info.value.code == ExitCode.FILE_NOT_FOUND

    def test_rejects_invalid_extension(self, tmp_path: Path) -> None:
        """Rejects non-Excel file."""
        path = tmp_path / "test.txt"
        path.write_text("not excel")

        with pytest.raises(SystemExit) as exc_info:
            validate_input_path(str(path))

        assert exc_info.value.code == ExitCode.VALIDATION_ERROR
