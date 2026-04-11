"""
Reusable CLI argument parsing, path validation, and JSON input handling.

Every tool uses these helpers for consistent argument names, validation
behavior, and error reporting. This ensures the AI agent sees a uniform
interface across all 53 tools.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Set

from excel_agent.utils.exit_codes import ExitCode, exit_with


# Valid Excel file extensions
_VALID_EXCEL_EXTENSIONS: Set[str] = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})

# Extensions that support VBA macros
_VBA_EXTENSIONS: Set[str] = frozenset({".xlsm", ".xltm"})


def create_parser(description: str) -> argparse.ArgumentParser:
    """Create an ArgumentParser with the standard excel-agent-tools format.

    Args:
        description: Tool-specific help text shown in --help.

    Returns:
        A configured ArgumentParser ready for add_common_args / add_governance_args.
    """
    return argparse.ArgumentParser(
        description=description,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )


def add_common_args(parser: argparse.ArgumentParser) -> None:
    """Add standard arguments shared by most tools.

    Adds:
        --input   : Path to input workbook (required for most tools)
        --output  : Path to output workbook (optional, defaults to in-place)
        --sheet   : Target sheet name (optional, defaults to active sheet)
        --format  : Output format — json or jsonl (default: json)
    """
    parser.add_argument(
        "--input",
        type=str,
        required=True,
        help="Path to the input Excel workbook (.xlsx or .xlsm)",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help=(
            "Path to the output workbook (default: overwrite input — requires --force for safety)"
        ),
    )
    parser.add_argument(
        "--sheet",
        type=str,
        default=None,
        help="Target sheet name (default: active sheet)",
    )
    parser.add_argument(
        "--format",
        type=str,
        choices=["json", "jsonl"],
        default="json",
        help="Output format: json (default) or jsonl (streaming, one object per line)",
    )


def add_governance_args(parser: argparse.ArgumentParser) -> None:
    """Add governance-related arguments for destructive operations.

    Adds:
        --token              : HMAC-SHA256 approval token string
        --force              : Skip confirmation prompts (still requires token for gated ops)
        --acknowledge-impact : Acknowledge pre-flight impact report and proceed
    """
    parser.add_argument(
        "--token",
        type=str,
        default=None,
        help="HMAC-SHA256 approval token for governance-gated operations",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        default=False,
        help="Force operation even if impact report shows warnings",
    )
    parser.add_argument(
        "--acknowledge-impact",
        action="store_true",
        default=False,
        help="Acknowledge that this operation may break formula references",
    )


def validate_input_path(path_str: str) -> Path:
    """Validate that the input file exists and is readable.

    Args:
        path_str: Raw string path from CLI argument.

    Returns:
        Resolved Path object.

    Raises:
        SystemExit: With exit code 2 if file not found or not readable.
    """
    path = Path(path_str).resolve()

    if not path.exists():
        exit_with(
            ExitCode.FILE_NOT_FOUND,
            f"Input file not found: {path}",
            details={"path": str(path)},
        )

    if not path.is_file():
        exit_with(
            ExitCode.FILE_NOT_FOUND,
            f"Input path is not a file: {path}",
            details={"path": str(path)},
        )

    if path.suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Unsupported file format: {path.suffix}. Expected .xlsx, .xlsm, .xltx, or .xltm",
            details={"path": str(path), "suffix": path.suffix},
        )

    return path


def validate_output_suffix(path: Path, allowed: Set[str] | None = None) -> None:
    """Validate output file extension.

    Args:
        path: Output path to validate.
        allowed: Set of allowed extensions (default: _VALID_EXCEL_EXTENSIONS).

    Raises:
        SystemExit: With exit code 1 if extension not in allowed set.
    """
    allowed = allowed or _VALID_EXCEL_EXTENSIONS
    ext = path.suffix.lower()

    if ext not in allowed:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Invalid output extension: {ext}",
            details={
                "path": str(path),
                "extension": ext,
                "allowed_extensions": sorted(allowed),
            },
        )


def check_macro_contract(input_path: Path, output_path: Path) -> str | None:
    """Check if macros will be lost during save.

    Args:
        input_path: Source file path.
        output_path: Target file path.

    Returns:
        Warning message if macros will be lost, None otherwise.
    """
    input_ext = input_path.suffix.lower()
    output_ext = output_path.suffix.lower()

    # Input has macros, output doesn't support them
    if input_ext in _VBA_EXTENSIONS and output_ext not in _VBA_EXTENSIONS:
        return (
            f"Converting {input_ext} to {output_ext} will strip VBA macros. "
            f"Use {output_path.stem}.xlsm extension to preserve macros."
        )

    return None


def validate_output_path(
    path_str: str,
    *,
    create_parents: bool = False,
    allowed_suffixes: Set[str] | None = None,
    overwrite: bool = True,  # Default True for backward compatibility
) -> Path:
    """Validate output path with comprehensive checks.

    Enhanced version that validates:
    - Parent directory exists (or creates if create_parents)
    - Extension is in allowed set
    - File doesn't exist (unless overwrite=True)

    Args:
        path_str: Raw string path from CLI argument.
        create_parents: If True, create parent directories as needed.
        allowed_suffixes: Set of allowed extensions (default: _VALID_EXCEL_EXTENSIONS).
        overwrite: If True, allow overwriting existing files.

    Returns:
        Resolved Path object.

    Raises:
        SystemExit: With exit code 1 if validation fails.
    """
    path = Path(path_str).resolve()

    # Check/create parent directory
    if create_parents:
        path.parent.mkdir(parents=True, exist_ok=True)
    elif not path.parent.exists():
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Output directory does not exist: {path.parent}",
            details={"path": str(path), "parent": str(path.parent)},
        )

    # Validate extension
    if allowed_suffixes is not None:
        validate_output_suffix(path, allowed_suffixes)
    else:
        validate_output_suffix(path, _VALID_EXCEL_EXTENSIONS)

    # Check overwrite
    if path.exists() and not overwrite:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Output file exists: {path}. Use --overwrite to replace.",
            details={"path": str(path)},
        )

    return path


def validate_output_path_legacy(path_str: str, *, create_parents: bool = False) -> Path:
    """Original validate_output_path for backward compatibility.

    This is the legacy version that only checks parent directory.
    Use validate_output_path() for new code.

    Args:
        path_str: Raw string path from CLI argument.
        create_parents: If True, create parent directories as needed.

    Returns:
        Resolved Path object.

    Raises:
        SystemExit: With exit code 1 if parent directory doesn't exist.
    """
    path = Path(path_str).resolve()

    if create_parents:
        path.parent.mkdir(parents=True, exist_ok=True)
    elif not path.parent.exists():
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Output directory does not exist: {path.parent}",
            details={"path": str(path), "parent": str(path.parent)},
        )

    return path


def load_json_stdin() -> dict[str, Any]:
    """Read and parse JSON from stdin.

    Used by tools that accept complex input (data arrays, style specs)
    piped from the AI agent orchestrator.

    Returns:
        Parsed dict.

    Raises:
        SystemExit: With exit code 1 if stdin is empty or contains invalid JSON.
    """
    try:
        raw = sys.stdin.read()
    except Exception as exc:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Failed to read from stdin: {exc}",
        )

    if not raw.strip():
        exit_with(
            ExitCode.VALIDATION_ERROR,
            "No JSON data provided on stdin",
        )

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Invalid JSON on stdin: {exc}",
            details={"line": exc.lineno, "column": exc.colno},
        )

    if not isinstance(data, dict):
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Expected JSON object on stdin, got {type(data).__name__}",
        )

    return data


def parse_json_arg(value: str) -> Any:
    """Parse a JSON string from a CLI argument (e.g., --data '[...]').

    Args:
        value: Raw JSON string from CLI.

    Returns:
        Parsed Python object.

    Raises:
        SystemExit: With exit code 1 if invalid JSON.
    """
    try:
        return json.loads(value)
    except json.JSONDecodeError as exc:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Invalid JSON in argument: {exc}",
            details={"raw_value": value[:200]},  # truncate for safety
        )
