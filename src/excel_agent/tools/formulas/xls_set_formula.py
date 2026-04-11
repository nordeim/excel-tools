"""xls_set_formula: Set formula in a cell with syntax validation."""

from __future__ import annotations

from openpyxl.formula import Tokenizer

from excel_agent.core.edit_session import EditSession
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.exceptions import ValidationError
from excel_agent.utils.json_io import build_response


def _validate_formula_syntax(formula: str) -> list[str]:
    """Validate formula syntax using the openpyxl Tokenizer.

    Returns a list of warning strings (empty if valid).
    """
    warnings: list[str] = []
    if not formula.startswith("="):
        warnings.append("Formula must start with '='")
        return warnings
    try:
        tok = Tokenizer(formula)
        tokens = tok.items
        if not tokens:
            warnings.append("Formula parsed to zero tokens")
        # Check for unclosed parentheses
        open_count = sum(1 for t in tokens if t.value == "(" or t.subtype == "OPEN")
        close_count = sum(1 for t in tokens if t.value == ")" or t.subtype == "CLOSE")
        if open_count != close_count:
            warnings.append(f"Mismatched parentheses: {open_count} open, {close_count} close")
    except Exception as exc:
        warnings.append(f"Formula syntax error: {exc}")
    return warnings


def _run() -> dict[str, object]:
    parser = create_parser("Set a formula in a cell with syntax validation.")
    add_common_args(parser)
    parser.add_argument("--cell", type=str, required=True, help="Target cell (e.g., A1)")
    parser.add_argument(
        "--formula",
        type=str,
        required=True,
        help="Formula string (e.g., =SUM(B1:B10))",
    )
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(args.output or args.input, create_parents=True)

    formula = args.formula
    if not formula.startswith("="):
        formula = f"={formula}"

    # Validate syntax
    syntax_warnings = _validate_formula_syntax(formula)
    if any("error" in w.lower() for w in syntax_warnings):
        raise ValidationError(
            f"Formula syntax validation failed: {'; '.join(syntax_warnings)}",
            details={"formula": formula, "warnings": syntax_warnings},
        )

    session = EditSession.prepare(input_path, output_path)
    with session:
        wb = session.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]
        ws[args.cell] = formula

        # Capture version hash before exiting context
        version_hash = session.version_hash

        # EditSession handles save automatically on exit

    return build_response(
        "success",
        {
            "cell": args.cell,
            "sheet": sheet_name,
            "formula": formula,
            "syntax_warnings": syntax_warnings,
        },
        workbook_version=version_hash,
        impact={"cells_modified": 1, "formulas_updated": 1},
        warnings=syntax_warnings if syntax_warnings else None,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
