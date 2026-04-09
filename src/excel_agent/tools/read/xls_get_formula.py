"""xls_get_formula: Get formula from a specific cell with parsed references."""

from __future__ import annotations

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _extract_references(formula: str) -> list[str]:
    """Extract cell references from a formula string."""
    try:
        tok = Tokenizer(formula)
    except Exception:
        return []
    return [t.value for t in tok.items if t.type == Token.OPERAND and t.subtype == Token.RANGE]


def _run() -> dict:
    parser = create_parser("Get the formula from a specific cell.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    parser.add_argument("--cell", type=str, required=True, help="Cell reference (e.g., A1)")
    parser.add_argument("--sheet", type=str, default=None, help="Sheet name")
    args = parser.parse_args()
    path = validate_input_path(args.input)

    with ExcelAgent(path, mode="r") as agent:
        wb = agent.workbook
        sheet_name = args.sheet or wb.sheetnames[0]
        ws = wb[sheet_name]
        cell = ws[args.cell]

        formula: str | None = None
        references: list[str] = []

        if cell.data_type == "f":
            formula = str(cell.value)
            references = _extract_references(formula)

        return build_response(
            "success",
            {
                "cell": args.cell,
                "sheet": sheet_name,
                "formula": formula,
                "references": references,
                "has_formula": formula is not None,
            },
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
