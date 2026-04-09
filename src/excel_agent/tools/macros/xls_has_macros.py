"""xls_has_macros: Boolean VBA presence check.

Quick check to determine if a workbook contains VBA macros.
This is useful for filtering and initial assessment.
"""

from __future__ import annotations

from excel_agent.core.macro_handler import has_macros
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict[str, object]:
    parser = create_parser("Check if a workbook contains VBA macros.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()

    path = validate_input_path(args.input)
    result = has_macros(path)

    return build_response(
        "success",
        {
            "has_macros": result,
            "file": str(path),
        },
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
