"""xls_get_cell_style: Get formatting for a cell as JSON."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.style_serializer import serialize_cell_style
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Get the full formatting/style of a cell as JSON.")
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

        style_data = serialize_cell_style(cell)

        return build_response(
            "success",
            {"cell": args.cell, "sheet": sheet_name, **style_data},
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
