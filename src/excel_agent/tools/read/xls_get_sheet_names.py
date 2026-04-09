"""xls_get_sheet_names: List all sheets with index, name, visibility."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("List all sheets in a workbook with index, name, and visibility.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()
    path = validate_input_path(args.input)

    with ExcelAgent(path, mode="r") as agent:
        wb = agent.workbook
        sheets = []
        for idx, name in enumerate(wb.sheetnames):
            ws = wb[name]
            sheets.append(
                {
                    "index": idx,
                    "name": name,
                    "visibility": ws.sheet_state,
                }
            )

        return build_response(
            "success",
            {"sheets": sheets, "count": len(sheets)},
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
