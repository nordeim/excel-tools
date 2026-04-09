"""xls_get_table_info: List Excel Tables (ListObjects) with schema."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("List all Excel Tables (ListObjects) in a workbook.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()
    path = validate_input_path(args.input)

    with ExcelAgent(path, mode="r") as agent:
        wb = agent.workbook
        tables: list[dict] = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for table_name, table in ws.tables.items():
                columns = [col.name for col in table.tableColumns]
                tables.append(
                    {
                        "name": table.name,
                        "sheet": sheet_name,
                        "range": table.ref,
                        "columns": columns,
                        "has_totals_row": table.totalsRowShown or False,
                        "style": table.tableStyleInfo.name if table.tableStyleInfo else None,
                    }
                )

        return build_response(
            "success",
            {"tables": tables, "count": len(tables)},
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
