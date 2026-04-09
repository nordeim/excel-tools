"""xls_get_defined_names: List all named ranges (global and sheet-scoped)."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("List all named ranges in a workbook.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()
    path = validate_input_path(args.input)

    with ExcelAgent(path, mode="r") as agent:
        wb = agent.workbook
        named_ranges: list[dict] = []

        for defn in wb.defined_names.definedName:
            scope = "Workbook"
            if defn.localSheetId is not None:
                idx = int(defn.localSheetId)
                if idx < len(wb.sheetnames):
                    scope = wb.sheetnames[idx]
                else:
                    scope = f"Sheet(index={idx})"

            named_ranges.append(
                {
                    "name": defn.name,
                    "scope": scope,
                    "refers_to": defn.attr_text or "",
                    "hidden": defn.hidden or False,
                    "is_reserved": defn.is_reserved or False,
                }
            )

        return build_response(
            "success",
            {"named_ranges": named_ranges, "count": len(named_ranges)},
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
