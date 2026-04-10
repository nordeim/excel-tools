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

    try:
        with ExcelAgent(path, mode="r") as agent:
            wb = agent.workbook
            named_ranges: list[dict] = []

            # Safe access to defined names - handle workbooks with no named ranges
            if wb.defined_names is None:
                return build_response(
                    "success",
                    {"named_ranges": [], "count": 0},
                    workbook_version=agent.version_hash,
                )

            # Handle different openpyxl API styles
            defined_names_list = getattr(wb.defined_names, "definedName", None)
            if defined_names_list is None:
                # Try alternative access pattern
                try:
                    defined_names_list = list(wb.defined_names.values())
                except (AttributeError, TypeError):
                    defined_names_list = []

            for defn in defined_names_list:
                scope = "Workbook"
                local_sheet_id = getattr(defn, "localSheetId", None)

                if local_sheet_id is not None:
                    try:
                        idx = int(local_sheet_id)
                        if idx < len(wb.sheetnames):
                            scope = wb.sheetnames[idx]
                        else:
                            scope = f"Sheet(index={idx})"
                    except (ValueError, TypeError):
                        scope = "Unknown"

                named_ranges.append(
                    {
                        "name": getattr(defn, "name", "Unknown"),
                        "scope": scope,
                        "refers_to": getattr(defn, "attr_text", "") or "",
                        "hidden": getattr(defn, "hidden", False) or False,
                        "is_reserved": getattr(defn, "is_reserved", False) or False,
                    }
                )

            return build_response(
                "success",
                {"named_ranges": named_ranges, "count": len(named_ranges)},
                workbook_version=agent.version_hash,
            )
    except Exception as e:
        return build_response(
            "error",
            None,
            exit_code=5,
            error=f"Failed to read named ranges: {str(e)}",
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
