"""xls_validate_workbook: OOXML compliance, broken refs, circular ref detection."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.dependency import DependencyTracker
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import add_common_args, create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Validate workbook: OOXML compliance, broken refs, circular refs.")
    add_common_args(parser)
    args = parser.parse_args()

    path = validate_input_path(args.input)

    with ExcelAgent(path, mode="r") as agent:
        wb = agent.workbook
        errors: list[str] = []
        warnings: list[str] = []

        # Check for error values in cells
        error_cells: list[dict] = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("#"):
                        if cell.value in (
                            "#REF!",
                            "#VALUE!",
                            "#DIV/0!",
                            "#NAME?",
                            "#N/A",
                            "#NUM!",
                            "#NULL!",
                        ):
                            error_cells.append(
                                {
                                    "sheet": sheet_name,
                                    "cell": f"{cell.column_letter}{cell.row}",
                                    "error": cell.value,
                                }
                            )
        if error_cells:
            errors.append(f"Found {len(error_cells)} error value(s) in cells")

        # Build dependency graph for circular ref detection
        tracker = DependencyTracker(wb)
        tracker.build_graph()
        circular = tracker.detect_circular_references()
        if circular:
            warnings.append(f"Found {len(circular)} circular reference chain(s)")

        valid = len(errors) == 0
        return build_response(
            "success" if valid else "warning",
            {
                "valid": valid,
                "errors": errors,
                "error_cells": error_cells[:20],
                "circular_refs": circular,
                "stats": tracker.get_stats(),
            },
            workbook_version=agent.version_hash,
            warnings=warnings,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
