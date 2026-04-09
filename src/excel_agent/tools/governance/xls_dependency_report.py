"""xls_dependency_report: Full dependency graph export as JSON adjacency list."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.dependency import DependencyTracker
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import add_common_args, create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Export the full formula dependency graph as a JSON adjacency list.")
    add_common_args(parser)
    args = parser.parse_args()
    path = validate_input_path(args.input)

    with ExcelAgent(path, mode="r") as agent:
        tracker = DependencyTracker(agent.workbook)
        sheets_filter = [args.sheet] if args.sheet else None
        tracker.build_graph(sheets=sheets_filter)

        return build_response(
            "success",
            {
                "stats": tracker.get_stats(),
                "graph": tracker.get_adjacency_list(),
                "circular_refs": tracker.detect_circular_references(),
            },
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
