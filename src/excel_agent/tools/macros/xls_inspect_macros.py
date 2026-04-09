"""xls_inspect_macros: List VBA modules and check signature status.

Detailed inspection of VBA macros including:
- Module names and source code (first 500 chars)
- Digital signature status
- Module count
- Stream information
"""

from __future__ import annotations

from excel_agent.core.macro_handler import get_analyzer
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict[str, object]:
    parser = create_parser("Inspect VBA macros: list modules, view code, check signature.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    parser.add_argument(
        "--code-preview-length",
        type=int,
        default=500,
        help="Characters of code to preview (default: 500)",
    )
    args = parser.parse_args()

    path = validate_input_path(args.input)
    analyzer = get_analyzer()
    result = analyzer.analyze(path)

    # Adjust code preview length
    if args.code_preview_length >= 0:
        for module_dict in result.to_dict().get("modules", []):
            full_code = module_dict.get("code_preview", "")
            module_dict["code_preview"] = full_code[: args.code_preview_length]

    return build_response(
        "success" if not result.errors else "warning",
        result.to_dict(),
        warnings=result.errors if result.errors else None,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
