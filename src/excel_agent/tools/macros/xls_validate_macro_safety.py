"""xls_validate_macro_safety: Risk scan for VBA macros.

Analyzes VBA macros for suspicious patterns:
- Auto-execute functions (AutoOpen, Document_Open, etc.)
- Shell execution (Shell, CreateObject, etc.)
- Network activity (WinHttp, URLDownloadToFile, etc.)
- Obfuscation techniques

Provides a risk score and level (none/low/medium/high).
"""

from __future__ import annotations

from excel_agent.core.macro_handler import get_analyzer
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict[str, object]:
    parser = create_parser(
        "Validate VBA macro safety: scan for auto-exec, shell, network, obfuscation."
    )
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()

    path = validate_input_path(args.input)
    analyzer = get_analyzer()
    result = analyzer.analyze(path)

    status = "success"
    if result.risk_level == "high":
        status = "warning"
    elif result.risk_level == "medium":
        status = "warning"

    return build_response(
        status,
        result.to_dict(),
        warnings=(
            [
                f"Risk level: {result.risk_level}",
                f"Risk score: {result.risk_score}/100",
            ]
            if result.risk_level != "none"
            else None
        ),
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
