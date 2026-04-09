"""xls_version_hash: Compute geometry hash of workbook."""

from __future__ import annotations

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.version_hash import compute_file_hash
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Compute geometry and file hashes of a workbook.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()
    path = validate_input_path(args.input)

    file_hash = compute_file_hash(path)
    with ExcelAgent(path, mode="r") as agent:
        geometry_hash = agent.version_hash

    return build_response(
        "success",
        {"geometry_hash": geometry_hash, "file_hash": file_hash},
        workbook_version=geometry_hash,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
