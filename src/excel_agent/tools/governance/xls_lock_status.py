"""xls_lock_status: Check OS-level file lock state."""

from __future__ import annotations

from excel_agent.core.locking import FileLock
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Check if a workbook is currently locked by another process.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()
    path = validate_input_path(args.input)

    locked = FileLock.is_locked(path)
    lock_path = path.parent / f".{path.name}.lock"

    return build_response(
        "success",
        {
            "locked": locked,
            "lock_file_exists": lock_path.exists(),
            "lock_file_path": str(lock_path),
        },
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
