"""xls_clone_workbook: Atomic copy of workbook to a work directory."""

from __future__ import annotations

import shutil
from datetime import datetime, timezone
from pathlib import Path

from excel_agent.core.version_hash import compute_file_hash
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Atomic copy of workbook to a safe working directory.")
    parser.add_argument("--input", type=str, required=True, help="Source workbook path")
    parser.add_argument(
        "--output-dir", type=str, default="./work", help="Target directory (default: ./work)"
    )
    args = parser.parse_args()

    source = validate_input_path(args.input)
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    source_hash = compute_file_hash(source)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S")
    short_hash = source_hash.replace("sha256:", "")[:8]
    clone_name = f"{source.stem}_{timestamp}_{short_hash}{source.suffix}"
    clone_path = output_dir / clone_name

    shutil.copy2(source, clone_path)
    clone_hash = compute_file_hash(clone_path)

    return build_response(
        "success",
        {
            "clone_path": str(clone_path),
            "clone_name": clone_name,
            "source_path": str(source),
            "source_hash": source_hash,
            "clone_hash": clone_hash,
        },
        workbook_version=source_hash,
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
