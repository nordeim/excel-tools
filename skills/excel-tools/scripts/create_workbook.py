#!/usr/bin/env python3
"""
Helper script to create Excel workbooks programmatically.
Part of excel-tools-skill.
"""

import argparse
import json
import sys
from pathlib import Path


def create_workbook(
    output: str, sheets: list[str] | None = None, data: dict | None = None
) -> dict:
    """Create a workbook with optional data.

    Args:
        output: Output file path
        sheets: List of sheet names
        data: Dict of sheet_name -> 2D array data

    Returns:
        Result dict with status and path
    """
    import subprocess

    # Create workbook
    sheets_arg = ",".join(sheets) if sheets else "Sheet1"
    result = subprocess.run(
        ["xls-create-new", "--output", output, "--sheets", sheets_arg],
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        # excel-agent-tools writes JSON errors to stdout, not stderr
        try:
            error_data = json.loads(result.stdout)
            return error_data
        except json.JSONDecodeError:
            return {"status": "error", "error": result.stdout or result.stderr}

    # Write data if provided
    if data:
        for sheet, rows in data.items():
            data_json = json.dumps(rows)
            write_result = subprocess.run(
                [
                    "xls-write-range",
                    "--input",
                    output,
                    "--output",
                    output,
                    "--sheet",
                    sheet,
                    "--range",
                    "A1",
                    "--data",
                    data_json,
                ],
                capture_output=True,
                text=True,
            )
            if write_result.returncode != 0:
                return {"status": "error", "error": f"Failed to write to {sheet}"}

    return {"status": "success", "path": output}


def main():
    parser = argparse.ArgumentParser(description="Create Excel workbook")
    parser.add_argument("--output", required=True, help="Output file path")
    parser.add_argument("--sheets", help="Comma-separated sheet names")
    parser.add_argument("--data", help='JSON data: {"Sheet1": [[...]], ...}')

    args = parser.parse_args()

    sheets = args.sheets.split(",") if args.sheets else None
    data = json.loads(args.data) if args.data else None

    result = create_workbook(args.output, sheets, data)
    print(json.dumps(result, indent=2))

    return 0 if result["status"] == "success" else 1


if __name__ == "__main__":
    sys.exit(main())
