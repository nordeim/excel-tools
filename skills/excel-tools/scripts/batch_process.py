#!/usr/bin/env python3
"""
Batch process multiple Excel files.
Part of excel-tools-skill.
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path


def run_tool(tool: str, **kwargs) -> dict:
    """Run an excel-agent tool."""
    cmd = [f"xls-{tool}"]
    for key, value in kwargs.items():
        cmd.extend([f"--{key.replace('_', '-')}", str(value)])

    result = subprocess.run(cmd, capture_output=True, text=True)

    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError:
        return {"status": "error", "error": result.stderr or result.stdout}


def process_file(input_path: str, output_dir: str, operations: list) -> dict:
    """Process a single file with multiple operations.

    Args:
        input_path: Input file path
        output_dir: Output directory
        operations: List of operation dicts

    Returns:
        Result dict
    """
    from pathlib import Path

    input_file = Path(input_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Clone first
    clone_result = run_tool("clone-workbook", input=input_path, output_dir=str(output_dir))

    if clone_result.get("status") != "success":
        return {"status": "error", "file": input_path, "error": "Clone failed"}

    clone_path = clone_result["data"]["clone_path"]

    # Apply operations
    for op in operations:
        tool = op["tool"]
        params = op.get("params", {})
        params["input"] = clone_path
        params["output"] = clone_path

        result = run_tool(tool, **params)
        if result.get("status") != "success":
            return {"status": "error", "file": input_path, "operation": tool}

    return {"status": "success", "file": input_path, "output": clone_path}


def main():
    parser = argparse.ArgumentParser(description="Batch process Excel files")
    parser.add_argument("--input", required=True, help="Input directory or file pattern")
    parser.add_argument("--output", required=True, help="Output directory")
    parser.add_argument(
        "--operations", required=True, help='JSON operations: [{"tool": "...", "params": {...}}]'
    )

    args = parser.parse_args()

    operations = json.loads(args.operations)

    # Get input files
    input_path = Path(args.input)
    if input_path.is_dir():
        files = list(input_path.glob("*.xlsx"))
    else:
        # Treat as glob pattern
        files = list(Path(".").glob(args.input))

    # Process each file
    results = []
    for file in files:
        result = process_file(str(file), args.output, operations)
        results.append(result)
        print(json.dumps(result))

    # Summary
    success = sum(1 for r in results if r["status"] == "success")
    failed = len(results) - success

    print(f"\nProcessed: {len(results)}, Success: {success}, Failed: {failed}")

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
