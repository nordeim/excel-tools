#!/usr/bin/env python3
"""Fix EditSession calls to allow in-place editing."""

import re
from pathlib import Path

TOOLS_DIR = Path("/home/project/excel-tools/src/excel_agent/tools")

# Pattern to find EditSession.prepare calls without force_inplace
# session = EditSession.prepare(input_path, output_path)
PATTERN = r"session = EditSession\.prepare\(input_path, output_path\)"
REPLACEMENT = """output_arg = args.output or args.input
    output_path = validate_output_path(output_arg, create_parents=True)

    # Use EditSession for proper copy-on-write and save semantics
    # Allow in-place editing when output == input (default behavior)
    force_inplace = str(output_path) == str(input_path)
    session = EditSession.prepare(input_path, output_path, force_inplace=force_inplace)"""


def fix_tool(filepath: Path) -> bool:
    """Fix a single tool file."""
    content = filepath.read_text()
    original = content

    # Find and replace the pattern
    if re.search(PATTERN, content):
        # Remove the old output_path line if it exists
        content = re.sub(
            r"output_path = validate_output_path\(args\.output or args\.input, create_parents=True\)\s*\n\s*",
            "",
            content,
        )

        # Replace the EditSession.prepare line
        content = re.sub(
            PATTERN,
            """output_arg = args.output or args.input
    output_path = validate_output_path(output_arg, create_parents=True)

    # Use EditSession for proper copy-on-write and save semantics
    # Allow in-place editing when output == input (default behavior)
    force_inplace = str(output_path) == str(input_path)
    session = EditSession.prepare(input_path, output_path, force_inplace=force_inplace)""",
            content,
        )

        if content != original:
            filepath.write_text(content)
            print(f"  FIXED: {filepath.name}")
            return True

    return False


def main():
    """Fix all tools using EditSession."""
    fixed = 0

    for filepath in TOOLS_DIR.rglob("*.py"):
        if "EditSession" in filepath.read_text():
            if fix_tool(filepath):
                fixed += 1

    print(f"\nFixed {fixed} files")


if __name__ == "__main__":
    main()
