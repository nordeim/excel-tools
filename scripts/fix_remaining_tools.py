#!/usr/bin/env python3
"""Fix double-save bug in remaining Excel tools.

This script performs surgical edits to convert from ExcelAgent to EditSession pattern.
"""

import re
from pathlib import Path

TOOLS_DIR = Path("/home/project/excel-tools/src/excel_agent/tools")

# Files that still need fixing (have been verified to have the bug)
FILES_TO_FIX = [
    "structure/xls_delete_columns.py",
    "structure/xls_insert_rows.py",
    "structure/xls_insert_columns.py",
    "structure/xls_rename_sheet.py",
    "structure/xls_delete_sheet.py",
    "structure/xls_move_sheet.py",
    "cells/xls_unmerge_cells.py",
    "cells/xls_delete_range.py",
    "cells/xls_update_references.py",
    "formulas/xls_set_formula.py",
]


def fix_tool(filepath: Path) -> bool:
    """Fix a single tool file."""
    content = filepath.read_text()
    original = content

    # Skip if already has EditSession
    if "from excel_agent.core.edit_session import EditSession" in content:
        print(f"  SKIP: {filepath.name} - already uses EditSession")
        return False

    # Skip if no ExcelAgent import
    if "from excel_agent.core.agent import ExcelAgent" not in content:
        print(f"  SKIP: {filepath.name} - no ExcelAgent import found")
        return False

    # Pattern 1: Replace import
    content = content.replace(
        "from excel_agent.core.agent import ExcelAgent",
        "from excel_agent.core.edit_session import EditSession",
    )

    # Pattern 2: Replace context manager and add session prep
    # This pattern handles the standard case
    content = re.sub(
        r'with ExcelAgent\(input_path, mode="rw"\) as agent:',
        "session = EditSession.prepare(input_path, output_path)\n    with session:",
        content,
    )

    # Pattern 3: Replace workbook access
    content = content.replace("wb = agent.workbook", "wb = session.workbook")

    # Pattern 4: Replace conditional save block
    content = re.sub(
        r"\n        if str\(output_path\) != str\(input_path\):\s*\n            wb\.save\(str\(output_path\)\)",
        "\n        # EditSession handles save automatically on exit",
        content,
    )

    # Pattern 5: Replace agent.version_hash
    content = content.replace("agent.version_hash", "session.version_hash")

    # Pattern 6: Replace agent.file_hash
    content = content.replace("agent.file_hash", "session.file_hash")

    # Pattern 7: Replace file_hash = compute_file_hash with session.file_hash
    content = re.sub(
        r"file_hash = compute_file_hash\(input_path\)",
        "# file_hash obtained from session",
        content,
    )

    if content != original:
        filepath.write_text(content)
        print(f"  FIXED: {filepath.name}")
        return True
    else:
        print(f"  NO CHANGE: {filepath.name}")
        return False


def main():
    """Fix all tools."""
    fixed = 0
    for rel_path in FILES_TO_FIX:
        filepath = TOOLS_DIR / rel_path
        if filepath.exists():
            if fix_tool(filepath):
                fixed += 1
        else:
            print(f"  NOT FOUND: {rel_path}")

    print(f"\nFixed {fixed} files")


if __name__ == "__main__":
    main()
