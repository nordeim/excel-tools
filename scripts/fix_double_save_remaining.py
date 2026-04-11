#!/usr/bin/env python3
"""Fix double-save bug in the remaining 9 tools."""

from pathlib import Path
import re

TOOLS_DIR = Path("/home/project/excel-tools/src/excel_agent/tools")

# Only the 9 tools with actual double-save bug
FILES_TO_FIX = [
    "cells/xls_delete_range.py",
    "cells/xls_unmerge_cells.py",
    "cells/xls_update_references.py",
    "formulas/xls_set_formula.py",
    "structure/xls_delete_sheet.py",
    "structure/xls_insert_columns.py",
    "structure/xls_insert_rows.py",
    "structure/xls_move_sheet.py",
    "structure/xls_rename_sheet.py",
]


def fix_tool(filepath: Path) -> bool:
    """Fix a single tool file."""
    content = filepath.read_text()
    original = content

    changes = []

    # Change 1: Replace import
    if "from excel_agent.core.agent import ExcelAgent" in content:
        content = content.replace(
            "from excel_agent.core.agent import ExcelAgent",
            "from excel_agent.core.edit_session import EditSession",
        )
        changes.append("import")

    # Change 2: Replace context manager
    if 'with ExcelAgent(input_path, mode="rw") as agent:' in content:
        content = content.replace(
            'with ExcelAgent(input_path, mode="rw") as agent:',
            "session = EditSession.prepare(input_path, output_path)\n    with session:",
        )
        changes.append("context")

    # Change 3: Replace workbook access
    if "wb = agent.workbook" in content:
        content = content.replace("wb = agent.workbook", "wb = session.workbook")
        changes.append("workbook")

    # Change 4: Remove conditional save
    if "if str(output_path) != str(input_path):" in content:
        # Remove the conditional save block
        content = re.sub(
            r"\n        if str\(output_path\) != str\(input_path\):\s*\n            wb\.save\(str\(output_path\)\)",
            "\n        # EditSession handles save automatically on exit",
            content,
        )
        changes.append("save")

    # Change 5: Replace version_hash
    if "agent.version_hash" in content:
        content = content.replace("agent.version_hash", "session.version_hash")
        changes.append("version_hash")

    # Change 6: Replace file_hash in audit
    if "agent.file_hash" in content:
        content = content.replace("agent.file_hash", "session.file_hash")
        changes.append("file_hash")

    if content != original:
        filepath.write_text(content)
        print(f"  FIXED: {filepath.name} ({', '.join(changes)})")
        return True
    else:
        print(f"  SKIP: {filepath.name} - no changes needed")
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

    print(f"\nFixed {fixed}/{len(FILES_TO_FIX)} files")


if __name__ == "__main__":
    main()
