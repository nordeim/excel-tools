#!/usr/bin/env python3
"""Fix double-save bug in all Excel tools.

Converts tools from:
- Using raw ExcelAgent with conditional wb.save()

To:
- Using EditSession which handles copy-on-write and save automatically
"""

from pathlib import Path
import re

# Files that need fixing (have double-save pattern)
TOOLS_TO_FIX = [
    "src/excel_agent/tools/structure/xls_add_sheet.py",
    "src/excel_agent/tools/structure/xls_delete_rows.py",
    "src/excel_agent/tools/structure/xls_delete_columns.py",
    "src/excel_agent/tools/structure/xls_insert_rows.py",
    "src/excel_agent/tools/structure/xls_insert_columns.py",
    "src/excel_agent/tools/structure/xls_rename_sheet.py",
    "src/excel_agent/tools/structure/xls_delete_sheet.py",
    "src/excel_agent/tools/structure/xls_move_sheet.py",
    "src/excel_agent/tools/cells/xls_merge_cells.py",
    "src/excel_agent/tools/cells/xls_unmerge_cells.py",
    "src/excel_agent/tools/cells/xls_delete_range.py",
    "src/excel_agent/tools/cells/xls_update_references.py",
    "src/excel_agent/tools/formulas/xls_set_formula.py",
]


def fix_file(filepath: Path) -> bool:
    """Fix a single file. Returns True if modified."""
    content = filepath.read_text()
    original = content

    # Skip if already using EditSession
    if "EditSession" in content:
        print(f"  Skipping {filepath.name} - already uses EditSession")
        return False

    # Skip if doesn't have ExcelAgent
    if "ExcelAgent" not in content:
        print(f"  Skipping {filepath.name} - no ExcelAgent")
        return False

    # 1. Replace import
    content = content.replace(
        "from excel_agent.core.agent import ExcelAgent",
        "from excel_agent.core.edit_session import EditSession",
    )

    # 2. Replace context manager pattern
    # Pattern: with ExcelAgent(input_path, mode="rw") as agent:
    content = re.sub(
        r'with ExcelAgent\(input_path, mode="rw"\) as agent:',
        "session = EditSession.prepare(input_path, output_path)\n    with session:",
        content,
    )

    # 3. Replace wb = agent.workbook
    content = content.replace("wb = agent.workbook", "wb = session.workbook")

    # 4. Remove the conditional save block
    # Pattern: if str(output_path) != str(input_path):\n    #             wb.save(str(output_path))
    content = re.sub(
        r"\n        if str\(output_path\) != str\(input_path\):\s*\n            wb\.save\(str\(output_path\)\)",
        "",
        content,
    )

    # Also try without the newline prefix
    content = re.sub(
        r"if str\(output_path\) != str\(input_path\):\s*\n            wb\.save\(str\(output_path\)\)",
        "",
        content,
    )

    # 5. Replace agent.version_hash with session.version_hash
    content = content.replace("agent.version_hash", "session.version_hash")

    # 6. Replace agent.file_hash with session.file_hash
    content = content.replace("agent.file_hash", "session.file_hash")

    if content != original:
        filepath.write_text(content)
        print(f"  Fixed {filepath.name}")
        return True
    else:
        print(f"  No changes needed for {filepath.name}")
        return False


def main():
    """Fix all files."""
    base_path = Path("/home/project/excel-tools")

    fixed_count = 0
    for tool_path in TOOLS_TO_FIX:
        filepath = base_path / tool_path
        if filepath.exists():
            if fix_file(filepath):
                fixed_count += 1
        else:
            print(f"  File not found: {tool_path}")

    print(f"\nFixed {fixed_count} files")


if __name__ == "__main__":
    main()
