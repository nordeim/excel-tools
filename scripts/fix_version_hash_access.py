#!/usr/bin/env python3
"""Fix session.version_hash accessed outside context manager.

This script fixes tools that access session.version_hash outside the with block.
"""

import re
from pathlib import Path

TOOLS_DIR = Path("/home/project/excel-tools/src/excel_agent/tools")


def fix_file(filepath: Path) -> bool:
    """Fix a single file."""
    content = filepath.read_text()
    original = content

    # Find files with workbook_version=session.version_hash outside context
    # Pattern: return build_response(... workbook_version=session.version_hash ...)
    # This needs to be captured inside the with block first

    if "workbook_version=session.version_hash" not in content:
        return False

    # Pattern to find the build_response call outside the with block
    # We need to add version_hash = session.version_hash before exiting
    # and change workbook_version=session.version_hash to workbook_version=version_hash

    # First, let's add the capture before the context exits
    # Look for "# EditSession handles save automatically on exit" comment
    if "# EditSession handles save automatically on exit" in content:
        # Add version hash capture before the comment
        content = content.replace(
            "        # EditSession handles save automatically on exit",
            "        # Capture version hash before exiting context\n        version_hash = session.version_hash\n\n        # EditSession handles save automatically on exit",
        )

        # Replace workbook_version=session.version_hash with workbook_version=version_hash
        content = content.replace(
            "workbook_version=session.version_hash", "workbook_version=version_hash"
        )

        if content != original:
            filepath.write_text(content)
            print(f"  FIXED: {filepath.name}")
            return True

    return False


def main():
    """Fix all files."""
    fixed = 0

    for filepath in TOOLS_DIR.rglob("*.py"):
        if filepath.is_file():
            if fix_file(filepath):
                fixed += 1

    print(f"\nFixed {fixed} files")


if __name__ == "__main__":
    main()
