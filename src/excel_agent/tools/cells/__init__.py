"""Cell Operations tools for excel-agent-tools.

This module provides 4 CLI tools for manipulating cells:
- xls_merge_cells: Merge a range of cells with hidden data pre-check
- xls_unmerge_cells: Restore grid from merged range
- xls_delete_range: Clear a range and shift cells up or left
- xls_update_references: Batch-update cell references in formulas

These tools use openpyxl's merge_cells/unmerge_cells and move_range APIs,
with careful attention to:
1. Data loss prevention (merge warns about non-anchor cell data)
2. Mutation-safe iteration (unmerge collects ranges before iterating)
3. Formula reference updating (via formula_updater module)
"""

from __future__ import annotations

__all__ = []
