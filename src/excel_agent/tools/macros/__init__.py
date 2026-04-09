"""Macro Safety tools for excel-agent-tools.

This module provides 5 CLI tools for VBA macro manipulation:
- xls_has_macros: Boolean VBA presence check
- xls_inspect_macros: List VBA modules + signature status
- xls_validate_macro_safety: Risk scan: auto-exec, Shell, IOCs
- xls_remove_macros: Strip VBA (double-token required)
- xls_inject_vba_project: Inject pre-extracted .bin (token-gated)

These tools use the oletools library for VBA analysis, which can
detect, extract and analyze VBA macros, OLE objects, Excel 4
macros (XLM) and DDE links.
"""

from __future__ import annotations

__all__ = []
