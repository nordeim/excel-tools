"""
Core components for excel-agent-tools.

This package contains the foundational building blocks:
    - ExcelAgent: Context manager for safe, locked workbook manipulation
    - FileLock: Cross-platform OS-level file locking
    - RangeSerializer: Unified range parsing (A1/R1C1/Named/Table)
    - version_hash: Geometry-aware workbook hashing
"""

from __future__ import annotations

__all__ = [
    "CellCoordinate",
    "ExcelAgent",
    "FileLock",
    "RangeCoordinate",
    "RangeSerializer",
]


def __getattr__(name: str) -> object:
    if name == "ExcelAgent":
        from excel_agent.core.agent import ExcelAgent  # type: ignore[import-untyped]

        return ExcelAgent
    if name == "FileLock":
        from excel_agent.core.locking import FileLock

        return FileLock
    if name in ("RangeSerializer", "CellCoordinate", "RangeCoordinate"):
        import excel_agent.core.serializers as serializers

        return getattr(serializers, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
