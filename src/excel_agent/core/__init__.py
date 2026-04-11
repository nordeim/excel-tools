"""
Core components for excel-agent-tools.

This package contains the foundational building blocks:
- ExcelAgent: Context manager for safe, locked workbook manipulation
- EditSession: Unified edit target semantics, eliminates double-save
- FileLock: Cross-platform OS-level file locking
- RangeSerializer: Unified range parsing (A1/R1C1/Named/Table)
- DependencyTracker: Formula dependency graph for safety analysis
- version_hash: Geometry-aware workbook hashing
"""

from __future__ import annotations

__all__ = [
    "CellCoordinate",
    "DependencyTracker",
    "EditSession",
    "ExcelAgent",
    "FileLock",
    "ImpactReport",
    "prepare_edit_target",
    "RangeCoordinate",
    "RangeSerializer",
]


def __getattr__(name: str) -> object:
    if name == "ExcelAgent":
        from excel_agent.core.agent import ExcelAgent  # type: ignore[import-untyped]

        return ExcelAgent
    if name == "EditSession":
        from excel_agent.core.edit_session import EditSession

        return EditSession
    if name == "prepare_edit_target":
        from excel_agent.core.edit_session import prepare_edit_target

        return prepare_edit_target
    if name == "DependencyTracker":
        from excel_agent.core.dependency import DependencyTracker

        return DependencyTracker
    if name == "ImpactReport":
        from excel_agent.core.dependency import ImpactReport

        return ImpactReport
    if name == "FileLock":
        from excel_agent.core.locking import FileLock

        return FileLock
    if name in ("RangeSerializer", "CellCoordinate", "RangeCoordinate"):
        import excel_agent.core.serializers as serializers

        return getattr(serializers, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
