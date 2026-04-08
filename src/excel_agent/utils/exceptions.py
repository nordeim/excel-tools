"""
Custom exception hierarchy for excel-agent-tools.

Every exception maps to a specific exit code. Tool entry points catch
ExcelAgentError and convert to the appropriate JSON response + exit code.

Hierarchy:
    ExcelAgentError (base)
    ├── ExcelFileNotFoundError        → exit code 2
    ├── LockContentionError     → exit code 3
    ├── PermissionDeniedError   → exit code 4
    ├── ValidationError         → exit code 1
    ├── ImpactDeniedError       → exit code 1 (with prescriptive guidance)
    └── ConcurrentModificationError → exit code 5
"""

from __future__ import annotations

from typing import Any


class ExcelAgentError(Exception):
    """Base exception for all excel-agent errors.

    Attributes:
        exit_code: The standardized exit code for this error category.
        details: Additional context for debugging (included in JSON response).
    """

    exit_code: int = 5  # default: internal error

    def __init__(self, message: str, *, details: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.details = details or {}


class ExcelFileNotFoundError(ExcelAgentError):
    """Input file does not exist or is not readable."""

    exit_code: int = 2


class LockContentionError(ExcelAgentError):
    """File is locked by another process; timeout exceeded."""

    exit_code: int = 3


class PermissionDeniedError(ExcelAgentError):
    """Approval token is invalid, expired, revoked, or wrong scope.

    Attributes:
        reason: Specific reason for denial (for agent diagnostic).
    """

    exit_code: int = 4

    def __init__(
        self,
        message: str,
        *,
        reason: str = "",
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message, details=details)
        self.reason = reason


class ValidationError(ExcelAgentError):
    """Input validation failed — malformed data, bad range, schema violation."""

    exit_code: int = 1


class ImpactDeniedError(ExcelAgentError):
    """Destructive operation denied because it would break formula references.

    This is the denial-with-prescriptive-guidance pattern: the error includes
    the full impact report and a suggested next step for the agent.

    Attributes:
        impact_report: Structured impact analysis (broken refs, affected sheets).
        guidance: Prescriptive next action for the agent.
    """

    exit_code: int = 1

    def __init__(
        self,
        message: str,
        *,
        impact_report: dict[str, Any],
        guidance: str,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message, details=details)
        self.impact_report = impact_report
        self.guidance = guidance


class ConcurrentModificationError(ExcelAgentError):
    """File was modified by another process during our edit session.

    This means the file on disk changed between our __enter__ (load)
    and __exit__ (save). We refuse to save to prevent data loss.
    """

    exit_code: int = 5
