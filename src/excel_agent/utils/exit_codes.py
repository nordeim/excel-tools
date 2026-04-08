"""
Standardized exit codes for all 53 excel-agent-tools CLI tools.

Every tool MUST exit with one of these codes. The AI agent orchestrator
uses exit codes to determine success/failure without parsing JSON.

Exit Code Semantics:
    0 — Success: operation completed, JSON output is authoritative.
    1 — Validation Error: input rejected (bad schema, malformed range,
        impact denial). The agent should fix input and retry.
    2 — File Not Found: input file does not exist or is not readable.
    3 — Lock Contention: file is locked by another process; retry after delay.
    4 — Permission Denied: approval token is invalid, expired, revoked,
        or scoped to a different operation. Generate a new token.
    5 — Internal Error: unexpected failure (bug, corrupt file, concurrent
        modification). Report to operator.
"""

from __future__ import annotations

import json
import sys
from enum import IntEnum
from typing import NoReturn


class ExitCode(IntEnum):
    """Standardized exit codes for all excel-agent-tools."""

    SUCCESS = 0
    VALIDATION_ERROR = 1
    FILE_NOT_FOUND = 2
    LOCK_CONTENTION = 3
    PERMISSION_DENIED = 4
    INTERNAL_ERROR = 5

    @property
    def description(self) -> str:
        """Human-readable description of this exit code."""
        descriptions: dict[int, str] = {
            0: "Operation completed successfully",
            1: "Input validation failed",
            2: "Input file does not exist or is not readable",
            3: "File is locked by another process",
            4: "Approval token invalid, expired, or wrong scope",
            5: "Unexpected internal error",
        }
        return descriptions.get(self.value, "Unknown exit code")


def exit_with(
    code: ExitCode, message: str, *, details: dict[str, object] | None = None
) -> NoReturn:
    """Print a JSON error to stdout and exit with the given code.

    This is the canonical way for tools to report errors. The JSON is always
    written to stdout (never stderr) so the AI agent can parse it reliably.
    """
    response = {
        "status": "error",
        "exit_code": int(code),
        "error": message,
        "details": details or {},
    }
    print(json.dumps(response, indent=2))
    sys.exit(int(code))
