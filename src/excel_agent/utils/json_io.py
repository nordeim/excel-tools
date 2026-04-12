"""
Standardized JSON output formatting for all excel-agent-tools.

Every tool produces a consistent JSON envelope via build_response().
The ExcelAgentEncoder handles non-standard types (datetime, Path, bytes, Decimal).
print_json() writes exclusively to stdout — never stderr — so the AI agent
can reliably capture structured output.
"""

from __future__ import annotations

import json
import sys
from datetime import UTC, date, datetime, timedelta, time
from decimal import Decimal
from pathlib import Path
from typing import Any


class ExcelAgentEncoder(json.JSONEncoder):
    """Custom JSON encoder for types commonly found in Excel workbook data.

    Handles:
        datetime/date/time → ISO 8601 string
        timedelta           → total seconds (float)
        Path                → string (POSIX path)
        bytes               → hex string
        Decimal             → float
        set/frozenset       → sorted list
    """

    def default(self, o: object) -> Any:
        if isinstance(o, datetime):
            return o.isoformat()
        if isinstance(o, date):
            return o.isoformat()
        if isinstance(o, time):
            return o.isoformat()
        if isinstance(o, timedelta):
            return o.total_seconds()
        if isinstance(o, Path):
            return str(o)
        if isinstance(o, bytes):
            return o.hex()
        if isinstance(o, Decimal):
            return float(o)
        if isinstance(o, (set, frozenset)):
            return sorted(o)
        return super().default(o)


def build_response(
    status: str,
    data: Any,
    *,
    workbook_version: str = "",
    impact: dict[str, Any] | None = None,
    warnings: list[str] | None = None,
    exit_code: int = 0,
    guidance: str | None = None,
) -> dict[str, Any]:
    """Build the standardized JSON response envelope.

    Args:
        status: One of "success", "error", "warning", "denied".
        data: The primary payload (tool-specific).
        workbook_version: Geometry hash of the workbook (sha256:...).
        impact: Mutation metrics (cells_modified, formulas_updated, etc.).
        warnings: Non-fatal issues encountered during execution.
        exit_code: Integer exit code (0-5).
        guidance: Prescriptive next-step for the agent (used in denials).

    Returns:
        A dict ready for JSON serialization and printing.
    """
    response: dict[str, Any] = {
        "status": status,
        "exit_code": exit_code,
        "timestamp": datetime.now(UTC).isoformat(),
        "workbook_version": workbook_version,
        "data": data,
    }

    if impact is not None:
        response["impact"] = impact
    else:
        response["impact"] = {"cells_modified": 0, "formulas_updated": 0}

    if warnings:
        response["warnings"] = warnings
    else:
        response["warnings"] = []

    if guidance is not None:
        response["guidance"] = guidance

    return response


def print_json(data: dict[str, Any], *, indent: int = 2) -> None:
    """Print JSON to stdout. Never writes to stderr.

    The AI agent captures stdout for structured data. All diagnostic
    messages go through the JSON envelope, never as raw text.
    """
    output = json.dumps(data, indent=indent, cls=ExcelAgentEncoder, ensure_ascii=False)
    sys.stdout.write(output)
    sys.stdout.write("\n")
    sys.stdout.flush()
