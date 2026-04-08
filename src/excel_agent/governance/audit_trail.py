"""
Pluggable audit logging for excel-agent-tools.

This module implements a pluggable audit trail system that logs all
destructive operations. The default backend is JSONL (one JSON object
per line), but the system supports:
- File-based JSONL (default): .excel_agent_audit.jsonl
- Webhook: POST to configurable endpoint
- Composite: Fan-out to multiple backends simultaneously
- Custom: User-defined backends via AuditBackend protocol

Each audit event includes:
- Timestamp (ISO 8601 UTC)
- Operation details (tool name, scope, target)
- Actor identification (token nonce, optionally user)
- Impact metrics (cells modified, formulas affected)
- File version hash (for forensics)
"""

from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Protocol

import requests

logger = logging.getLogger(__name__)

DEFAULT_AUDIT_FILE = ".excel_agent_audit.jsonl"


@dataclass
class AuditEvent:
    """An audit event representing a single destructive operation.

    Attributes:
        timestamp: ISO 8601 UTC timestamp
        tool: Name of the CLI tool that performed the operation
        scope: Token scope (e.g., "sheet:delete")
        target_file: Path to the affected workbook
        file_version_hash: Geometry hash of file at operation time
        actor_nonce: Token nonce identifying the operation
        operation_details: Additional details (sheet name, range, etc.)
        impact: Impact metrics (cells_modified, formulas_updated)
        success: Whether operation succeeded
        exit_code: Tool exit code
    """

    timestamp: str
    tool: str
    scope: str
    target_file: str
    file_version_hash: str
    actor_nonce: str
    operation_details: dict[str, Any] = field(default_factory=dict)
    impact: dict[str, Any] = field(default_factory=dict)
    success: bool = True
    exit_code: int = 0

    @classmethod
    def create(
        cls,
        tool: str,
        scope: str,
        target_file: Path,
        file_version_hash: str,
        actor_nonce: str,
        operation_details: dict[str, Any] | None = None,
        impact: dict[str, Any] | None = None,
        success: bool = True,
        exit_code: int = 0,
    ) -> AuditEvent:
        """Factory method to create an AuditEvent with current timestamp."""
        return cls(
            timestamp=datetime.now(UTC).isoformat(),
            tool=tool,
            scope=scope,
            target_file=str(target_file),
            file_version_hash=file_version_hash,
            actor_nonce=actor_nonce,
            operation_details=operation_details or {},
            impact=impact or {},
            success=success,
            exit_code=exit_code,
        )

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for serialization."""
        return asdict(self)

    def to_json(self) -> str:
        """Serialize to JSON string."""
        return json.dumps(self.to_dict(), ensure_ascii=False)


class AuditBackend(Protocol):
    """Protocol for audit backends.

    Implement this protocol to create custom audit backends.
    """

    def log_event(self, event: AuditEvent) -> None:
        """Log an audit event."""
        ...


class JsonlAuditBackend:
    """Default audit backend: append to JSONL file.

    Each event is a single line of JSON. File is opened in append mode
    and flushed after each write for durability.
    """

    def __init__(self, audit_file: Path | None = None):
        """Initialize with optional custom audit file path.

        Args:
            audit_file: Path to JSONL file. Defaults to .excel_agent_audit.jsonl
        """
        self._audit_file = audit_file or Path(DEFAULT_AUDIT_FILE)

    def log_event(self, event: AuditEvent) -> None:
        """Append event to JSONL file."""
        line = event.to_json() + "\n"
        with open(self._audit_file, "a", encoding="utf-8") as f:
            f.write(line)
            f.flush()

        logger.debug("Audit event logged to %s: %s", self._audit_file, event.tool)


class WebhookAuditBackend:
    """Webhook audit backend: POST events to HTTP endpoint.

    Useful for SIEM integration or centralized logging.
    """

    def __init__(
        self,
        endpoint: str,
        headers: dict[str, Any] | None = None,
        timeout: float = 5.0,
    ):
        """Initialize with webhook configuration.

        Args:
            endpoint: HTTP URL to POST events to
            headers: Optional additional headers
            timeout: Request timeout in seconds
        """
        self._endpoint = endpoint
        self._headers: dict[str, Any] = headers or {}
        self._timeout = timeout

    def log_event(self, event: AuditEvent) -> None:
        """POST event to webhook endpoint."""
        try:
            response = requests.post(
                self._endpoint,
                json=event.to_dict(),
                headers=self._headers,
                timeout=self._timeout,
            )
            response.raise_for_status()
            logger.debug("Audit event sent to webhook: %s", self._endpoint)
        except requests.RequestException as e:
            logger.warning("Failed to send audit event to webhook: %s", e)


class CompositeAuditBackend:
    """Composite backend: fan-out events to multiple backends.

    Useful when you need both local JSONL and remote webhook logging.
    """

    def __init__(self, backends: list[AuditBackend]):
        """Initialize with list of backends.

        Args:
            backends: List of AuditBackend instances to fan-out to.
        """
        self._backends = backends

    def log_event(self, event: AuditEvent) -> None:
        """Log event to all backends (best effort)."""
        for backend in self._backends:
            try:
                backend.log_event(event)
            except Exception as e:
                logger.warning("Audit backend %s failed: %s", type(backend).__name__, e)


class AuditTrail:
    """Pluggable audit trail for excel-agent-tools.

    Usage:
        # Default JSONL backend
        audit = AuditTrail()

        # Custom backend
        audit = AuditTrail(backend=WebhookAuditBackend("https://logs.example.com"))

        # Composite backend (JSONL + Webhook)
        audit = AuditTrail(backend=CompositeAuditBackend([
            JsonlAuditBackend(),
            WebhookAuditBackend("https://logs.example.com"),
        ]))

        # Log an event
        audit.log(
            tool="xls-delete-sheet",
            scope="sheet:delete",
            target_file=Path("workbook.xlsx"),
            file_version_hash="sha256:abc123...",
            actor_nonce="def456...",
            operation_details={"sheet_name": "OldData"},
            impact={"sheets_deleted": 1},
        )
    """

    def __init__(self, backend: AuditBackend | None = None):
        """Initialize audit trail with optional custom backend.

        Args:
            backend: AuditBackend instance. Defaults to JsonlAuditBackend.
        """
        self._backend = backend or JsonlAuditBackend()

    def log(
        self,
        tool: str,
        scope: str,
        target_file: Path,
        file_version_hash: str,
        actor_nonce: str,
        operation_details: dict[str, Any] | None = None,
        impact: dict[str, Any] | None = None,
        success: bool = True,
        exit_code: int = 0,
    ) -> None:
        """Log an audit event.

        Args:
            tool: Name of the CLI tool
            scope: Token scope
            target_file: Path to affected workbook
            file_version_hash: Version hash at time of operation
            actor_nonce: Token nonce identifying the operation
            operation_details: Additional details (optional)
            impact: Impact metrics (optional)
            success: Whether operation succeeded
            exit_code: Tool exit code
        """
        event = AuditEvent.create(
            tool=tool,
            scope=scope,
            target_file=target_file,
            file_version_hash=file_version_hash,
            actor_nonce=actor_nonce,
            operation_details=operation_details,
            impact=impact,
            success=success,
            exit_code=exit_code,
        )

        self._backend.log_event(event)

    def read_events(
        self,
        audit_file: Path | None = None,
        limit: int | None = None,
    ) -> list[AuditEvent]:
        """Read audit events from JSONL file (for inspection).

        Args:
            audit_file: Path to JSONL file. Defaults to .excel_agent_audit.jsonl
            limit: Maximum number of events to read (newest first).

        Returns:
            List of AuditEvent objects.
        """
        file_path = audit_file or Path(DEFAULT_AUDIT_FILE)
        if not file_path.exists():
            return []

        events: list[AuditEvent] = []
        with open(file_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    data = json.loads(line)
                    events.append(AuditEvent(**data))
                except (json.JSONDecodeError, TypeError) as e:
                    logger.warning("Failed to parse audit event: %s", e)

        # Return newest first, limited
        events.reverse()
        if limit:
            events = events[:limit]

        return events
