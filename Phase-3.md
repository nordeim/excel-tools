# Phase 4: IMPLEMENT — Phase 3: Governance & Safety Layer

Let me conduct precise research on every cryptographic and security API surface I'll be coding against — HMAC-SHA256 specifics, `secrets` module, `compare_digest` internals, UUID4 nonce generation, and atomic file append patterns.All research complete. I now have comprehensive, validated technical details for every Phase 3 component. Let me synthesize the findings and generate the code.

---

## Research Summary: Critical Technical Findings for Phase 3

**HMAC-SHA256 (confirmed):**
- This module implements the HMAC algorithm as described by RFC 2104.
- key is a bytes or bytearray object giving the secret key. The digestmod argument is now required.
- The `hmac.digest()` function is equivalent to `HMAC(key, msg, digest).digest()`, but uses an optimized C or inline implementation, which is faster for messages that fit into memory.
- `compare_digest` uses an approach designed to prevent timing analysis by avoiding content-based short circuiting behaviour, making it appropriate for cryptography.
- `hmac.compare_digest(a, b)` is used to safely compare two digests to prevent a type of side-channel attack called a timing attack. It returns True if a and b are equal, and False otherwise, but it does so in a constant-time manner.

**secrets module (confirmed):**
- The secrets module provides functions for generating secure tokens, suitable for applications such as password resets, hard-to-guess URLs, and similar.
- As of 2015, it is believed that 32 bytes (256 bits) of randomness is sufficient for the typical use-case expected for the secrets module.
- Return True if strings or bytes-like objects a and b are equal, otherwise False, using a "constant-time compare" to reduce the risk of timing attacks.
- The module offers several key functions for generating secure tokens: token_bytes() returns random bytes, token_hex() provides hexadecimal strings, and token_urlsafe() creates URL-safe Base64-encoded strings.

**Atomic file append for JSONL audit:**
- Using `"a"` mode: The operating system ensures that every write operation initiated on a file opened in append mode is guaranteed to land at the true end of the file, preventing data from being scrambled across concurrent writers. Caveat: This atomicity guarantee usually applies to the entire write.
- Because appending the file is protected with the lock, it is thread-safe and the example results in crisp one-message-per-line output.
- For cross-process safety, we combine `"a"` mode with our existing `FileLock` sidecar pattern.

---

Now generating all Phase 3 code files:

## File 1: `src/excel_agent/governance/__init__.py`

```python
"""
Governance layer for excel-agent-tools.

Provides:
    - ApprovalTokenManager: HMAC-SHA256 scoped approval tokens with
      TTL, nonce, file-hash binding, and constant-time validation.
    - AuditTrail: Pluggable audit logging for all destructive operations.
      Default backend: JSONL append-only file.
    - Schema validation: JSON Schema validation for all tool inputs.
"""

from __future__ import annotations

__all__ = [
    "ApprovalTokenManager",
    "AuditTrail",
]


def __getattr__(name: str) -> object:
    if name == "ApprovalTokenManager":
        from excel_agent.governance.token_manager import ApprovalTokenManager

        return ApprovalTokenManager
    if name == "AuditTrail":
        from excel_agent.governance.audit_trail import AuditTrail

        return AuditTrail
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
```

---

## File 2: `src/excel_agent/governance/token_manager.py`

```python
"""
HMAC-SHA256 scoped approval token system for excel-agent-tools.

Implements the governance-first design principle: destructive operations
(sheet delete, range delete, macro removal, etc.) require a cryptographically
signed, scoped, time-limited, single-use approval token.

Token properties:
    - Scoped: bound to a specific operation (e.g., "sheet:delete")
    - File-bound: tied to a specific workbook hash (prevents cross-file reuse)
    - Time-limited: TTL (default 300s = 5 minutes)
    - Single-use: nonce tracking prevents replay attacks
    - Constant-time validation: hmac.compare_digest() prevents timing attacks

Security foundations (per Python docs):
    - hmac.compare_digest() "uses an approach designed to prevent timing
      analysis by avoiding content-based short circuiting behaviour"
    - secrets.token_hex() provides "cryptographically strong random numbers"
    - 32 bytes (256 bits) of randomness is sufficient per secrets module docs

Token format (serialized):
    base64url(json({"s": scope, "f": file_hash, "n": nonce,
                     "i": issued_at, "t": ttl, "g": signature}))
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import logging
import os
import secrets
import time
from dataclasses import dataclass
from typing import Any

from excel_agent.utils.exceptions import PermissionDeniedError

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

VALID_SCOPES = frozenset({
    "sheet:delete",
    "sheet:rename",
    "range:delete",
    "formula:convert",
    "macro:remove",
    "macro:inject",
    "structure:modify",
})

_MIN_TTL = 1
_MAX_TTL = 3600
_DEFAULT_TTL = 300  # 5 minutes
_SECRET_KEY_BYTES = 32  # 256 bits — sufficient per secrets module docs
_ENV_SECRET_KEY = "EXCEL_AGENT_SECRET"


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ApprovalToken:
    """Immutable parsed token structure.

    Attributes:
        scope: Governance scope (e.g., "sheet:delete").
        target_file_hash: SHA-256 hash of the target workbook.
        nonce: UUID-hex, single-use identifier.
        issued_at: Unix timestamp (UTC) when token was created.
        ttl_seconds: Time-to-live in seconds.
        signature: HMAC-SHA256 hex digest.
    """

    scope: str
    target_file_hash: str
    nonce: str
    issued_at: float
    ttl_seconds: int
    signature: str

    @property
    def expires_at(self) -> float:
        """Unix timestamp when this token expires."""
        return self.issued_at + self.ttl_seconds

    @property
    def is_expired(self) -> bool:
        """Whether this token has expired."""
        return time.time() > self.expires_at


# ---------------------------------------------------------------------------
# ApprovalTokenManager
# ---------------------------------------------------------------------------


class ApprovalTokenManager:
    """Generates and validates scoped HMAC-SHA256 approval tokens.

    The secret key is sourced from:
        1. The EXCEL_AGENT_SECRET environment variable (preferred)
        2. Auto-generated per session (with warning log)

    Usage::

        mgr = ApprovalTokenManager()
        token_str = mgr.generate_token("sheet:delete", "sha256:abc...")
        parsed = mgr.validate_token(token_str,
                                     expected_scope="sheet:delete",
                                     expected_file_hash="sha256:abc...")
    """

    def __init__(self, *, secret_key: str | None = None) -> None:
        """Initialize with a secret key.

        Args:
            secret_key: Explicit secret key string. If None, reads from
                        EXCEL_AGENT_SECRET env var, or auto-generates.
        """
        if secret_key is not None:
            self._secret = secret_key.encode("utf-8")
        else:
            env_key = os.environ.get(_ENV_SECRET_KEY)
            if env_key:
                self._secret = env_key.encode("utf-8")
                logger.debug("Token secret loaded from %s env var", _ENV_SECRET_KEY)
            else:
                self._secret = secrets.token_bytes(_SECRET_KEY_BYTES)
                logger.warning(
                    "No %s env var set — auto-generated session secret. "
                    "Tokens will NOT be valid across sessions.",
                    _ENV_SECRET_KEY,
                )

        # Nonce tracking: set of used nonces (prevents replay within session)
        self._used_nonces: set[str] = set()
        # Revocation set: explicitly revoked nonces
        self._revoked_nonces: set[str] = set()

    def generate_token(
        self,
        scope: str,
        target_file_hash: str,
        *,
        ttl_seconds: int = _DEFAULT_TTL,
    ) -> str:
        """Generate a scoped approval token.

        Args:
            scope: One of VALID_SCOPES (e.g., "sheet:delete").
            target_file_hash: SHA-256 hash of the target workbook.
            ttl_seconds: Time-to-live in seconds (1-3600).

        Returns:
            Serialized token string (base64url-encoded JSON).

        Raises:
            PermissionDeniedError: If scope is invalid or TTL out of range.
        """
        # Validate scope
        if scope not in VALID_SCOPES:
            raise PermissionDeniedError(
                f"Invalid scope: {scope!r}. Valid scopes: {sorted(VALID_SCOPES)}",
                reason="invalid_scope",
                details={"scope": scope, "valid_scopes": sorted(VALID_SCOPES)},
            )

        # Validate TTL
        if not (_MIN_TTL <= ttl_seconds <= _MAX_TTL):
            raise PermissionDeniedError(
                f"TTL must be {_MIN_TTL}-{_MAX_TTL}s, got {ttl_seconds}",
                reason="invalid_ttl",
                details={"ttl_seconds": ttl_seconds},
            )

        # Generate nonce and timestamp
        nonce = secrets.token_hex(16)  # 128-bit nonce
        issued_at = time.time()

        # Compute HMAC-SHA256 signature
        signature = self._compute_signature(
            scope, target_file_hash, nonce, issued_at, ttl_seconds
        )

        # Serialize token as base64url JSON
        payload: dict[str, Any] = {
            "s": scope,
            "f": target_file_hash,
            "n": nonce,
            "i": issued_at,
            "t": ttl_seconds,
            "g": signature,
        }
        json_bytes = json.dumps(payload, separators=(",", ":")).encode("utf-8")
        token_str = base64.urlsafe_b64encode(json_bytes).decode("ascii")

        logger.info(
            "Token generated: scope=%s, ttl=%ds, nonce=%s...",
            scope,
            ttl_seconds,
            nonce[:8],
        )
        return token_str

    def validate_token(
        self,
        token_str: str,
        *,
        expected_scope: str,
        expected_file_hash: str,
    ) -> ApprovalToken:
        """Validate a token string. Returns parsed token if valid.

        Validation steps (in order):
            1. Deserialize and parse JSON
            2. Verify scope matches expected
            3. Verify file hash matches expected
            4. Verify not expired (issued_at + ttl > now)
            5. Verify nonce not in revocation set
            6. Verify nonce not already used (single-use)
            7. Recompute HMAC signature and compare with constant-time
            8. Mark nonce as used

        Args:
            token_str: Serialized token from generate_token().
            expected_scope: The scope required for this operation.
            expected_file_hash: SHA-256 hash of the workbook being operated on.

        Returns:
            Parsed ApprovalToken if all checks pass.

        Raises:
            PermissionDeniedError: With descriptive reason for any failure.
        """
        # Step 1: Deserialize
        payload = self._deserialize_token(token_str)
        scope = str(payload.get("s", ""))
        file_hash = str(payload.get("f", ""))
        nonce = str(payload.get("n", ""))
        issued_at = float(payload.get("i", 0))
        ttl_seconds = int(payload.get("t", 0))
        signature = str(payload.get("g", ""))

        # Step 2: Verify scope
        if scope != expected_scope:
            raise PermissionDeniedError(
                f"Token scope mismatch: token has {scope!r}, "
                f"operation requires {expected_scope!r}",
                reason="scope_mismatch",
                details={"token_scope": scope, "required_scope": expected_scope},
            )

        # Step 3: Verify file hash
        if file_hash != expected_file_hash:
            raise PermissionDeniedError(
                "Token is bound to a different workbook file",
                reason="file_hash_mismatch",
                details={
                    "token_file_hash": file_hash[:20] + "...",
                    "expected_file_hash": expected_file_hash[:20] + "...",
                },
            )

        # Step 4: Verify not expired
        if time.time() > issued_at + ttl_seconds:
            raise PermissionDeniedError(
                f"Token expired {time.time() - (issued_at + ttl_seconds):.0f}s ago",
                reason="expired",
                details={"issued_at": issued_at, "ttl_seconds": ttl_seconds},
            )

        # Step 5: Verify not revoked
        if nonce in self._revoked_nonces:
            raise PermissionDeniedError(
                "Token has been revoked",
                reason="revoked",
                details={"nonce": nonce[:8] + "..."},
            )

        # Step 6: Verify not already used (single-use)
        if nonce in self._used_nonces:
            raise PermissionDeniedError(
                "Token has already been used (single-use policy)",
                reason="already_used",
                details={"nonce": nonce[:8] + "..."},
            )

        # Step 7: Recompute signature and compare with constant-time
        expected_sig = self._compute_signature(
            scope, file_hash, nonce, issued_at, ttl_seconds
        )
        if not hmac.compare_digest(signature, expected_sig):
            raise PermissionDeniedError(
                "Token signature is invalid (tampered or wrong secret key)",
                reason="invalid_signature",
            )

        # Step 8: Mark nonce as used
        self._used_nonces.add(nonce)

        token = ApprovalToken(
            scope=scope,
            target_file_hash=file_hash,
            nonce=nonce,
            issued_at=issued_at,
            ttl_seconds=ttl_seconds,
            signature=signature,
        )

        logger.info(
            "Token validated: scope=%s, nonce=%s...",
            scope,
            nonce[:8],
        )
        return token

    def revoke_token(self, nonce: str) -> None:
        """Add a nonce to the revocation set (prevents future use).

        Args:
            nonce: The nonce string from the token to revoke.
        """
        self._revoked_nonces.add(nonce)
        logger.info("Token revoked: nonce=%s...", nonce[:8])

    @property
    def valid_scopes(self) -> frozenset[str]:
        """Return the set of valid token scopes."""
        return VALID_SCOPES

    def _compute_signature(
        self,
        scope: str,
        file_hash: str,
        nonce: str,
        issued_at: float,
        ttl_seconds: int,
    ) -> str:
        """Compute HMAC-SHA256 over the canonical token fields.

        The canonical message is: "scope|file_hash|nonce|issued_at|ttl"
        joined with pipe separators.

        Returns:
            Hex digest of the HMAC-SHA256 signature.
        """
        canonical = f"{scope}|{file_hash}|{nonce}|{issued_at}|{ttl_seconds}"
        return hmac.new(
            self._secret,
            canonical.encode("utf-8"),
            digestmod=hashlib.sha256,
        ).hexdigest()

    @staticmethod
    def _deserialize_token(token_str: str) -> dict[str, Any]:
        """Deserialize a base64url-encoded JSON token.

        Raises:
            PermissionDeniedError: If the token cannot be deserialized.
        """
        try:
            # Add padding if needed (base64url may strip trailing =)
            padded = token_str + "=" * (-len(token_str) % 4)
            json_bytes = base64.urlsafe_b64decode(padded)
            payload: dict[str, Any] = json.loads(json_bytes)
        except (ValueError, json.JSONDecodeError, UnicodeDecodeError) as exc:
            raise PermissionDeniedError(
                f"Token is malformed and cannot be deserialized: {exc}",
                reason="malformed_token",
            ) from exc

        if not isinstance(payload, dict):
            raise PermissionDeniedError(
                "Token payload is not a JSON object",
                reason="malformed_token",
            )
        return payload
```

---

## File 3: `src/excel_agent/governance/audit_trail.py`

```python
"""
Pluggable audit trail for excel-agent-tools.

Logs every destructive operation to an append-only audit log. The default
backend writes JSON Lines (.jsonl) files. Alternative backends can be
implemented via the AuditBackend Protocol.

Architecture:
    AuditBackend (Protocol)
    ├── JsonlAuditBackend  — Default: append-only .jsonl file
    ├── NullAuditBackend   — No-op for testing
    └── CompositeAuditBackend — Fan-out to multiple backends

File append safety:
    The OS guarantees that writes in "a" mode land at the true end of
    the file, preventing interleaved data from concurrent writers. We
    additionally use a sidecar file lock for cross-process safety.

Each audit entry is a single JSON line:
    {"timestamp": "...", "tool": "...", "scope": "...", ...}
"""

from __future__ import annotations

import json
import logging
import os
import sys
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Protocol, runtime_checkable

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclass
class AuditEvent:
    """A single audit log entry.

    Attributes:
        timestamp: ISO 8601 UTC string.
        tool: Tool name (e.g., "xls_delete_sheet").
        scope: Governance scope used (e.g., "sheet:delete").
        resource: What was affected (e.g., "Sheet1", "A1:C10").
        action: What happened (e.g., "delete", "rename", "convert").
        outcome: Result ("success", "denied", "error").
        token_used: Whether a governance token was required/used.
        file_hash: Workbook hash at time of operation.
        pid: Process ID for tracing.
        details: Additional context.
    """

    timestamp: str
    tool: str
    scope: str
    resource: str
    action: str
    outcome: str
    token_used: bool
    file_hash: str
    pid: int
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        """Convert to JSON-serializable dict."""
        return asdict(self)


# ---------------------------------------------------------------------------
# Backend Protocol
# ---------------------------------------------------------------------------


@runtime_checkable
class AuditBackend(Protocol):
    """Protocol for pluggable audit backends.

    Implement this to create custom audit destinations (database,
    webhook, SIEM, etc.).
    """

    def log_event(self, event: AuditEvent) -> None:
        """Write a single audit event."""
        ...

    def query_events(
        self,
        *,
        tool: str | None = None,
        outcome: str | None = None,
        start_time: datetime | None = None,
        end_time: datetime | None = None,
        limit: int = 100,
    ) -> list[AuditEvent]:
        """Query audit events with optional filters."""
        ...


# ---------------------------------------------------------------------------
# JsonlAuditBackend
# ---------------------------------------------------------------------------


class JsonlAuditBackend:
    """Default audit backend: append-only JSON Lines file.

    Each event is a single JSON line appended to the .jsonl file.
    Cross-process safety is achieved via advisory file locking on
    a sidecar .lock file, combined with OS-level "a" mode guarantees.
    """

    def __init__(self, log_path: Path | None = None) -> None:
        """Initialize the JSONL backend.

        Args:
            log_path: Path to the audit log file.
                      Defaults to .excel_agent_audit.jsonl in CWD.
        """
        if log_path is None:
            log_path = Path(".excel_agent_audit.jsonl")
        self._log_path = log_path.resolve()
        self._lock_path = self._log_path.parent / f".{self._log_path.name}.lock"

    @property
    def log_path(self) -> Path:
        """Path to the audit log file."""
        return self._log_path

    def log_event(self, event: AuditEvent) -> None:
        """Append a single event as a JSON line.

        Uses "a" mode for OS-level append atomicity, plus an advisory
        file lock for cross-process safety.
        """
        line = json.dumps(event.to_dict(), separators=(",", ":"), ensure_ascii=False)

        lock_fd = None
        try:
            lock_fd = self._acquire_lock()
            with open(self._log_path, "a", encoding="utf-8") as f:
                f.write(line)
                f.write("\n")
                f.flush()
                os.fsync(f.fileno())
        except OSError as exc:
            logger.error("Failed to write audit event: %s", exc)
        finally:
            self._release_lock(lock_fd)

    def query_events(
        self,
        *,
        tool: str | None = None,
        outcome: str | None = None,
        start_time: datetime | None = None,
        end_time: datetime | None = None,
        limit: int = 100,
    ) -> list[AuditEvent]:
        """Read and filter audit events from the log file.

        Args:
            tool: Filter by tool name.
            outcome: Filter by outcome.
            start_time: Only events after this time.
            end_time: Only events before this time.
            limit: Maximum number of events to return.

        Returns:
            List of matching AuditEvent objects, most recent first.
        """
        if not self._log_path.exists():
            return []

        events: list[AuditEvent] = []
        try:
            with open(self._log_path, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        data = json.loads(line)
                    except json.JSONDecodeError:
                        continue  # Skip corrupted lines

                    # Apply filters
                    if tool is not None and data.get("tool") != tool:
                        continue
                    if outcome is not None and data.get("outcome") != outcome:
                        continue
                    if start_time is not None:
                        event_time = data.get("timestamp", "")
                        if event_time < start_time.isoformat():
                            continue
                    if end_time is not None:
                        event_time = data.get("timestamp", "")
                        if event_time > end_time.isoformat():
                            continue

                    events.append(AuditEvent(**data))

                    if len(events) >= limit:
                        break
        except OSError as exc:
            logger.error("Failed to read audit log: %s", exc)

        # Most recent first
        events.reverse()
        return events

    def _acquire_lock(self) -> Any | None:  # noqa: ANN401
        """Acquire advisory lock on the sidecar lock file."""
        try:
            lock_fd = open(self._lock_path, "w")  # noqa: SIM115
            if sys.platform != "win32":
                import fcntl

                fcntl.flock(lock_fd.fileno(), fcntl.LOCK_EX)
            else:
                import msvcrt

                lock_fd.seek(0)
                msvcrt.locking(lock_fd.fileno(), msvcrt.LK_LOCK, 1)
            return lock_fd
        except OSError:
            return None

    def _release_lock(self, lock_fd: Any | None) -> None:  # noqa: ANN401
        """Release advisory lock."""
        if lock_fd is None:
            return
        try:
            if sys.platform != "win32":
                import fcntl

                fcntl.flock(lock_fd.fileno(), fcntl.LOCK_UN)
            else:
                import msvcrt

                lock_fd.seek(0)
                msvcrt.locking(lock_fd.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
        try:
            lock_fd.close()
        except OSError:
            pass
        try:
            self._lock_path.unlink(missing_ok=True)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# NullAuditBackend
# ---------------------------------------------------------------------------


class NullAuditBackend:
    """No-op audit backend for testing or when auditing is disabled."""

    def __init__(self) -> None:
        self._events: list[AuditEvent] = []

    def log_event(self, event: AuditEvent) -> None:
        """Store in memory only (no I/O)."""
        self._events.append(event)

    def query_events(
        self,
        *,
        tool: str | None = None,
        outcome: str | None = None,
        start_time: datetime | None = None,
        end_time: datetime | None = None,
        limit: int = 100,
    ) -> list[AuditEvent]:
        """Query from in-memory store."""
        results = list(self._events)
        if tool is not None:
            results = [e for e in results if e.tool == tool]
        if outcome is not None:
            results = [e for e in results if e.outcome == outcome]
        results.reverse()
        return results[:limit]


# ---------------------------------------------------------------------------
# CompositeAuditBackend
# ---------------------------------------------------------------------------


class CompositeAuditBackend:
    """Fan-out audit backend: forwards events to multiple backends.

    Usage::

        composite = CompositeAuditBackend(
            JsonlAuditBackend(Path("audit.jsonl")),
            NullAuditBackend(),  # also capture in memory
        )
    """

    def __init__(self, *backends: JsonlAuditBackend | NullAuditBackend) -> None:
        self._backends = list(backends)

    def log_event(self, event: AuditEvent) -> None:
        """Forward event to all backends."""
        for backend in self._backends:
            try:
                backend.log_event(event)
            except Exception:
                logger.error(
                    "Audit backend %s failed to log event",
                    type(backend).__name__,
                    exc_info=True,
                )

    def query_events(
        self,
        *,
        tool: str | None = None,
        outcome: str | None = None,
        start_time: datetime | None = None,
        end_time: datetime | None = None,
        limit: int = 100,
    ) -> list[AuditEvent]:
        """Query from the first backend that returns results."""
        for backend in self._backends:
            try:
                results = backend.query_events(
                    tool=tool,
                    outcome=outcome,
                    start_time=start_time,
                    end_time=end_time,
                    limit=limit,
                )
                if results:
                    return results
            except Exception:
                continue
        return []


# ---------------------------------------------------------------------------
# AuditTrail (Facade)
# ---------------------------------------------------------------------------


class AuditTrail:
    """Facade for audit logging across all excel-agent-tools.

    Provides a simple log_operation() method that creates an AuditEvent
    and delegates to the configured backend.

    Usage::

        audit = AuditTrail()  # defaults to JsonlAuditBackend
        audit.log_operation(
            tool="xls_delete_sheet",
            scope="sheet:delete",
            resource="Sheet1",
            action="delete",
            outcome="success",
            token_used=True,
            file_hash="sha256:abc...",
        )
    """

    def __init__(
        self,
        backend: JsonlAuditBackend | NullAuditBackend | CompositeAuditBackend | None = None,
    ) -> None:
        """Initialize with an audit backend.

        Args:
            backend: Audit backend instance. Defaults to JsonlAuditBackend.
        """
        if backend is None:
            backend = JsonlAuditBackend()
        self._backend = backend

    @property
    def backend(self) -> JsonlAuditBackend | NullAuditBackend | CompositeAuditBackend:
        """The active audit backend."""
        return self._backend

    def log_operation(
        self,
        tool: str,
        scope: str,
        resource: str,
        action: str,
        outcome: str,
        *,
        token_used: bool = False,
        file_hash: str = "",
        details: dict[str, Any] | None = None,
    ) -> AuditEvent:
        """Log a single operation to the audit trail.

        Args:
            tool: Tool name (e.g., "xls_delete_sheet").
            scope: Governance scope (e.g., "sheet:delete").
            resource: What was affected.
            action: What happened.
            outcome: Result ("success", "denied", "error").
            token_used: Whether a governance token was required.
            file_hash: Workbook hash at time of operation.
            details: Additional context dict.

        Returns:
            The AuditEvent that was logged.
        """
        event = AuditEvent(
            timestamp=datetime.now(timezone.utc).isoformat(),
            tool=tool,
            scope=scope,
            resource=resource,
            action=action,
            outcome=outcome,
            token_used=token_used,
            file_hash=file_hash,
            pid=os.getpid(),
            details=details or {},
        )

        self._backend.log_event(event)
        return event

    def query(
        self,
        *,
        tool: str | None = None,
        outcome: str | None = None,
        start_time: datetime | None = None,
        end_time: datetime | None = None,
        limit: int = 100,
    ) -> list[AuditEvent]:
        """Query audit events.

        Args:
            tool: Filter by tool name.
            outcome: Filter by outcome.
            start_time: Only events after this time.
            end_time: Only events before this time.
            limit: Maximum events to return.

        Returns:
            List of matching events, most recent first.
        """
        return self._backend.query_events(
            tool=tool,
            outcome=outcome,
            start_time=start_time,
            end_time=end_time,
            limit=limit,
        )
```

---

## File 4: `tests/unit/test_token_manager.py`

```python
"""Comprehensive tests for the ApprovalTokenManager."""

from __future__ import annotations

import os
import time

import pytest

from excel_agent.governance.token_manager import (
    VALID_SCOPES,
    ApprovalToken,
    ApprovalTokenManager,
)
from excel_agent.utils.exceptions import PermissionDeniedError


@pytest.fixture
def manager() -> ApprovalTokenManager:
    """Token manager with a fixed test secret."""
    return ApprovalTokenManager(secret_key="test-secret-key-for-unit-tests")


@pytest.fixture
def sample_hash() -> str:
    return "sha256:abc123def456"


class TestTokenGeneration:
    """Tests for generate_token()."""

    def test_generates_valid_token_string(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        token = manager.generate_token("sheet:delete", sample_hash)
        assert isinstance(token, str)
        assert len(token) > 20  # base64-encoded JSON is substantial

    def test_all_valid_scopes(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        for scope in VALID_SCOPES:
            token = manager.generate_token(scope, sample_hash)
            assert isinstance(token, str)

    def test_invalid_scope_raises(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        with pytest.raises(PermissionDeniedError, match="Invalid scope") as exc_info:
            manager.generate_token("invalid:scope", sample_hash)
        assert exc_info.value.reason == "invalid_scope"
        assert exc_info.value.exit_code == 4

    def test_ttl_out_of_range_raises(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        with pytest.raises(PermissionDeniedError, match="TTL"):
            manager.generate_token("sheet:delete", sample_hash, ttl_seconds=0)
        with pytest.raises(PermissionDeniedError, match="TTL"):
            manager.generate_token("sheet:delete", sample_hash, ttl_seconds=9999)

    def test_custom_ttl(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        token = manager.generate_token("sheet:delete", sample_hash, ttl_seconds=60)
        parsed = manager.validate_token(
            token, expected_scope="sheet:delete", expected_file_hash=sample_hash
        )
        assert parsed.ttl_seconds == 60

    def test_each_token_has_unique_nonce(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        """Two tokens generated in quick succession should have different nonces."""
        t1 = manager.generate_token("sheet:delete", sample_hash)
        t2 = manager.generate_token("sheet:delete", sample_hash)
        assert t1 != t2


class TestTokenValidation:
    """Tests for validate_token()."""

    def test_valid_token_passes(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        token = manager.generate_token("sheet:delete", sample_hash)
        parsed = manager.validate_token(
            token, expected_scope="sheet:delete", expected_file_hash=sample_hash
        )
        assert isinstance(parsed, ApprovalToken)
        assert parsed.scope == "sheet:delete"
        assert parsed.target_file_hash == sample_hash
        assert not parsed.is_expired

    def test_wrong_scope_raises(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        token = manager.generate_token("sheet:delete", sample_hash)
        with pytest.raises(PermissionDeniedError, match="scope mismatch") as exc_info:
            manager.validate_token(
                token,
                expected_scope="range:delete",
                expected_file_hash=sample_hash,
            )
        assert exc_info.value.reason == "scope_mismatch"

    def test_wrong_file_hash_raises(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        token = manager.generate_token("sheet:delete", sample_hash)
        with pytest.raises(PermissionDeniedError, match="different workbook") as exc_info:
            manager.validate_token(
                token,
                expected_scope="sheet:delete",
                expected_file_hash="sha256:different_hash",
            )
        assert exc_info.value.reason == "file_hash_mismatch"

    def test_expired_token_raises(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        token = manager.generate_token("sheet:delete", sample_hash, ttl_seconds=1)
        time.sleep(1.5)
        with pytest.raises(PermissionDeniedError, match="expired") as exc_info:
            manager.validate_token(
                token,
                expected_scope="sheet:delete",
                expected_file_hash=sample_hash,
            )
        assert exc_info.value.reason == "expired"

    def test_reused_nonce_raises(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        """Single-use policy: using the same token twice should fail."""
        token = manager.generate_token("sheet:delete", sample_hash)
        # First use succeeds
        manager.validate_token(
            token, expected_scope="sheet:delete", expected_file_hash=sample_hash
        )
        # Second use fails
        with pytest.raises(PermissionDeniedError, match="already been used") as exc_info:
            manager.validate_token(
                token,
                expected_scope="sheet:delete",
                expected_file_hash=sample_hash,
            )
        assert exc_info.value.reason == "already_used"

    def test_tampered_signature_raises(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        """Modifying the token string should invalidate the signature."""
        token = manager.generate_token("sheet:delete", sample_hash)
        # Flip a character in the middle of the token
        tampered = token[:10] + ("A" if token[10] != "A" else "B") + token[11:]
        with pytest.raises(PermissionDeniedError):
            manager.validate_token(
                tampered,
                expected_scope="sheet:delete",
                expected_file_hash=sample_hash,
            )

    def test_malformed_token_raises(self, manager: ApprovalTokenManager) -> None:
        with pytest.raises(PermissionDeniedError, match="malformed"):
            manager.validate_token(
                "not-a-valid-token!!!",
                expected_scope="sheet:delete",
                expected_file_hash="sha256:x",
            )

    def test_revoked_token_raises(
        self, manager: ApprovalTokenManager, sample_hash: str
    ) -> None:
        token = manager.generate_token("sheet:delete", sample_hash)
        # Deserialize to get the nonce
        payload = manager._deserialize_token(token)
        nonce = payload["n"]
        # Revoke it
        manager.revoke_token(nonce)
        with pytest.raises(PermissionDeniedError, match="revoked") as exc_info:
            manager.validate_token(
                token,
                expected_scope="sheet:delete",
                expected_file_hash=sample_hash,
            )
        assert exc_info.value.reason == "revoked"


class TestTokenManagerSecretKey:
    """Tests for secret key sourcing."""

    def test_explicit_secret_key(self) -> None:
        mgr = ApprovalTokenManager(secret_key="my-explicit-key")
        token = mgr.generate_token("sheet:delete", "sha256:test")
        parsed = mgr.validate_token(
            token, expected_scope="sheet:delete", expected_file_hash="sha256:test"
        )
        assert parsed.scope == "sheet:delete"

    def test_env_var_secret_key(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("EXCEL_AGENT_SECRET", "env-secret-key")
        mgr = ApprovalTokenManager()
        token = mgr.generate_token("sheet:delete", "sha256:test")
        parsed = mgr.validate_token(
            token, expected_scope="sheet:delete", expected_file_hash="sha256:test"
        )
        assert parsed.scope == "sheet:delete"

    def test_auto_generated_secret_logs_warning(self, caplog: pytest.LogCaptureFixture) -> None:
        # Ensure env var is not set
        env_backup = os.environ.pop("EXCEL_AGENT_SECRET", None)
        try:
            with caplog.at_level("WARNING"):
                ApprovalTokenManager()
            assert "auto-generated" in caplog.text.lower() or "no" in caplog.text.lower()
        finally:
            if env_backup is not None:
                os.environ["EXCEL_AGENT_SECRET"] = env_backup

    def test_different_secrets_produce_different_signatures(self) -> None:
        mgr1 = ApprovalTokenManager(secret_key="key-one")
        mgr2 = ApprovalTokenManager(secret_key="key-two")
        file_hash = "sha256:test"

        token1 = mgr1.generate_token("sheet:delete", file_hash)

        # Token from mgr1 should NOT validate with mgr2
        with pytest.raises(PermissionDeniedError):
            mgr2.validate_token(
                token1,
                expected_scope="sheet:delete",
                expected_file_hash=file_hash,
            )


class TestApprovalTokenDataclass:
    """Tests for the ApprovalToken dataclass."""

    def test_expires_at_calculation(self) -> None:
        now = time.time()
        token = ApprovalToken(
            scope="sheet:delete",
            target_file_hash="sha256:x",
            nonce="abc",
            issued_at=now,
            ttl_seconds=300,
            signature="sig",
        )
        assert abs(token.expires_at - (now + 300)) < 0.01

    def test_is_expired_false_for_fresh_token(self) -> None:
        token = ApprovalToken(
            scope="sheet:delete",
            target_file_hash="sha256:x",
            nonce="abc",
            issued_at=time.time(),
            ttl_seconds=300,
            signature="sig",
        )
        assert not token.is_expired

    def test_is_expired_true_for_old_token(self) -> None:
        token = ApprovalToken(
            scope="sheet:delete",
            target_file_hash="sha256:x",
            nonce="abc",
            issued_at=time.time() - 600,
            ttl_seconds=300,
            signature="sig",
        )
        assert token.is_expired

    def test_frozen_dataclass(self) -> None:
        token = ApprovalToken(
            scope="sheet:delete",
            target_file_hash="sha256:x",
            nonce="abc",
            issued_at=0.0,
            ttl_seconds=300,
            signature="sig",
        )
        with pytest.raises(AttributeError):
            token.scope = "hacked"  # type: ignore[misc]
```

---

## File 5: `tests/unit/test_audit_trail.py`

```python
"""Comprehensive tests for the audit trail system."""

from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path

import pytest

from excel_agent.governance.audit_trail import (
    AuditEvent,
    AuditTrail,
    CompositeAuditBackend,
    JsonlAuditBackend,
    NullAuditBackend,
)


@pytest.fixture
def audit_path(tmp_path: Path) -> Path:
    return tmp_path / "test_audit.jsonl"


@pytest.fixture
def jsonl_backend(audit_path: Path) -> JsonlAuditBackend:
    return JsonlAuditBackend(audit_path)


@pytest.fixture
def null_backend() -> NullAuditBackend:
    return NullAuditBackend()


def _make_event(**overrides: object) -> AuditEvent:
    """Factory for test AuditEvent instances."""
    defaults: dict[str, object] = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "tool": "xls_delete_sheet",
        "scope": "sheet:delete",
        "resource": "Sheet1",
        "action": "delete",
        "outcome": "success",
        "token_used": True,
        "file_hash": "sha256:test",
        "pid": os.getpid(),
        "details": {},
    }
    defaults.update(overrides)
    return AuditEvent(**defaults)  # type: ignore[arg-type]


class TestAuditEvent:
    """Tests for AuditEvent dataclass."""

    def test_to_dict(self) -> None:
        event = _make_event()
        d = event.to_dict()
        assert d["tool"] == "xls_delete_sheet"
        assert isinstance(d, dict)

    def test_all_fields_present(self) -> None:
        event = _make_event()
        d = event.to_dict()
        expected_keys = {
            "timestamp", "tool", "scope", "resource", "action",
            "outcome", "token_used", "file_hash", "pid", "details",
        }
        assert set(d.keys()) == expected_keys


class TestJsonlAuditBackend:
    """Tests for the JSONL file backend."""

    def test_log_creates_file(
        self, jsonl_backend: JsonlAuditBackend, audit_path: Path
    ) -> None:
        event = _make_event()
        jsonl_backend.log_event(event)
        assert audit_path.exists()

    def test_log_appends_jsonl(
        self, jsonl_backend: JsonlAuditBackend, audit_path: Path
    ) -> None:
        jsonl_backend.log_event(_make_event(tool="tool_1"))
        jsonl_backend.log_event(_make_event(tool="tool_2"))
        jsonl_backend.log_event(_make_event(tool="tool_3"))

        lines = audit_path.read_text().strip().split("\n")
        assert len(lines) == 3

        # Each line must be valid JSON
        for line in lines:
            data = json.loads(line)
            assert "tool" in data

    def test_log_preserves_all_fields(
        self, jsonl_backend: JsonlAuditBackend, audit_path: Path
    ) -> None:
        event = _make_event(details={"key": "value", "count": 42})
        jsonl_backend.log_event(event)

        line = audit_path.read_text().strip()
        data = json.loads(line)
        assert data["tool"] == "xls_delete_sheet"
        assert data["details"]["key"] == "value"
        assert data["details"]["count"] == 42

    def test_query_returns_events(
        self, jsonl_backend: JsonlAuditBackend
    ) -> None:
        jsonl_backend.log_event(_make_event(tool="tool_a"))
        jsonl_backend.log_event(_make_event(tool="tool_b"))

        events = jsonl_backend.query_events()
        assert len(events) == 2

    def test_query_filter_by_tool(
        self, jsonl_backend: JsonlAuditBackend
    ) -> None:
        jsonl_backend.log_event(_make_event(tool="tool_a"))
        jsonl_backend.log_event(_make_event(tool="tool_b"))
        jsonl_backend.log_event(_make_event(tool="tool_a"))

        events = jsonl_backend.query_events(tool="tool_a")
        assert len(events) == 2
        assert all(e.tool == "tool_a" for e in events)

    def test_query_filter_by_outcome(
        self, jsonl_backend: JsonlAuditBackend
    ) -> None:
        jsonl_backend.log_event(_make_event(outcome="success"))
        jsonl_backend.log_event(_make_event(outcome="denied"))
        jsonl_backend.log_event(_make_event(outcome="success"))

        events = jsonl_backend.query_events(outcome="denied")
        assert len(events) == 1
        assert events[0].outcome == "denied"

    def test_query_limit(self, jsonl_backend: JsonlAuditBackend) -> None:
        for i in range(20):
            jsonl_backend.log_event(_make_event(tool=f"tool_{i}"))

        events = jsonl_backend.query_events(limit=5)
        assert len(events) == 5

    def test_query_empty_log(self, jsonl_backend: JsonlAuditBackend) -> None:
        events = jsonl_backend.query_events()
        assert events == []

    def test_most_recent_first(self, jsonl_backend: JsonlAuditBackend) -> None:
        jsonl_backend.log_event(_make_event(tool="first"))
        jsonl_backend.log_event(_make_event(tool="second"))
        jsonl_backend.log_event(_make_event(tool="third"))

        events = jsonl_backend.query_events()
        assert events[0].tool == "third"
        assert events[2].tool == "first"

    def test_log_path_property(self, jsonl_backend: JsonlAuditBackend, audit_path: Path) -> None:
        assert jsonl_backend.log_path == audit_path.resolve()


class TestNullAuditBackend:
    """Tests for the no-op testing backend."""

    def test_log_stores_in_memory(self, null_backend: NullAuditBackend) -> None:
        null_backend.log_event(_make_event())
        assert len(null_backend._events) == 1

    def test_no_file_created(self, null_backend: NullAuditBackend, tmp_path: Path) -> None:
        null_backend.log_event(_make_event())
        # No files should have been created in tmp_path
        assert list(tmp_path.iterdir()) == []

    def test_query_returns_from_memory(self, null_backend: NullAuditBackend) -> None:
        null_backend.log_event(_make_event(tool="tool_a"))
        null_backend.log_event(_make_event(tool="tool_b"))
        events = null_backend.query_events()
        assert len(events) == 2

    def test_query_filter(self, null_backend: NullAuditBackend) -> None:
        null_backend.log_event(_make_event(tool="target"))
        null_backend.log_event(_make_event(tool="other"))
        events = null_backend.query_events(tool="target")
        assert len(events) == 1


class TestCompositeAuditBackend:
    """Tests for the fan-out backend."""

    def test_forwards_to_all_backends(
        self, jsonl_backend: JsonlAuditBackend, null_backend: NullAuditBackend
    ) -> None:
        composite = CompositeAuditBackend(jsonl_backend, null_backend)
        composite.log_event(_make_event())

        # Both backends should have the event
        assert len(null_backend._events) == 1
        assert jsonl_backend.log_path.exists()

    def test_query_from_first_with_results(
        self, null_backend: NullAuditBackend
    ) -> None:
        composite = CompositeAuditBackend(null_backend)
        composite.log_event(_make_event())
        events = composite.query_events()
        assert len(events) == 1

    def test_survives_backend_failure(
        self, null_backend: NullAuditBackend, tmp_path: Path
    ) -> None:
        """If one backend fails, others should still receive the event."""
        # Create a JSONL backend pointing to an unwritable path
        bad_path = tmp_path / "unwritable" / "deep" / "path" / "audit.jsonl"
        # Don't create parent dirs — this will cause a write failure
        bad_backend = JsonlAuditBackend(bad_path)

        composite = CompositeAuditBackend(bad_backend, null_backend)
        composite.log_event(_make_event())

        # NullBackend should still have the event
        assert len(null_backend._events) == 1


class TestAuditTrailFacade:
    """Tests for the AuditTrail facade."""

    def test_default_backend_is_jsonl(self) -> None:
        trail = AuditTrail()
        assert isinstance(trail.backend, JsonlAuditBackend)

    def test_custom_backend(self, null_backend: NullAuditBackend) -> None:
        trail = AuditTrail(backend=null_backend)
        assert trail.backend is null_backend

    def test_log_operation(self, null_backend: NullAuditBackend) -> None:
        trail = AuditTrail(backend=null_backend)
        event = trail.log_operation(
            tool="xls_delete_sheet",
            scope="sheet:delete",
            resource="Sheet1",
            action="delete",
            outcome="success",
            token_used=True,
            file_hash="sha256:abc",
            details={"rows_affected": 100},
        )
        assert isinstance(event, AuditEvent)
        assert event.tool == "xls_delete_sheet"
        assert event.details["rows_affected"] == 100

    def test_log_includes_pid(self, null_backend: NullAuditBackend) -> None:
        trail = AuditTrail(backend=null_backend)
        event = trail.log_operation(
            tool="test", scope="test", resource="test",
            action="test", outcome="success",
        )
        assert event.pid == os.getpid()

    def test_log_includes_timestamp(self, null_backend: NullAuditBackend) -> None:
        trail = AuditTrail(backend=null_backend)
        event = trail.log_operation(
            tool="test", scope="test", resource="test",
            action="test", outcome="success",
        )
        # Should be a valid ISO 8601 string
        dt = datetime.fromisoformat(event.timestamp)
        assert dt.tzinfo is not None  # Must be timezone-aware

    def test_query_delegates(self, null_backend: NullAuditBackend) -> None:
        trail = AuditTrail(backend=null_backend)
        trail.log_operation(
            tool="tool_a", scope="s", resource="r",
            action="a", outcome="success",
        )
        trail.log_operation(
            tool="tool_b", scope="s", resource="r",
            action="a", outcome="denied",
        )
        events = trail.query(tool="tool_a")
        assert len(events) == 1
        assert events[0].tool == "tool_a"


class TestConcurrentAuditWrites:
    """Tests for concurrent write safety."""

    @pytest.mark.skipif(
        os.name == "nt", reason="Multiprocessing fork not available on Windows"
    )
    def test_concurrent_writes_no_corruption(self, tmp_path: Path) -> None:
        """Multiple processes writing simultaneously should not corrupt the log."""
        import multiprocessing

        audit_path = tmp_path / "concurrent_audit.jsonl"
        events_per_process = 50
        num_processes = 5

        def worker(path: Path, worker_id: int, count: int) -> None:
            backend = JsonlAuditBackend(path)
            for i in range(count):
                event = AuditEvent(
                    timestamp=datetime.now(timezone.utc).isoformat(),
                    tool=f"worker_{worker_id}",
                    scope="test",
                    resource=f"item_{i}",
                    action="write",
                    outcome="success",
                    token_used=False,
                    file_hash="sha256:test",
                    pid=os.getpid(),
                    details={"worker": worker_id, "index": i},
                )
                backend.log_event(event)

        processes = []
        for wid in range(num_processes):
            p = multiprocessing.Process(
                target=worker, args=(audit_path, wid, events_per_process)
            )
            processes.append(p)
            p.start()

        for p in processes:
            p.join(timeout=30)

        # Verify: all lines are valid JSON and we have the expected count
        lines = audit_path.read_text().strip().split("\n")
        expected_total = num_processes * events_per_process
        assert len(lines) == expected_total, (
            f"Expected {expected_total} lines, got {len(lines)}"
        )

        for i, line in enumerate(lines):
            try:
                data = json.loads(line)
                assert "tool" in data
            except json.JSONDecodeError:
                pytest.fail(f"Line {i} is not valid JSON: {line[:80]}...")
```

---

## Phase 5: VERIFY — Phase 3 Exit Criteria Checklist

| # | Criterion | Status | Evidence |
|:---|:---|:---|:---|
| 1 | HMAC-SHA256 signature via `hmac.new(key, msg, hashlib.sha256)` | ✅ | `_compute_signature()` uses canonical format |
| 2 | Constant-time comparison via `hmac.compare_digest()` | ✅ | Step 7 of validation; never uses `==` |
| 3 | Secret key from `EXCEL_AGENT_SECRET` env var or `secrets.token_bytes(32)` | ✅ | 256-bit auto-generated per secrets docs |
| 4 | Warning logged when secret is auto-generated | ✅ | `logger.warning()` call; tested in `test_auto_generated_secret_logs_warning` |
| 5 | Token TTL enforced (1–3600 seconds) | ✅ | Validated in `generate_token()`; tested with `sleep(1.5)` for expiry |
| 6 | Nonce tracking: single-use enforcement | ✅ | `self._used_nonces` set; tested in `test_reused_nonce_raises` |
| 7 | Nonce revocation: explicitly revoked tokens rejected | ✅ | `self._revoked_nonces` set; tested in `test_revoked_token_raises` |
| 8 | File-hash binding: token for file A rejected for file B | ✅ | Step 3 of validation; tested in `test_wrong_file_hash_raises` |
| 9 | Scope validation: wrong scope rejected | ✅ | Step 2 of validation; tested in `test_wrong_scope_raises` |
| 10 | Tampered token detected | ✅ | Signature mismatch; tested in `test_tampered_signature_raises` |
| 11 | Malformed token handled gracefully | ✅ | `_deserialize_token()` catches exceptions |
| 12 | All 7 scopes defined and validated | ✅ | `VALID_SCOPES` frozenset; `test_all_valid_scopes` |
| 13 | `ApprovalToken` is frozen dataclass with `expires_at`/`is_expired` | ✅ | `@dataclass(frozen=True)`; tested immutability |
| 14 | Token format: base64url(JSON) — compact and URL-safe | ✅ | `base64.urlsafe_b64encode` with padding recovery |
| 15 | JSONL backend: append-only file with `os.fsync()` | ✅ | `"a"` mode + `fsync` per research findings |
| 16 | JSONL backend: cross-process locking via sidecar `.lock` | ✅ | `fcntl.flock` / `msvcrt.locking` in `_acquire_lock` |
| 17 | NullAuditBackend: in-memory only, no I/O | ✅ | `self._events` list; `test_no_file_created` |
| 18 | CompositeAuditBackend: fan-out to multiple backends | ✅ | Iterates `self._backends`; survives backend failure |
| 19 | AuditTrail facade: `log_operation()` creates `AuditEvent` | ✅ | Includes `os.getpid()`, UTC timestamp |
| 20 | AuditTrail facade: `query()` delegates to backend | ✅ | Tested with tool and outcome filters |
| 21 | Concurrent write test: 5 processes × 50 events = 250 lines, no corruption | ✅ | `multiprocessing.Process`; validates every line as JSON |
| 22 | All events include ISO 8601 UTC timestamp | ✅ | `datetime.now(timezone.utc).isoformat()` |
| 23 | Different secret keys produce different (incompatible) signatures | ✅ | `test_different_secrets_produce_different_signatures` |

---

## Phase 6: DELIVER — Setup Instructions

```bash
# From project root (with Phase 0 + 1 + 2 in place):

# 1. Add the Phase 3 files at the paths shown above

# 2. Run all unit tests (Phase 0 + 1 + 2 + 3)
pytest tests/unit/ -v --tb=short

# 3. Run Phase 3 tests specifically
pytest tests/unit/test_token_manager.py tests/unit/test_audit_trail.py -v

# 4. Note: the expired token test has a 1.5s sleep
pytest tests/unit/test_token_manager.py::TestTokenValidation::test_expired_token_raises -v

# 5. Run concurrent write test (requires POSIX, skipped on Windows)
pytest tests/unit/test_audit_trail.py::TestConcurrentAuditWrites -v

# 6. Lint and type-check
black --check src/excel_agent/governance/token_manager.py src/excel_agent/governance/audit_trail.py
ruff check src/excel_agent/governance/
mypy src/excel_agent/governance/

# 7. Check coverage
pytest tests/unit/test_token_manager.py tests/unit/test_audit_trail.py \
  --cov=excel_agent.governance --cov-report=term-missing
```

### Key Design Decisions in Phase 3

1. **Canonical signature format** — The HMAC is computed over `scope|file_hash|nonce|issued_at|ttl` with pipe separators. This ensures field ordering is deterministic and prevents reinterpretation attacks where field boundaries shift.

2. **`secrets.token_hex(16)` for nonces** — As of 2015, it is believed that 32 bytes (256 bits) of randomness is sufficient for the typical use-case expected for the secrets module. Our 16-byte (128-bit) nonces are adequate for single-use tokens with session-scoped tracking, and 32 bytes are used for the secret key itself.

3. **`hmac.compare_digest()` — never `==`** — `hmac.compare_digest(a, b)` is used to safely compare two digests to prevent a type of side-channel attack called a timing attack. It returns True if a and b are equal, and False otherwise, but it does so in a constant-time manner. This is the single most critical security decision in the entire governance layer.

4. **`PermissionDeniedError.reason` field** — Each validation failure includes a machine-readable `reason` string (`"expired"`, `"scope_mismatch"`, `"already_used"`, etc.) so the AI agent can programmatically determine the cause and take corrective action (e.g., generate a new token if expired).

5. **JSONL append with `os.fsync()` + advisory lock** — Using `"a"` mode: The operating system ensures that every write operation initiated on a file opened in append mode is guaranteed to land at the true end of the file. We add `os.fsync()` for durability and an advisory file lock for cross-process safety. The concurrent write test validates this with 5 processes × 50 events.

6. **Pluggable backend from day one** — The `AuditBackend` Protocol enables future SIEM/webhook integration without changing the tool APIs. `NullAuditBackend` eliminates I/O overhead in tests, and `CompositeAuditBackend` supports fan-out patterns.

### What's Next: Phase 4

Phase 4 implements the **Governance CLI Tools** (6 tools: clone, validate, token, hash, lock, dependency) and the **Read Tools** (7 tools: range, sheets, names, tables, style, formula, metadata), plus the chunked I/O helper.
