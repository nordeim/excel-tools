# Phase 4: IMPLEMENT — Phase 1: Core Foundation

Let me first conduct targeted research to ensure every technical detail — file locking semantics, openpyxl's Tokenizer API, hash streaming patterns, and context manager edge cases — is precisely correct.All research complete. Now I have comprehensive, validated technical details for every Phase 1 component. Let me synthesize the findings and generate the code.

---

## Research Summary: Critical Technical Findings

**File Locking:**
- If LOCK_NB is used and the lock cannot be acquired, an OSError will be raised and the exception will have an errno attribute set to EACCES or EAGAIN (depending on the operating system; for portability, check for both values).
- We can never close the file unless we want to release the lock.
- msvcrt.LK_NBLCK locks the specified bytes. Lock part of a file based on file descriptor fd from the C runtime. Raises OSError on failure. The locked region of the file extends from the current file position for nbytes bytes.
- The `tox-dev/filelock` reference implementation uses `msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)` and on exception checks `if exception.errno != EACCES`, closing the fd on failure.

**openpyxl Tokenizer:**
- openpyxl supports limited parsing of formulas embedded in cells. The openpyxl.formula package contains a Tokenizer class to break formulas into their constituent tokens.
- Tokens have three attributes: `value`, `type`, `subtype`. The OPERAND type with subtype RANGE identifies all cell/range references.

**openpyxl load_workbook:**
- The full signature is `load_workbook(filename, read_only=False, keep_vba=False, data_only=False, keep_links=True, rich_text=False)`
- data_only controls whether cells with formulae have either the formula (default) or the value stored the last time Excel read the sheet. keep_vba controls whether any Visual Basic elements are preserved or not (default). If they are preserved they are still not editable.

**openpyxl Translator:**
- The Translator class modifies a formula so it can be translated from one cell to another. If A1 has formula "=B1+C1", then translating it to cell A2 results in formula "=B2+C2". The algorithm relies on the formula tokenizer to identify the parts that need to change.

---

Now generating all Phase 1 code files:

## File 1: `src/excel_agent/core/__init__.py`

```python
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
    "ExcelAgent",
    "FileLock",
    "RangeSerializer",
    "CellCoordinate",
    "RangeCoordinate",
]


def __getattr__(name: str) -> object:
    if name == "ExcelAgent":
        from excel_agent.core.agent import ExcelAgent

        return ExcelAgent
    if name == "FileLock":
        from excel_agent.core.locking import FileLock

        return FileLock
    if name in ("RangeSerializer", "CellCoordinate", "RangeCoordinate"):
        import excel_agent.core.serializers as serializers

        return getattr(serializers, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
```

---

## File 2: `src/excel_agent/core/locking.py`

```python
"""
Cross-platform atomic file locking for excel-agent-tools.

Provides OS-level file locking to prevent concurrent agent access to the
same workbook. Uses fcntl.flock on POSIX and msvcrt.locking on Windows.

The lock is implemented via a sidecar .lock file adjacent to the target,
so we never modify the Excel file itself for locking purposes.

Key semantics:
    - Exclusive lock (only one holder at a time)
    - Non-blocking attempts with polling + exponential backoff
    - Timeout raises LockContentionError (exit code 3)
    - Lock file is cleaned up on release
    - Always releases on __exit__, even if body raises
"""

from __future__ import annotations

import errno
import logging
import os
import sys
import time
from pathlib import Path
from typing import IO

from excel_agent.utils.exceptions import LockContentionError

logger = logging.getLogger(__name__)

# Maximum column for exponential backoff poll interval
_MAX_POLL_INTERVAL: float = 1.0


class FileLock:
    """OS-level file lock with timeout, retry, and contention detection.

    Usage::

        with FileLock(Path("workbook.xlsx"), timeout=30.0):
            # Exclusive access to workbook.xlsx
            ...

    The lock is held on a sidecar file ``.<filename>.lock`` adjacent to the
    target. This avoids any modification to the Excel file itself.
    """

    def __init__(
        self,
        path: Path,
        *,
        timeout: float = 30.0,
        poll_interval: float = 0.1,
    ) -> None:
        """Initialize a file lock.

        Args:
            path: Path to the file to protect (the .lock sidecar is derived).
            timeout: Maximum seconds to wait for lock acquisition.
            poll_interval: Initial interval between non-blocking attempts.
                           Doubles each attempt up to _MAX_POLL_INTERVAL.
        """
        self._target_path = path.resolve()
        self._lock_path = self._target_path.parent / f".{self._target_path.name}.lock"
        self._timeout = timeout
        self._initial_poll = poll_interval
        self._lock_file: IO[str] | None = None

    @property
    def lock_path(self) -> Path:
        """Path to the sidecar lock file."""
        return self._lock_path

    def __enter__(self) -> FileLock:
        """Acquire exclusive lock with exponential backoff polling.

        Returns:
            self

        Raises:
            LockContentionError: If the lock cannot be acquired within timeout.
        """
        self._acquire()
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: object,
    ) -> None:
        """Release the lock. Always releases, even on exception."""
        self._release()

    def _acquire(self) -> None:
        """Internal: attempt to acquire the lock with timeout."""
        deadline = time.monotonic() + self._timeout
        poll = self._initial_poll

        while True:
            try:
                self._try_lock()
                logger.debug("Lock acquired: %s", self._lock_path)
                return
            except OSError:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise LockContentionError(
                        f"Could not acquire lock on {self._target_path} "
                        f"within {self._timeout}s. Another process may hold it.",
                        details={
                            "lock_file": str(self._lock_path),
                            "timeout": self._timeout,
                        },
                    )
                sleep_time = min(poll, remaining, _MAX_POLL_INTERVAL)
                time.sleep(sleep_time)
                # Exponential backoff, capped
                poll = min(poll * 2, _MAX_POLL_INTERVAL)

    def _try_lock(self) -> None:
        """Attempt a single non-blocking lock acquisition.

        Raises:
            OSError: If the lock is held by another process.
        """
        # Open (or create) the sidecar lock file
        lock_file = open(self._lock_path, "w")  # noqa: SIM115
        try:
            self._platform_lock(lock_file)
        except OSError:
            lock_file.close()
            raise
        # Write our PID for debugging
        lock_file.write(str(os.getpid()))
        lock_file.flush()
        self._lock_file = lock_file

    def _release(self) -> None:
        """Release the lock and clean up the sidecar file."""
        if self._lock_file is None:
            return
        try:
            self._platform_unlock(self._lock_file)
        except OSError:
            logger.warning("Failed to unlock %s", self._lock_path, exc_info=True)
        try:
            self._lock_file.close()
        except OSError:
            pass
        # Best-effort cleanup of the lock file
        try:
            self._lock_path.unlink(missing_ok=True)
        except OSError:
            pass
        self._lock_file = None
        logger.debug("Lock released: %s", self._lock_path)

    @staticmethod
    def is_locked(path: Path) -> bool:
        """Non-blocking check: is this file currently locked by any process?

        Attempts to acquire the lock momentarily, then immediately releases.
        Returns True if the file is locked, False otherwise.

        Args:
            path: Path to the file to check (not the lock file).
        """
        lock_path = path.resolve().parent / f".{path.name}.lock"
        try:
            lock_file = open(lock_path, "w")  # noqa: SIM115
        except OSError:
            return True  # Can't even open the lock file — treat as locked
        try:
            FileLock._platform_lock(lock_file)
            FileLock._platform_unlock(lock_file)
            lock_file.close()
            try:
                lock_path.unlink(missing_ok=True)
            except OSError:
                pass
            return False
        except OSError:
            lock_file.close()
            return True

    @staticmethod
    def _platform_lock(lock_file: IO[str]) -> None:
        """Acquire a non-blocking exclusive lock (platform-specific).

        Raises:
            OSError: If the lock is already held.
        """
        if sys.platform == "win32":
            _win32_lock(lock_file)
        else:
            _posix_lock(lock_file)

    @staticmethod
    def _platform_unlock(lock_file: IO[str]) -> None:
        """Release an exclusive lock (platform-specific)."""
        if sys.platform == "win32":
            _win32_unlock(lock_file)
        else:
            _posix_unlock(lock_file)


# ---------------------------------------------------------------------------
# Platform-specific implementations
# ---------------------------------------------------------------------------


def _posix_lock(lock_file: IO[str]) -> None:
    """Acquire exclusive non-blocking lock via fcntl.flock.

    Per Python docs: If LOCK_NB is used and the lock cannot be acquired,
    an OSError is raised with errno EACCES or EAGAIN.
    """
    import fcntl

    try:
        fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError as exc:
        # Portability: check for both EACCES and EAGAIN
        if exc.errno in (errno.EACCES, errno.EAGAIN):
            raise
        raise


def _posix_unlock(lock_file: IO[str]) -> None:
    """Release lock via fcntl.flock."""
    import fcntl

    fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)


def _win32_lock(lock_file: IO[str]) -> None:
    """Acquire exclusive non-blocking lock via msvcrt.locking.

    Uses LK_NBLCK which raises OSError immediately if lock is held.
    Locks 1 byte at position 0 (standard pattern from tox-dev/filelock).
    """
    import msvcrt

    lock_file.seek(0)
    try:
        msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError as exc:
        if exc.errno == errno.EACCES:
            raise
        raise


def _win32_unlock(lock_file: IO[str]) -> None:
    """Release lock via msvcrt.locking."""
    import msvcrt

    lock_file.seek(0)
    msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
```

---

## File 3: `src/excel_agent/core/serializers.py`

```python
"""
Unified range parsing for excel-agent-tools.

Converts any Excel reference format (A1, R1C1, Named Range, Table[Column])
into internal coordinates and back. This is the single source of truth for
all range resolution across the 53 tools.

Supported formats:
    - A1 notation: "A1", "A1:C10", "$A$1:$C$10", "Sheet1!A1:C10"
    - R1C1 notation: "R1C1", "R1C1:R10C3"
    - Named ranges: "SalesData" (requires workbook context)
    - Table references: "Table1[Sales]" (requires workbook context)
    - Full rows/columns: "A:A", "1:1"
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
)

from excel_agent.utils.exceptions import ValidationError

if TYPE_CHECKING:
    from openpyxl import Workbook


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CellCoordinate:
    """A single cell coordinate (1-indexed)."""

    row: int
    col: int

    def __post_init__(self) -> None:
        if self.row < 1:
            raise ValidationError(f"Row must be >= 1, got {self.row}")
        if self.col < 1:
            raise ValidationError(f"Column must be >= 1, got {self.col}")


@dataclass(frozen=True)
class RangeCoordinate:
    """A rectangular cell range (1-indexed, inclusive).

    Attributes:
        sheet: Sheet name, or None for the active/default sheet.
        min_row: Top row (1-indexed).
        min_col: Left column (1-indexed).
        max_row: Bottom row, or None for single-cell / full-column.
        max_col: Right column, or None for single-cell / full-row.
    """

    sheet: str | None
    min_row: int
    min_col: int
    max_row: int | None = None
    max_col: int | None = None

    @property
    def is_single_cell(self) -> bool:
        return self.max_row is None and self.max_col is None


# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

# Matches A1-style cell references with optional $ anchors
_CELL_RE = re.compile(
    r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$"
)

# Matches A1:B2 range (each side is a cell ref)
_RANGE_RE = re.compile(
    r"^\$?([A-Za-z]{1,3})\$?(\d{1,7}):\$?([A-Za-z]{1,3})\$?(\d{1,7})$"
)

# R1C1 single cell
_R1C1_CELL_RE = re.compile(
    r"^R(\d+)C(\d+)$", re.IGNORECASE
)

# R1C1 range
_R1C1_RANGE_RE = re.compile(
    r"^R(\d+)C(\d+):R(\d+)C(\d+)$", re.IGNORECASE
)

# Full column: A:A or A:C
_FULL_COL_RE = re.compile(
    r"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$"
)

# Full row: 1:1 or 1:10
_FULL_ROW_RE = re.compile(
    r"^(\d+):(\d+)$"
)

# Sheet prefix: 'Sheet Name'! or Sheet1!
_SHEET_PREFIX_RE = re.compile(
    r"^(?:'([^']+)'|([A-Za-z0-9_.\-]+))!(.+)$"
)

# Table reference: Table1, Table1[Column], Table1[#All]
_TABLE_REF_RE = re.compile(
    r"^([A-Za-z_][A-Za-z0-9_.]*)(?:\[([^\]]*)\])?$"
)


# ---------------------------------------------------------------------------
# Column helpers
# ---------------------------------------------------------------------------


def col_letter_to_number(letter: str) -> int:
    """Convert column letter(s) to 1-indexed number.

    Examples:
        'A' → 1, 'Z' → 26, 'AA' → 27, 'XFD' → 16384
    """
    return column_index_from_string(letter.upper())


def col_number_to_letter(number: int) -> str:
    """Convert 1-indexed column number to letter(s).

    Examples:
        1 → 'A', 26 → 'Z', 27 → 'AA', 16384 → 'XFD'
    """
    return get_column_letter(number)


# ---------------------------------------------------------------------------
# RangeSerializer
# ---------------------------------------------------------------------------


class RangeSerializer:
    """Parses any supported Excel reference format to RangeCoordinate.

    Args:
        workbook: Optional workbook context for resolving named ranges
                  and table references. If None, only A1/R1C1 formats
                  are supported.
    """

    def __init__(self, workbook: Workbook | None = None) -> None:
        self._wb = workbook

    def parse(
        self, range_str: str, *, default_sheet: str | None = None
    ) -> RangeCoordinate:
        """Parse any supported format to RangeCoordinate.

        Args:
            range_str: The reference string (A1, R1C1, named range, table ref).
            default_sheet: Sheet name to use if not specified in range_str.

        Returns:
            Parsed RangeCoordinate.

        Raises:
            ValidationError: If the input cannot be parsed.
        """
        original = range_str.strip()
        if not original:
            raise ValidationError("Empty range string")

        # Step 1: Extract sheet prefix if present
        sheet: str | None = default_sheet
        ref = original
        m = _SHEET_PREFIX_RE.match(ref)
        if m:
            sheet = m.group(1) or m.group(2)  # quoted or unquoted name
            ref = m.group(3)

        # Step 2: Try each format in order

        # A1 range (A1:C10)
        m = _RANGE_RE.match(ref)
        if m:
            return RangeCoordinate(
                sheet=sheet,
                min_row=int(m.group(2)),
                min_col=col_letter_to_number(m.group(1)),
                max_row=int(m.group(4)),
                max_col=col_letter_to_number(m.group(3)),
            )

        # A1 single cell
        m = _CELL_RE.match(ref)
        if m:
            return RangeCoordinate(
                sheet=sheet,
                min_row=int(m.group(2)),
                min_col=col_letter_to_number(m.group(1)),
            )

        # R1C1 range
        m = _R1C1_RANGE_RE.match(ref)
        if m:
            return RangeCoordinate(
                sheet=sheet,
                min_row=int(m.group(1)),
                min_col=int(m.group(2)),
                max_row=int(m.group(3)),
                max_col=int(m.group(4)),
            )

        # R1C1 single cell
        m = _R1C1_CELL_RE.match(ref)
        if m:
            return RangeCoordinate(
                sheet=sheet,
                min_row=int(m.group(1)),
                min_col=int(m.group(2)),
            )

        # Full column (A:C)
        m = _FULL_COL_RE.match(ref)
        if m:
            return RangeCoordinate(
                sheet=sheet,
                min_row=1,
                min_col=col_letter_to_number(m.group(1)),
                max_row=None,  # signals "entire column"
                max_col=col_letter_to_number(m.group(2)),
            )

        # Full row (1:10)
        m = _FULL_ROW_RE.match(ref)
        if m:
            return RangeCoordinate(
                sheet=sheet,
                min_row=int(m.group(1)),
                min_col=1,
                max_row=int(m.group(2)),
                max_col=None,  # signals "entire row"
            )

        # Named range (requires workbook)
        if self._wb is not None:
            coord = self._try_named_range(ref, sheet)
            if coord is not None:
                return coord

        # Table reference (requires workbook)
        if self._wb is not None:
            m = _TABLE_REF_RE.match(ref)
            if m:
                coord = self._try_table_ref(m.group(1), m.group(2), sheet)
                if coord is not None:
                    return coord

        raise ValidationError(
            f"Cannot parse range: {original!r}. "
            "Expected A1, R1C1, named range, or table reference.",
            details={"input": original},
        )

    def _try_named_range(
        self, name: str, fallback_sheet: str | None
    ) -> RangeCoordinate | None:
        """Attempt to resolve a named range from the workbook."""
        if self._wb is None:
            return None

        for defn in self._wb.defined_names.definedName:
            if defn.name.lower() == name.lower():
                # DefinedName.attr_text is like "Sheet1!$A$1:$B$5"
                destinations = list(defn.destinations)
                if not destinations:
                    return None
                dest_sheet, dest_range = destinations[0]
                # Recursively parse the resolved reference
                return self.parse(
                    dest_range,
                    default_sheet=dest_sheet or fallback_sheet,
                )
        return None

    def _try_table_ref(
        self,
        table_name: str,
        column_spec: str | None,
        fallback_sheet: str | None,
    ) -> RangeCoordinate | None:
        """Attempt to resolve a table reference from the workbook."""
        if self._wb is None:
            return None

        for ws in self._wb.worksheets:
            for table in ws.tables.values():
                if table.name.lower() == table_name.lower():
                    sheet_name = ws.title
                    ref = table.ref  # e.g., "A1:D100"
                    if column_spec is None:
                        # Whole table
                        return self.parse(ref, default_sheet=sheet_name)
                    # Column-specific: need to find the column index
                    for idx, col in enumerate(table.tableColumns):
                        if col.name.lower() == column_spec.lower():
                            parsed = self.parse(ref, default_sheet=sheet_name)
                            target_col = parsed.min_col + idx
                            return RangeCoordinate(
                                sheet=sheet_name,
                                min_row=parsed.min_row,
                                min_col=target_col,
                                max_row=parsed.max_row,
                                max_col=target_col,
                            )
                    return None
        return None

    @staticmethod
    def to_a1(coord: RangeCoordinate) -> str:
        """Convert RangeCoordinate to A1 notation string.

        Examples:
            RangeCoordinate(None, 1, 1, 10, 3) → "A1:C10"
            RangeCoordinate("Sheet1", 1, 1, None, None) → "Sheet1!A1"
        """
        prefix = ""
        if coord.sheet is not None:
            # Quote sheet names containing spaces or special characters
            if re.search(r"[^A-Za-z0-9_.]", coord.sheet):
                prefix = f"'{coord.sheet}'!"
            else:
                prefix = f"{coord.sheet}!"

        start = f"{col_number_to_letter(coord.min_col)}{coord.min_row}"

        if coord.max_row is None and coord.max_col is None:
            return f"{prefix}{start}"

        if coord.max_row is not None and coord.max_col is not None:
            end = f"{col_number_to_letter(coord.max_col)}{coord.max_row}"
            return f"{prefix}{start}:{end}"

        # Partial: full row or full column
        if coord.max_row is not None and coord.max_col is None:
            # Full row range
            return f"{prefix}{coord.min_row}:{coord.max_row}"
        # max_col is not None and max_row is None → full column range
        return (
            f"{prefix}{col_number_to_letter(coord.min_col)}"
            f":{col_number_to_letter(coord.max_col)}"
        )

    @staticmethod
    def to_r1c1(coord: RangeCoordinate) -> str:
        """Convert RangeCoordinate to R1C1 notation string.

        Examples:
            RangeCoordinate(None, 1, 1, 10, 3) → "R1C1:R10C3"
        """
        start = f"R{coord.min_row}C{coord.min_col}"
        if coord.max_row is None and coord.max_col is None:
            return start
        if coord.max_row is not None and coord.max_col is not None:
            return f"{start}:R{coord.max_row}C{coord.max_col}"
        return start
```

---

## File 4: `src/excel_agent/core/version_hash.py`

```python
"""
Geometry-aware workbook hashing for excel-agent-tools.

Produces SHA-256 hashes that capture workbook *structure and formulas*
but intentionally exclude cell values. This is because:
    - Values change on recalculation (they're volatile)
    - Formulas and structure are the "contract" (they're stable)
    - Concurrent modification detection needs to catch structural edits

Two hash types:
    - Geometry hash: sheet names + cell coordinates + formulas
    - File hash: raw bytes on disk (for concurrent modification detection)
"""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl import Workbook
    from openpyxl.worksheet.worksheet import Worksheet

# Read files in 64KB chunks for large file support
_FILE_CHUNK_SIZE = 65536


def compute_workbook_hash(workbook: Workbook) -> str:
    """Compute a SHA-256 hash of workbook geometry.

    Includes:
        - Sheet names in order
        - Sheet visibility states
        - Cell coordinates that contain formulas
        - Formula strings

    Excludes:
        - Cell values (they change on recalculation)
        - Cell styles (they're presentational, not structural)
        - Workbook metadata (author, etc.)

    Returns:
        String in format "sha256:<hex_digest>"
    """
    h = hashlib.sha256()

    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        h.update(f"SHEET:{sheet_name}".encode("utf-8"))
        h.update(f"STATE:{ws.sheet_state}".encode("utf-8"))
        _hash_sheet_geometry(ws, h)

    return f"sha256:{h.hexdigest()}"


def compute_sheet_hash(sheet: Worksheet) -> str:
    """Compute a SHA-256 hash of a single sheet's geometry.

    Returns:
        String in format "sha256:<hex_digest>"
    """
    h = hashlib.sha256()
    h.update(f"SHEET:{sheet.title}".encode("utf-8"))
    h.update(f"STATE:{sheet.sheet_state}".encode("utf-8"))
    _hash_sheet_geometry(sheet, h)
    return f"sha256:{h.hexdigest()}"


def compute_file_hash(path: Path) -> str:
    """Compute a SHA-256 hash of the raw file bytes on disk.

    This is used for concurrent modification detection: if the file
    changes on disk between our load and save, someone else modified it.

    Reads in 64KB chunks for memory efficiency on large files.

    Returns:
        String in format "sha256:<hex_digest>"
    """
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(_FILE_CHUNK_SIZE)
            if not chunk:
                break
            h.update(chunk)
    return f"sha256:{h.hexdigest()}"


def _hash_sheet_geometry(sheet: Worksheet, h: hashlib._Hash) -> None:
    """Hash the structural geometry of a single sheet.

    Iterates cells and includes coordinates + formulas for formula cells.
    Uses sorted iteration for deterministic ordering.
    """
    # Collect formula cells in deterministic order
    formula_cells: list[tuple[int, int, str]] = []

    for row in sheet.iter_rows():
        for cell in row:
            if cell.data_type == "f":
                # cell.value contains the formula string (e.g., "=SUM(A1:A10)")
                formula_str = str(cell.value) if cell.value is not None else ""
                formula_cells.append((cell.row, cell.column, formula_str))

    # Sort by (row, col) for deterministic order
    formula_cells.sort()

    for row, col, formula in formula_cells:
        entry = f"F:{row}:{col}:{formula}"
        h.update(entry.encode("utf-8"))

    # Also hash the dimensions (used range extent)
    if sheet.min_row is not None:
        h.update(
            f"DIM:{sheet.min_row}:{sheet.min_column}:"
            f"{sheet.max_row}:{sheet.max_column}".encode("utf-8")
        )
```

---

## File 5: `src/excel_agent/core/agent.py`

```python
"""
ExcelAgent: The central hub context manager for excel-agent-tools.

Integrates file locking, workbook loading, geometry hashing, and save
verification into a single, safe context manager.

Lifecycle:
    __enter__:
        1. Acquire FileLock (exclusive, with timeout)
        2. Load workbook via openpyxl (preserving formulas and VBA)
        3. Compute entry file hash (for concurrent modification detection)
        4. Compute geometry hash (for version reporting to agent)

    __exit__ (no exception, mode='rw'):
        1. Re-read file hash from disk
        2. If changed: raise ConcurrentModificationError (do NOT save)
        3. Save workbook
        4. Release lock

    __exit__ (with exception):
        1. Release lock WITHOUT saving (prevent partial/corrupt writes)
        2. Re-raise the exception

Usage::

    with ExcelAgent(Path("workbook.xlsx"), mode="rw") as agent:
        ws = agent.workbook.active
        ws["A1"] = "Hello"
        # Saved automatically on __exit__ with hash verification
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import TYPE_CHECKING

from openpyxl import load_workbook

from excel_agent.core.locking import FileLock
from excel_agent.core.version_hash import compute_file_hash, compute_workbook_hash
from excel_agent.utils.exceptions import (
    ConcurrentModificationError,
    ExcelFileNotFoundError,
    ValidationError,
)

if TYPE_CHECKING:
    from openpyxl import Workbook

logger = logging.getLogger(__name__)

_VALID_MODES = frozenset({"r", "rw"})
_VBA_EXTENSIONS = frozenset({".xlsm", ".xltm"})


class ExcelAgent:
    """Stateful context manager for safe, locked, hash-verified workbook manipulation.

    Args:
        path: Path to the Excel workbook (.xlsx, .xlsm, .xltx, .xltm).
        mode: "r" for read-only, "rw" for read-write (default).
        keep_vba: Preserve VBA projects in .xlsm files. Auto-detected from
                  extension if not explicitly set.
        lock_timeout: Max seconds to wait for lock (default: 30).
        data_only: If True, read cached values instead of formulas.
                   Default False — we ALWAYS preserve formulas.
    """

    def __init__(
        self,
        path: Path,
        *,
        mode: str = "rw",
        keep_vba: bool | None = None,
        lock_timeout: float = 30.0,
        data_only: bool = False,
    ) -> None:
        if mode not in _VALID_MODES:
            raise ValidationError(
                f"Invalid mode {mode!r}. Must be 'r' or 'rw'.",
                details={"mode": mode},
            )

        self._path = Path(path).resolve()
        self._mode = mode
        self._lock_timeout = lock_timeout
        self._data_only = data_only

        # Auto-detect VBA from extension if not explicitly set
        if keep_vba is None:
            self._keep_vba = self._path.suffix.lower() in _VBA_EXTENSIONS
        else:
            self._keep_vba = keep_vba

        # State (set during __enter__)
        self._lock: FileLock | None = None
        self._wb: Workbook | None = None
        self._entry_file_hash: str = ""
        self._geometry_hash: str = ""
        self._entered = False

    def __enter__(self) -> ExcelAgent:
        """Acquire lock, load workbook, compute hashes."""
        # Validate file exists
        if not self._path.exists():
            raise ExcelFileNotFoundError(
                f"Workbook not found: {self._path}",
                details={"path": str(self._path)},
            )
        if not self._path.is_file():
            raise ExcelFileNotFoundError(
                f"Path is not a file: {self._path}",
                details={"path": str(self._path)},
            )

        # Step 1: Acquire exclusive lock
        self._lock = FileLock(self._path, timeout=self._lock_timeout)
        self._lock.__enter__()

        try:
            # Step 2: Compute entry file hash (for concurrent modification detection)
            self._entry_file_hash = compute_file_hash(self._path)

            # Step 3: Load workbook
            self._wb = load_workbook(
                str(self._path),
                read_only=False,
                keep_vba=self._keep_vba,
                data_only=self._data_only,
                keep_links=True,
            )

            # Step 4: Compute geometry hash (for version reporting)
            self._geometry_hash = compute_workbook_hash(self._wb)

            self._entered = True
            logger.info(
                "ExcelAgent entered: %s (mode=%s, vba=%s, hash=%s)",
                self._path.name,
                self._mode,
                self._keep_vba,
                self._geometry_hash[:20] + "...",
            )
            return self

        except Exception:
            # If anything fails after lock acquisition, release the lock
            if self._lock is not None:
                self._lock.__exit__(None, None, None)
                self._lock = None
            raise

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: object,
    ) -> None:
        """Save (if rw + no exception), verify hash, release lock."""
        try:
            if exc_type is None and self._mode == "rw" and self._wb is not None:
                # Verify no concurrent modification before saving
                self.verify_no_concurrent_modification()
                # Save the workbook
                self._wb.save(str(self._path))
                logger.info("Workbook saved: %s", self._path.name)
            elif exc_type is not None:
                logger.warning(
                    "ExcelAgent exiting with exception — NOT saving: %s",
                    exc_type.__name__,
                )
        finally:
            # ALWAYS release the lock, even if save fails
            self._wb = None
            self._entered = False
            if self._lock is not None:
                self._lock.__exit__(None, None, None)
                self._lock = None

    @property
    def workbook(self) -> Workbook:
        """The openpyxl Workbook object.

        Raises:
            RuntimeError: If accessed outside the context manager.
        """
        if not self._entered or self._wb is None:
            raise RuntimeError(
                "ExcelAgent.workbook accessed outside context manager. "
                "Use 'with ExcelAgent(...) as agent:' to access the workbook."
            )
        return self._wb

    @property
    def path(self) -> Path:
        """The resolved workbook file path."""
        return self._path

    @property
    def version_hash(self) -> str:
        """The geometry hash computed at entry time.

        Returns:
            String in format "sha256:<hex_digest>"
        """
        return self._geometry_hash

    @property
    def file_hash(self) -> str:
        """The file byte hash computed at entry time."""
        return self._entry_file_hash

    def verify_no_concurrent_modification(self) -> None:
        """Re-read file bytes and compare hash to entry-time hash.

        Raises:
            ConcurrentModificationError: If the file was modified externally.
        """
        current_hash = compute_file_hash(self._path)
        if current_hash != self._entry_file_hash:
            raise ConcurrentModificationError(
                f"Workbook {self._path.name} was modified by another process "
                f"during this edit session. Refusing to save to prevent data loss.",
                details={
                    "path": str(self._path),
                    "entry_hash": self._entry_file_hash,
                    "current_hash": current_hash,
                },
            )
```

---

## File 6: `tests/unit/test_locking.py`

```python
"""Tests for cross-platform file locking."""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

import pytest

from excel_agent.core.locking import FileLock
from excel_agent.utils.exceptions import LockContentionError


class TestFileLockBasic:
    """Basic lock acquire/release tests."""

    def test_acquire_and_release(self, tmp_path: Path) -> None:
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        with FileLock(target, timeout=5.0) as lock:
            assert lock.lock_path.exists()
            # Lock file should contain our PID
            content = lock.lock_path.read_text()
            assert str(os.getpid()) in content

        # After release, lock file should be cleaned up
        assert not lock.lock_path.exists()

    def test_lock_path_derivation(self, tmp_path: Path) -> None:
        target = tmp_path / "workbook.xlsx"
        target.write_text("data")
        lock = FileLock(target)
        expected = tmp_path / ".workbook.xlsx.lock"
        assert lock.lock_path == expected

    def test_release_on_exception(self, tmp_path: Path) -> None:
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        with pytest.raises(ValueError, match="deliberate"):
            with FileLock(target, timeout=5.0):
                raise ValueError("deliberate")

        # Lock should be released even after exception
        assert not FileLock.is_locked(target)

    def test_is_locked_returns_false_when_free(self, tmp_path: Path) -> None:
        target = tmp_path / "test.xlsx"
        target.write_text("data")
        assert not FileLock.is_locked(target)

    def test_is_locked_returns_true_when_held(self, tmp_path: Path) -> None:
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        with FileLock(target, timeout=5.0):
            assert FileLock.is_locked(target)

        # After release
        assert not FileLock.is_locked(target)


class TestFileLockContention:
    """Lock contention and timeout tests."""

    def test_second_lock_raises_on_timeout(self, tmp_path: Path) -> None:
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        with FileLock(target, timeout=5.0):
            # Second lock attempt with short timeout should fail
            with pytest.raises(LockContentionError):
                with FileLock(target, timeout=0.3):
                    pass  # Should never reach here

    def test_timeout_error_has_details(self, tmp_path: Path) -> None:
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        with FileLock(target, timeout=5.0):
            with pytest.raises(LockContentionError) as exc_info:
                with FileLock(target, timeout=0.2):
                    pass

            assert exc_info.value.exit_code == 3
            assert "lock_file" in exc_info.value.details

    @pytest.mark.skipif(sys.platform == "win32", reason="Uses fork-based multiprocessing")
    def test_concurrent_processes(self, tmp_path: Path) -> None:
        """Two processes compete for the same lock; one should win."""
        import multiprocessing

        target = tmp_path / "test.xlsx"
        target.write_text("data")
        result_file = tmp_path / "result.txt"

        def worker(target_path: Path, result_path: Path, worker_id: int) -> None:
            try:
                with FileLock(Path(target_path), timeout=0.5):
                    time.sleep(1.0)  # Hold the lock
                    result_path.write_text(f"worker_{worker_id}_got_lock")
            except LockContentionError:
                pass  # Expected for the loser

        p1 = multiprocessing.Process(target=worker, args=(target, result_file, 1))
        p2_result = tmp_path / "result2.txt"
        p2 = multiprocessing.Process(target=worker, args=(target, p2_result, 2))

        p1.start()
        time.sleep(0.1)  # Give p1 a head start
        p2.start()

        p1.join(timeout=5)
        p2.join(timeout=5)

        # At least one should have gotten the lock
        got_lock = result_file.exists() or p2_result.exists()
        assert got_lock
```

---

## File 7: `tests/unit/test_serializers.py`

```python
"""Tests for unified range parsing (RangeSerializer)."""

from __future__ import annotations

import pytest

from excel_agent.core.serializers import (
    CellCoordinate,
    RangeCoordinate,
    RangeSerializer,
    col_letter_to_number,
    col_number_to_letter,
)
from excel_agent.utils.exceptions import ValidationError


class TestColConversions:
    """Tests for column letter ↔ number conversions."""

    @pytest.mark.parametrize(
        ("letter", "expected"),
        [("A", 1), ("Z", 26), ("AA", 27), ("AZ", 52), ("XFD", 16384)],
    )
    def test_letter_to_number(self, letter: str, expected: int) -> None:
        assert col_letter_to_number(letter) == expected

    @pytest.mark.parametrize(
        ("number", "expected"),
        [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (16384, "XFD")],
    )
    def test_number_to_letter(self, number: int, expected: str) -> None:
        assert col_number_to_letter(number) == expected

    @pytest.mark.parametrize("n", [1, 26, 27, 52, 256, 702, 16384])
    def test_roundtrip_number(self, n: int) -> None:
        assert col_letter_to_number(col_number_to_letter(n)) == n

    @pytest.mark.parametrize("letter", ["A", "Z", "AA", "AZ", "XFD"])
    def test_roundtrip_letter(self, letter: str) -> None:
        assert col_number_to_letter(col_letter_to_number(letter)) == letter


class TestCellCoordinate:
    """Tests for CellCoordinate dataclass."""

    def test_valid(self) -> None:
        c = CellCoordinate(row=1, col=1)
        assert c.row == 1 and c.col == 1

    def test_invalid_row(self) -> None:
        with pytest.raises(ValidationError):
            CellCoordinate(row=0, col=1)

    def test_invalid_col(self) -> None:
        with pytest.raises(ValidationError):
            CellCoordinate(row=1, col=0)


class TestRangeSerializerA1:
    """Tests for A1 notation parsing."""

    def setup_method(self) -> None:
        self.s = RangeSerializer()

    def test_single_cell(self) -> None:
        r = self.s.parse("A1")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1)
        assert r.is_single_cell

    def test_single_cell_absolute(self) -> None:
        r = self.s.parse("$A$1")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1)

    def test_range(self) -> None:
        r = self.s.parse("A1:C10")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)

    def test_range_absolute(self) -> None:
        r = self.s.parse("$A$1:$C$10")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)

    def test_with_sheet_unquoted(self) -> None:
        r = self.s.parse("Sheet1!A1:C10")
        assert r == RangeCoordinate(
            sheet="Sheet1", min_row=1, min_col=1, max_row=10, max_col=3
        )

    def test_with_sheet_quoted(self) -> None:
        r = self.s.parse("'Sheet Name'!A1")
        assert r == RangeCoordinate(sheet="Sheet Name", min_row=1, min_col=1)

    def test_multi_letter_column(self) -> None:
        r = self.s.parse("AA100")
        assert r == RangeCoordinate(sheet=None, min_row=100, min_col=27)

    def test_default_sheet(self) -> None:
        r = self.s.parse("A1", default_sheet="Data")
        assert r.sheet == "Data"


class TestRangeSerializerR1C1:
    """Tests for R1C1 notation parsing."""

    def setup_method(self) -> None:
        self.s = RangeSerializer()

    def test_single_cell(self) -> None:
        r = self.s.parse("R1C1")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1)

    def test_range(self) -> None:
        r = self.s.parse("R1C1:R10C3")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)

    def test_case_insensitive(self) -> None:
        r = self.s.parse("r5c10")
        assert r == RangeCoordinate(sheet=None, min_row=5, min_col=10)


class TestRangeSerializerFullRowCol:
    """Tests for full row and column references."""

    def setup_method(self) -> None:
        self.s = RangeSerializer()

    def test_full_column(self) -> None:
        r = self.s.parse("A:C")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=None, max_col=3)

    def test_full_row(self) -> None:
        r = self.s.parse("1:10")
        assert r == RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=None)


class TestRangeSerializerNamedRange:
    """Tests for named range resolution (requires workbook)."""

    def test_named_range_resolution(self, sample_workbook: Path) -> None:
        from openpyxl import load_workbook

        wb = load_workbook(str(sample_workbook))
        s = RangeSerializer(workbook=wb)
        r = s.parse("SalesData")
        assert r.sheet == "Sheet3"
        assert r.min_row == 1
        assert r.min_col == 1


class TestRangeSerializerInvalid:
    """Tests for error handling on invalid inputs."""

    def setup_method(self) -> None:
        self.s = RangeSerializer()

    def test_empty_string(self) -> None:
        with pytest.raises(ValidationError, match="Empty range"):
            self.s.parse("")

    def test_garbage(self) -> None:
        with pytest.raises(ValidationError, match="Cannot parse"):
            self.s.parse("not_a_valid_range!!!")


class TestToA1:
    """Tests for RangeCoordinate → A1 conversion."""

    def test_single_cell(self) -> None:
        coord = RangeCoordinate(sheet=None, min_row=1, min_col=1)
        assert RangeSerializer.to_a1(coord) == "A1"

    def test_range(self) -> None:
        coord = RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)
        assert RangeSerializer.to_a1(coord) == "A1:C10"

    def test_with_sheet(self) -> None:
        coord = RangeCoordinate(sheet="Sheet1", min_row=1, min_col=1, max_row=10, max_col=3)
        assert RangeSerializer.to_a1(coord) == "Sheet1!A1:C10"

    def test_with_quoted_sheet(self) -> None:
        coord = RangeCoordinate(sheet="My Sheet", min_row=1, min_col=1)
        assert RangeSerializer.to_a1(coord) == "'My Sheet'!A1"


class TestToR1C1:
    """Tests for RangeCoordinate → R1C1 conversion."""

    def test_single_cell(self) -> None:
        coord = RangeCoordinate(sheet=None, min_row=5, min_col=10)
        assert RangeSerializer.to_r1c1(coord) == "R5C10"

    def test_range(self) -> None:
        coord = RangeCoordinate(sheet=None, min_row=1, min_col=1, max_row=10, max_col=3)
        assert RangeSerializer.to_r1c1(coord) == "R1C1:R10C3"


class TestRoundtrip:
    """Tests for parse → to_a1 roundtrip fidelity."""

    @pytest.mark.parametrize(
        "input_str",
        ["A1", "A1:C10", "Z26", "AA1:AZ100", "A1:A1"],
    )
    def test_a1_roundtrip(self, input_str: str) -> None:
        s = RangeSerializer()
        coord = s.parse(input_str)
        output = s.to_a1(coord)
        # Parse the output again — should produce same coord
        coord2 = s.parse(output)
        assert coord == coord2
```

---

## File 8: `tests/unit/test_version_hash.py`

```python
"""Tests for geometry-aware workbook hashing."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from excel_agent.core.version_hash import (
    compute_file_hash,
    compute_sheet_hash,
    compute_workbook_hash,
)


class TestWorkbookHash:
    """Tests for compute_workbook_hash."""

    def test_identical_workbooks_same_hash(self, tmp_path: Path) -> None:
        """Two structurally identical workbooks should produce the same hash."""
        wb1 = Workbook()
        ws1 = wb1.active
        assert ws1 is not None
        ws1["A1"] = "=B1+1"
        ws1["B1"] = 10
        path1 = tmp_path / "wb1.xlsx"
        wb1.save(str(path1))

        wb2 = Workbook()
        ws2 = wb2.active
        assert ws2 is not None
        ws2["A1"] = "=B1+1"
        ws2["B1"] = 10
        path2 = tmp_path / "wb2.xlsx"
        wb2.save(str(path2))

        from openpyxl import load_workbook as lw

        h1 = compute_workbook_hash(lw(str(path1)))
        h2 = compute_workbook_hash(lw(str(path2)))
        assert h1 == h2

    def test_value_change_same_hash(self, tmp_path: Path) -> None:
        """Changing a cell VALUE should NOT change the geometry hash."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "=B1+1"
        ws["B1"] = 10

        h1 = compute_workbook_hash(wb)

        ws["B1"] = 9999  # Value change only

        h2 = compute_workbook_hash(wb)
        assert h1 == h2

    def test_formula_change_different_hash(self, tmp_path: Path) -> None:
        """Changing a cell FORMULA should change the geometry hash."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "=B1+1"

        h1 = compute_workbook_hash(wb)

        ws["A1"] = "=B1+2"

        h2 = compute_workbook_hash(wb)
        assert h1 != h2

    def test_sheet_rename_different_hash(self) -> None:
        """Renaming a sheet should change the hash."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Data"

        h1 = compute_workbook_hash(wb)

        ws.title = "NewName"

        h2 = compute_workbook_hash(wb)
        assert h1 != h2

    def test_sheet_add_different_hash(self) -> None:
        """Adding a sheet should change the hash."""
        wb = Workbook()
        h1 = compute_workbook_hash(wb)

        wb.create_sheet("NewSheet")

        h2 = compute_workbook_hash(wb)
        assert h1 != h2

    def test_sheet_remove_different_hash(self) -> None:
        """Removing a sheet should change the hash."""
        wb = Workbook()
        wb.create_sheet("Extra")
        h1 = compute_workbook_hash(wb)

        wb.remove(wb["Extra"])

        h2 = compute_workbook_hash(wb)
        assert h1 != h2

    def test_sheet_reorder_different_hash(self) -> None:
        """Reordering sheets should change the hash."""
        wb = Workbook()
        wb.create_sheet("Alpha")
        wb.create_sheet("Beta")
        h1 = compute_workbook_hash(wb)

        wb.move_sheet("Beta", offset=-1)

        h2 = compute_workbook_hash(wb)
        assert h1 != h2

    def test_hash_format(self) -> None:
        """Hash should be in 'sha256:...' format."""
        wb = Workbook()
        h = compute_workbook_hash(wb)
        assert h.startswith("sha256:")
        # SHA256 hex digest is 64 characters
        assert len(h) == 7 + 64  # "sha256:" + 64 hex chars

    def test_empty_workbook_produces_hash(self) -> None:
        """Even an empty workbook should produce a valid hash."""
        wb = Workbook()
        h = compute_workbook_hash(wb)
        assert h.startswith("sha256:")


class TestSheetHash:
    """Tests for compute_sheet_hash."""

    def test_returns_valid_hash(self) -> None:
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        h = compute_sheet_hash(ws)
        assert h.startswith("sha256:")


class TestFileHash:
    """Tests for compute_file_hash."""

    def test_file_hash_changes_on_modification(self, tmp_path: Path) -> None:
        f = tmp_path / "test.txt"
        f.write_text("hello")
        h1 = compute_file_hash(f)

        f.write_text("world")
        h2 = compute_file_hash(f)

        assert h1 != h2

    def test_same_content_same_hash(self, tmp_path: Path) -> None:
        f1 = tmp_path / "a.txt"
        f2 = tmp_path / "b.txt"
        f1.write_text("identical content")
        f2.write_text("identical content")

        assert compute_file_hash(f1) == compute_file_hash(f2)

    def test_hash_format(self, tmp_path: Path) -> None:
        f = tmp_path / "test.txt"
        f.write_text("data")
        h = compute_file_hash(f)
        assert h.startswith("sha256:")
        assert len(h) == 7 + 64
```

---

## File 9: `tests/unit/test_agent.py`

```python
"""Tests for the ExcelAgent context manager."""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from excel_agent.core.agent import ExcelAgent
from excel_agent.utils.exceptions import (
    ConcurrentModificationError,
    ExcelFileNotFoundError,
    ValidationError,
)


class TestExcelAgentBasic:
    """Basic lifecycle tests."""

    def test_load_and_read(self, sample_workbook: Path) -> None:
        """Agent can load a workbook and read data."""
        with ExcelAgent(sample_workbook, mode="r") as agent:
            ws = agent.workbook.active
            assert ws is not None
            assert ws.title == "Sheet1"
            assert agent.path == sample_workbook.resolve()

    def test_load_modify_save(self, sample_workbook: Path, tmp_path: Path) -> None:
        """Agent can modify and save a workbook."""
        work_copy = tmp_path / "work.xlsx"
        shutil.copy2(sample_workbook, work_copy)

        with ExcelAgent(work_copy, mode="rw") as agent:
            ws = agent.workbook.active
            assert ws is not None
            ws["Z1"] = "Modified by agent"

        # Verify the modification persisted
        wb = load_workbook(str(work_copy))
        assert wb.active is not None
        assert wb.active["Z1"].value == "Modified by agent"

    def test_read_only_does_not_save(self, sample_workbook: Path, tmp_path: Path) -> None:
        """Read-only mode does NOT save, even if workbook is modified in memory."""
        work_copy = tmp_path / "readonly.xlsx"
        shutil.copy2(sample_workbook, work_copy)
        original_size = work_copy.stat().st_size

        from excel_agent.core.version_hash import compute_file_hash

        original_hash = compute_file_hash(work_copy)

        with ExcelAgent(work_copy, mode="r") as agent:
            ws = agent.workbook.active
            assert ws is not None
            ws["Z99"] = "This should NOT be saved"

        # File should be unchanged
        assert compute_file_hash(work_copy) == original_hash

    def test_version_hash_available(self, sample_workbook: Path) -> None:
        with ExcelAgent(sample_workbook, mode="r") as agent:
            h = agent.version_hash
            assert h.startswith("sha256:")
            assert len(h) == 7 + 64


class TestExcelAgentErrorHandling:
    """Error handling and edge case tests."""

    def test_file_not_found(self, tmp_path: Path) -> None:
        with pytest.raises(ExcelFileNotFoundError):
            with ExcelAgent(tmp_path / "nonexistent.xlsx", mode="r"):
                pass

    def test_invalid_mode(self, sample_workbook: Path) -> None:
        with pytest.raises(ValidationError, match="Invalid mode"):
            ExcelAgent(sample_workbook, mode="x")

    def test_workbook_access_outside_context(self, sample_workbook: Path) -> None:
        agent = ExcelAgent(sample_workbook, mode="r")
        with pytest.raises(RuntimeError, match="outside context"):
            _ = agent.workbook

    def test_exception_in_body_releases_lock(
        self, sample_workbook: Path, tmp_path: Path
    ) -> None:
        """Exception in context body → lock released, file NOT modified."""
        work_copy = tmp_path / "exception_test.xlsx"
        shutil.copy2(sample_workbook, work_copy)

        from excel_agent.core.locking import FileLock
        from excel_agent.core.version_hash import compute_file_hash

        original_hash = compute_file_hash(work_copy)

        with pytest.raises(ValueError, match="deliberate"):
            with ExcelAgent(work_copy, mode="rw") as agent:
                ws = agent.workbook.active
                assert ws is not None
                ws["Z1"] = "Should NOT be saved"
                raise ValueError("deliberate error")

        # File should be unchanged (not saved)
        assert compute_file_hash(work_copy) == original_hash
        # Lock should be released
        assert not FileLock.is_locked(work_copy)


class TestExcelAgentConcurrency:
    """Concurrent modification detection tests."""

    def test_concurrent_modification_detected(
        self, sample_workbook: Path, tmp_path: Path
    ) -> None:
        """External file modification during session raises error."""
        work_copy = tmp_path / "concurrent.xlsx"
        shutil.copy2(sample_workbook, work_copy)

        with pytest.raises(ConcurrentModificationError):
            with ExcelAgent(work_copy, mode="rw") as agent:
                # Simulate another process modifying the file
                # Write directly to bypass lock (simulating a bug or race condition)
                wb2 = load_workbook(str(work_copy))
                ws2 = wb2.active
                assert ws2 is not None
                ws2["A1"] = "EXTERNAL MODIFICATION"
                wb2.save(str(work_copy))
                # Now __exit__ should detect the modification


class TestExcelAgentVBA:
    """VBA preservation tests."""

    def test_auto_detect_vba_from_extension(self, tmp_path: Path) -> None:
        """Agent should auto-set keep_vba=True for .xlsm files."""
        # Create a basic .xlsx and rename to .xlsm (won't have real VBA,
        # but tests the auto-detection logic)
        wb = Workbook()
        xlsm_path = tmp_path / "test.xlsm"
        wb.save(str(xlsm_path))

        agent = ExcelAgent(xlsm_path, mode="r")
        assert agent._keep_vba is True

    def test_xlsx_does_not_auto_enable_vba(self, sample_workbook: Path) -> None:
        agent = ExcelAgent(sample_workbook, mode="r")
        assert agent._keep_vba is False

    def test_explicit_keep_vba_override(self, sample_workbook: Path) -> None:
        agent = ExcelAgent(sample_workbook, mode="r", keep_vba=True)
        assert agent._keep_vba is True
```

---

## File 10: `tests/property/test_range_serializer.py`

```python
"""Property-based tests for RangeSerializer using Hypothesis.

Tests the fundamental invariant: parse(to_a1(parse(x))) == parse(x)
for randomly generated valid A1 reference strings.
"""

from __future__ import annotations

import hypothesis.strategies as st
from hypothesis import given, settings

from excel_agent.core.serializers import RangeCoordinate, RangeSerializer, col_number_to_letter


# ---------------------------------------------------------------------------
# Strategies for generating valid Excel references
# ---------------------------------------------------------------------------

# Column numbers: 1 to 16384 (XFD)
col_numbers = st.integers(min_value=1, max_value=16384)

# Row numbers: 1 to 1048576
row_numbers = st.integers(min_value=1, max_value=1048576)


@st.composite
def single_cell_a1(draw: st.DrawFn) -> str:
    """Generate a random valid A1 single-cell reference."""
    col = draw(col_numbers)
    row = draw(row_numbers)
    return f"{col_number_to_letter(col)}{row}"


@st.composite
def range_a1(draw: st.DrawFn) -> str:
    """Generate a random valid A1 range reference."""
    col1 = draw(st.integers(min_value=1, max_value=16000))
    col2 = draw(st.integers(min_value=col1, max_value=min(col1 + 100, 16384)))
    row1 = draw(st.integers(min_value=1, max_value=1000000))
    row2 = draw(st.integers(min_value=row1, max_value=min(row1 + 100, 1048576)))
    return f"{col_number_to_letter(col1)}{row1}:{col_number_to_letter(col2)}{row2}"


# ---------------------------------------------------------------------------
# Property tests
# ---------------------------------------------------------------------------

serializer = RangeSerializer()


class TestColumnRoundtrip:
    """Property: col_number_to_letter(col_letter_to_number(x)) == x"""

    @given(n=col_numbers)
    @settings(max_examples=300)
    def test_number_roundtrip(self, n: int) -> None:
        from excel_agent.core.serializers import col_letter_to_number

        letter = col_number_to_letter(n)
        assert col_letter_to_number(letter) == n


class TestParseToA1Roundtrip:
    """Property: parse(to_a1(parse(s))) produces the same RangeCoordinate."""

    @given(cell_str=single_cell_a1())
    @settings(max_examples=200)
    def test_single_cell_roundtrip(self, cell_str: str) -> None:
        coord1 = serializer.parse(cell_str)
        a1_str = RangeSerializer.to_a1(coord1)
        coord2 = serializer.parse(a1_str)
        assert coord1 == coord2

    @given(range_str=range_a1())
    @settings(max_examples=200)
    def test_range_roundtrip(self, range_str: str) -> None:
        coord1 = serializer.parse(range_str)
        a1_str = RangeSerializer.to_a1(coord1)
        coord2 = serializer.parse(a1_str)
        assert coord1 == coord2


class TestCoordinateInvariants:
    """Property: parsed coordinates have valid bounds."""

    @given(cell_str=single_cell_a1())
    @settings(max_examples=200)
    def test_single_cell_invariants(self, cell_str: str) -> None:
        coord = serializer.parse(cell_str)
        assert coord.min_row >= 1
        assert coord.min_col >= 1
        assert coord.max_row is None
        assert coord.max_col is None
        assert coord.is_single_cell

    @given(range_str=range_a1())
    @settings(max_examples=200)
    def test_range_invariants(self, range_str: str) -> None:
        coord = serializer.parse(range_str)
        assert coord.min_row >= 1
        assert coord.min_col >= 1
        assert coord.max_row is not None
        assert coord.max_col is not None
        assert coord.max_row >= coord.min_row
        assert coord.max_col >= coord.min_col
        assert not coord.is_single_cell
```

---

## File 11: `tests/property/__init__.py`

```python
"""Property-based tests using Hypothesis."""
```

---

## File 12: `tests/integration/__init__.py`

```python
"""Integration tests (subprocess-based, simulating AI agent tool calls)."""
```

---

## File 13: `tests/performance/__init__.py`

```python
"""Performance benchmark tests."""
```

---

## Phase 5: VERIFY — Phase 1 Exit Criteria Checklist

| # | Criterion | Status | Evidence |
|:---|:---|:---|:---|
| 1 | `locking.py` — Unix `fcntl.flock(LOCK_EX \| LOCK_NB)` with exponential backoff | ✅ | Checks both `EACCES` and `EAGAIN` per Python docs |
| 2 | `locking.py` — Windows `msvcrt.locking(LK_NBLCK, 1)` | ✅ | Follows `tox-dev/filelock` reference pattern |
| 3 | `locking.py` — Sidecar `.lock` file (never modifies the Excel file) | ✅ | Lock path: `.{filename}.lock` |
| 4 | `locking.py` — `LockContentionError` with exit code 3 | ✅ | Raises with `details` dict |
| 5 | `locking.py` — `is_locked()` non-blocking check | ✅ | Acquires and immediately releases |
| 6 | `serializers.py` — A1 notation with optional `$` anchors | ✅ | Regex strips absolute markers |
| 7 | `serializers.py` — R1C1 notation (case insensitive) | ✅ | `re.IGNORECASE` flag |
| 8 | `serializers.py` — Named range resolution via workbook | ✅ | Uses `workbook.defined_names` |
| 9 | `serializers.py` — Table reference resolution | ✅ | Iterates `ws.tables.values()` |
| 10 | `serializers.py` — Full row/column (`A:A`, `1:1`) | ✅ | Separate regex patterns |
| 11 | `serializers.py` — Quoted sheet names (`'Sheet Name'!A1`) | ✅ | Sheet prefix regex |
| 12 | `serializers.py` — `to_a1()` and `to_r1c1()` reverse conversion | ✅ | Roundtrip tested |
| 13 | `version_hash.py` — Hashes formulas, excludes values | ✅ | Only processes `cell.data_type == "f"` |
| 14 | `version_hash.py` — SHA-256 with `sha256:` prefix | ✅ | Format: `sha256:` + 64 hex chars |
| 15 | `version_hash.py` — File hash for concurrent modification | ✅ | 64KB chunked reads |
| 16 | `agent.py` — Lock → Load → Hash lifecycle | ✅ | 3-step `__enter__` |
| 17 | `agent.py` — Verify → Save → Release on clean exit | ✅ | Hash comparison before save |
| 18 | `agent.py` — Release WITHOUT save on exception | ✅ | `finally` block in `__exit__` |
| 19 | `agent.py` — `ConcurrentModificationError` on external edit | ✅ | Compares file hashes |
| 20 | `agent.py` — Auto-detect `keep_vba` from `.xlsm` extension | ✅ | `_VBA_EXTENSIONS` set |
| 21 | `agent.py` — `load_workbook(data_only=False)` preserves formulas | ✅ | Explicit `data_only=False` default |
| 22 | Hypothesis property tests — 200+ examples, roundtrip fidelity | ✅ | 3 property classes |
| 23 | All tests use `tmp_path` for isolation | ✅ | No shared state |
| 24 | No mypy errors anticipated in strict mode | ✅ | All type annotations present |

---

## Phase 6: DELIVER — Setup Instructions

```bash
# From the project root (with Phase 0 already in place):

# 1. Create Phase 1 files at the paths shown above
#    (replace the placeholder core/__init__.py from Phase 0)

# 2. Run Phase 0 + Phase 1 unit tests
pytest tests/unit/test_exit_codes.py \
       tests/unit/test_json_io.py \
       tests/unit/test_exceptions.py \
       tests/unit/test_locking.py \
       tests/unit/test_serializers.py \
       tests/unit/test_version_hash.py \
       tests/unit/test_agent.py \
       tests/property/test_range_serializer.py \
       -v --tb=short

# 3. Run linters
black --check src/excel_agent/core/ tests/
ruff check src/excel_agent/core/
mypy src/excel_agent/core/

# 4. Check coverage
pytest tests/unit/test_locking.py tests/unit/test_serializers.py \
       tests/unit/test_version_hash.py tests/unit/test_agent.py \
       --cov=excel_agent.core --cov-report=term-missing
```

### Key Design Decisions in Phase 1

1. **Sidecar lock file pattern** — Locking directly on the `.xlsx` file would require opening it in a mode compatible with both locking and openpyxl's ZIP reader. Using a sidecar `.{filename}.lock` file avoids this entirely and follows the `tox-dev/filelock` and `portalocker` reference implementations.

2. **Exponential backoff** — The polling interval doubles each attempt (0.1s → 0.2s → 0.4s → 0.8s → 1.0s cap). This balances responsiveness (initial fast retry) with resource efficiency (backs off under contention).

3. **Geometry hash excludes values** — Cell values change on every recalculation. Formula strings and sheet structure are the "structural contract." This ensures `version_hash` is stable across recalcs but changes on structural mutations.

4. **Two-hash strategy in `ExcelAgent`** — `file_hash` (raw bytes) detects *any* external modification for concurrent modification protection. `geometry_hash` (formula structure) is the version identifier reported to the agent. Different purposes, different algorithms.

5. **`Translator` integration prep** — The openpyxl `Translator` class is used in later phases (Phase 6: `xls_copy_formula_down.py`) for translating formulas across cells. The `serializers.py` module provides the coordinate system that feeds into it.

### What's Next: Phase 2

Phase 2 implements the **DependencyTracker** (formula dependency graph) and **JSON Schema** validation infrastructure — the most safety-critical component in the entire project.
