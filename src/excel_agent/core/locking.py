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
                poll = min(poll * 2, _MAX_POLL_INTERVAL)

    def _try_lock(self) -> None:
        """Attempt a single non-blocking lock acquisition.

        Raises:
            OSError: If the lock is held by another process.
        """
        lock_file = open(self._lock_path, "w")  # noqa: SIM115
        try:
            self._platform_lock(lock_file)
        except OSError:
            lock_file.close()
            raise
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
            return True
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


def _posix_lock(lock_file: IO[str]) -> None:
    """Acquire exclusive non-blocking lock via fcntl.flock."""
    import fcntl

    try:
        fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError as exc:
        if exc.errno in (errno.EACCES, errno.EAGAIN):
            raise
        raise


def _posix_unlock(lock_file: IO[str]) -> None:
    """Release lock via fcntl.flock."""
    import fcntl

    fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)


def _win32_lock(lock_file: IO[str]) -> None:
    """Acquire exclusive non-blocking lock via msvcrt.locking."""
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
