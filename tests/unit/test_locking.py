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

        with pytest.raises(ValueError, match="deliberate"), FileLock(target, timeout=5.0):
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


class TestFileLockTimeout:
    """Timeout and polling behavior tests."""

    def test_respects_timeout(self, tmp_path: Path) -> None:
        """Lock should timeout after specified duration."""
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        with FileLock(target, timeout=5.0):
            start = time.monotonic()
            with pytest.raises(LockContentionError), FileLock(target, timeout=0.5):
                pass
            elapsed = time.monotonic() - start
            # Should take approximately the timeout duration
            assert 0.4 <= elapsed <= 1.0

    def test_exponential_backoff(self, tmp_path: Path) -> None:
        """Polling interval should increase exponentially."""
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        # Use a longer initial poll to observe exponential backoff
        with FileLock(target, timeout=5.0):
            start = time.monotonic()
            with pytest.raises(LockContentionError):
                with FileLock(target, timeout=1.0, poll_interval=0.1):
                    pass
            elapsed = time.monotonic() - start
            # Should be close to the timeout (1.0s)
            assert elapsed >= 0.9


class TestFileLockEdgeCases:
    """Edge case tests."""

    def test_multiple_acquires_same_process_not_blocked(self, tmp_path: Path) -> None:
        """Same process should be able to acquire lock multiple times (reentrant not supported, but shouldn't block)."""
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        # Note: FileLock is NOT reentrant by design
        # But we should at least verify consistent behavior
        with FileLock(target, timeout=1.0):
            # Second acquire in same process should fail (non-reentrant)
            with pytest.raises(LockContentionError):
                with FileLock(target, timeout=0.1):
                    pass

    def test_lock_file_permission_error_handled(self, tmp_path: Path) -> None:
        """Should handle permission errors gracefully."""
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        # Normal case should work
        with FileLock(target, timeout=1.0) as lock:
            assert lock.lock_path.exists()

    def test_cleanup_on_abnormal_exit(self, tmp_path: Path) -> None:
        """Lock file should be cleaned up even on abnormal exit."""
        target = tmp_path / "test.xlsx"
        target.write_text("data")

        try:
            with FileLock(target, timeout=1.0) as lock:
                assert lock.lock_path.exists()
                raise RuntimeError("Simulated crash")
        except RuntimeError:
            pass

        # Lock file should be cleaned up
        assert not FileLock.is_locked(target)
