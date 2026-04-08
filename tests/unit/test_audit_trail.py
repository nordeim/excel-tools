"""Tests for AuditTrail - pluggable audit logging."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_agent.governance.audit_trail import (
    AuditEvent,
    AuditTrail,
    JsonlAuditBackend,
)


class TestAuditEvent:
    """Tests for AuditEvent data structure."""

    def test_event_creation(self) -> None:
        """Should create event with current timestamp."""
        event = AuditEvent.create(
            tool="xls-delete-sheet",
            scope="sheet:delete",
            target_file=Path("test.xlsx"),
            file_version_hash="sha256:abc123",
            actor_nonce="nonce456",
            operation_details={"sheet_name": "Data"},
            impact={"sheets_deleted": 1},
        )

        assert event.tool == "xls-delete-sheet"
        assert event.scope == "sheet:delete"
        assert event.file_version_hash == "sha256:abc123"
        assert event.actor_nonce == "nonce456"
        assert event.operation_details["sheet_name"] == "Data"
        assert event.impact["sheets_deleted"] == 1
        assert event.success is True
        assert event.exit_code == 0
        assert event.timestamp  # Should have timestamp

    def test_event_to_dict(self) -> None:
        """Should convert to dictionary."""
        event = AuditEvent.create(
            tool="xls-delete-sheet",
            scope="sheet:delete",
            target_file=Path("test.xlsx"),
            file_version_hash="sha256:abc123",
            actor_nonce="nonce456",
        )

        d = event.to_dict()
        assert d["tool"] == "xls-delete-sheet"
        assert d["scope"] == "sheet:delete"
        assert d["success"] is True

    def test_event_to_json(self) -> None:
        """Should serialize to JSON."""
        event = AuditEvent.create(
            tool="xls-delete-sheet",
            scope="sheet:delete",
            target_file=Path("test.xlsx"),
            file_version_hash="sha256:abc123",
            actor_nonce="nonce456",
        )

        json_str = event.to_json()
        # Should be valid JSON
        data = json.loads(json_str)
        assert data["tool"] == "xls-delete-sheet"


class TestJsonlAuditBackend:
    """Tests for JSONL audit backend."""

    def test_log_event_creates_file(self, tmp_path: Path) -> None:
        """Logging should create JSONL file."""
        audit_file = tmp_path / ".test_audit.jsonl"
        backend = JsonlAuditBackend(audit_file)

        event = AuditEvent.create(
            tool="xls-delete-sheet",
            scope="sheet:delete",
            target_file=Path("test.xlsx"),
            file_version_hash="sha256:abc123",
            actor_nonce="nonce456",
        )

        backend.log_event(event)

        assert audit_file.exists()

    def test_log_event_appends(self, tmp_path: Path) -> None:
        """Multiple events should append to file."""
        audit_file = tmp_path / ".test_audit.jsonl"
        backend = JsonlAuditBackend(audit_file)

        for i in range(3):
            event = AuditEvent.create(
                tool=f"tool-{i}",
                scope="sheet:delete",
                target_file=Path(f"test{i}.xlsx"),
                file_version_hash=f"sha256:hash{i}",
                actor_nonce=f"nonce{i}",
            )
            backend.log_event(event)

        lines = audit_file.read_text().strip().split("\n")
        assert len(lines) == 3

    def test_log_event_is_jsonl(self, tmp_path: Path) -> None:
        """Each line should be valid JSON."""
        audit_file = tmp_path / ".test_audit.jsonl"
        backend = JsonlAuditBackend(audit_file)

        event = AuditEvent.create(
            tool="xls-delete-sheet",
            scope="sheet:delete",
            target_file=Path("test.xlsx"),
            file_version_hash="sha256:abc123",
            actor_nonce="nonce456",
        )

        backend.log_event(event)

        lines = audit_file.read_text().strip().split("\n")
        for line in lines:
            data = json.loads(line)
            assert "tool" in data
            assert "timestamp" in data


class TestAuditTrail:
    """Tests for AuditTrail high-level interface."""

    def test_log_creates_event(self, tmp_path: Path) -> None:
        """log() should create and log an event."""
        audit_file = tmp_path / ".test_audit.jsonl"
        audit = AuditTrail(backend=JsonlAuditBackend(audit_file))

        audit.log(
            tool="xls-delete-sheet",
            scope="sheet:delete",
            target_file=Path("test.xlsx"),
            file_version_hash="sha256:abc123",
            actor_nonce="nonce456",
            operation_details={"sheet_name": "Data"},
            impact={"sheets_deleted": 1},
        )

        assert audit_file.exists()

        lines = audit_file.read_text().strip().split("\n")
        data = json.loads(lines[0])
        assert data["tool"] == "xls-delete-sheet"
        assert data["operation_details"]["sheet_name"] == "Data"

    def test_read_events_returns_list(self, tmp_path: Path) -> None:
        """read_events() should return list of AuditEvent."""
        audit_file = tmp_path / ".test_audit.jsonl"
        audit = AuditTrail(backend=JsonlAuditBackend(audit_file))

        for i in range(3):
            audit.log(
                tool=f"tool-{i}",
                scope="sheet:delete",
                target_file=Path(f"test{i}.xlsx"),
                file_version_hash=f"sha256:hash{i}",
                actor_nonce=f"nonce{i}",
            )

        events = audit.read_events(audit_file)

        assert len(events) == 3
        assert events[0].tool == "tool-2"  # Newest first
        assert events[2].tool == "tool-0"

    def test_read_events_limit(self, tmp_path: Path) -> None:
        """read_events(limit=N) should return N newest events."""
        audit_file = tmp_path / ".test_audit.jsonl"
        audit = AuditTrail(backend=JsonlAuditBackend(audit_file))

        for i in range(5):
            audit.log(
                tool=f"tool-{i}",
                scope="sheet:delete",
                target_file=Path(f"test{i}.xlsx"),
                file_version_hash=f"sha256:hash{i}",
                actor_nonce=f"nonce{i}",
            )

        events = audit.read_events(audit_file, limit=2)

        assert len(events) == 2
        assert events[0].tool == "tool-4"  # Newest
        assert events[1].tool == "tool-3"

    def test_read_events_empty_file(self, tmp_path: Path) -> None:
        """read_events() should return empty list for non-existent file."""
        audit_file = tmp_path / ".nonexistent_audit.jsonl"

        audit = AuditTrail()
        events = audit.read_events(audit_file)

        assert events == []

    def test_default_backend_uses_jsonl(self, tmp_path: Path) -> None:
        """Default backend should be JsonlAuditBackend."""
        # Change to temp dir for default file location
        import os

        old_cwd = os.getcwd()
        os.chdir(str(tmp_path))

        try:
            audit = AuditTrail()  # No backend specified
            audit.log(
                tool="xls-delete-sheet",
                scope="sheet:delete",
                target_file=Path("test.xlsx"),
                file_version_hash="sha256:abc123",
                actor_nonce="nonce456",
            )

            assert (tmp_path / ".excel_agent_audit.jsonl").exists()
        finally:
            os.chdir(old_cwd)
