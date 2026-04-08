"""
Governance components for excel-agent-tools.

This package contains the safety and compliance infrastructure:
- ApprovalTokenManager: HMAC-SHA256 scoped approval tokens
- AuditTrail: Pluggable audit logging for destructive operations
"""

from __future__ import annotations

__all__ = [
    "ApprovalTokenManager",
    "AuditEvent",
    "AuditTrail",
]


def __getattr__(name: str) -> object:
    if name == "ApprovalTokenManager":
        from excel_agent.governance.token_manager import ApprovalTokenManager

        return ApprovalTokenManager
    if name == "AuditTrail":
        from excel_agent.governance.audit_trail import AuditTrail

        return AuditTrail
    if name == "AuditEvent":
        from excel_agent.governance.audit_trail import AuditEvent

        return AuditEvent
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
