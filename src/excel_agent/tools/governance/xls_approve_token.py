"""xls_approve_token: Generate scoped HMAC-SHA256 approval tokens."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from excel_agent.governance.token_manager import ApprovalTokenManager
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import create_parser, validate_input_path
from excel_agent.utils.json_io import build_response


def _run() -> dict:
    parser = create_parser("Generate a scoped HMAC-SHA256 approval token.")
    parser.add_argument(
        "--scope", type=str, required=True, help="Token scope (e.g., sheet:delete)"
    )
    parser.add_argument("--file", type=str, required=True, help="Target workbook path")
    parser.add_argument("--ttl", type=int, default=300, help="Token TTL in seconds (default: 300)")
    args = parser.parse_args()

    path = validate_input_path(args.file)

    mgr = ApprovalTokenManager()
    token = mgr.generate_token(args.scope, path, ttl_seconds=args.ttl)
    expires_at = datetime.now(timezone.utc).timestamp() + args.ttl

    return build_response(
        "success",
        {
            "token": token,
            "scope": args.scope,
            "ttl_seconds": args.ttl,
            "expires_at": datetime.fromtimestamp(expires_at, tz=timezone.utc).isoformat(),
        },
        workbook_version="",
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
