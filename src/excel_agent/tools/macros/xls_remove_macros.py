"""xls_remove_macros: Strip VBA macros from workbook (IRREVERSIBLE, double-token).

Removes all VBA macro code from a workbook by:
1. Loading the workbook with openpyxl
2. Removing all VBA project data
3. Saving as a new .xlsx file

This operation:
- Is IRREVERSIBLE (macros cannot be recovered)
- Requires TWO approval tokens (scope: macro:remove)
- Changes the workbook structure
- Creates a clean .xlsx without any VBA
"""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

from excel_agent.core.macro_handler import has_macros
from excel_agent.core.version_hash import compute_file_hash
from excel_agent.governance.audit_trail import AuditTrail
from excel_agent.governance.token_manager import ApprovalTokenManager
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import (
    add_common_args,
    add_governance_args,
    create_parser,
    validate_input_path,
    validate_output_path,
)
from excel_agent.utils.exceptions import ValidationError
from excel_agent.utils.json_io import build_response


def _remove_vba(input_path: Path, output_path: Path) -> dict[str, object]:
    """Remove VBA from workbook by stripping vbaProject.bin."""
    # Read the original xlsm as a zip
    with zipfile.ZipFile(input_path, "r") as zin:
        # Create new zip in memory
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                # Skip VBA-related files
                if "vba" in item.filename.lower():
                    continue
                data = zin.read(item.filename)
                zout.writestr(item, data)

    # Write to output
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as f:
        f.write(buffer.getvalue())

    return {"vba_removed": True, "output_path": str(output_path)}


def _run() -> dict[str, object]:
    parser = create_parser("Remove VBA macros from workbook (IRREVERSIBLE, requires 2 tokens).")
    add_common_args(parser)
    add_governance_args(parser)
    parser.add_argument("--token2", type=str, required=True, help="Second approval token")
    args = parser.parse_args()

    input_path = validate_input_path(args.input)
    output_path = validate_output_path(
        args.output or args.input.replace(".xlsm", "_cleaned.xlsx"),
        create_parents=True,
    )

    file_hash = compute_file_hash(input_path)

    # Verify both tokens
    if not args.token:
        raise ValidationError("First approval token required (--token)")
    if not args.token2:
        raise ValidationError("Second approval token required (--token2)")

    # Validate tokens
    mgr = ApprovalTokenManager()
    mgr.validate_token(args.token, "macro:remove", input_path)
    mgr.validate_token(args.token2, "macro:remove", input_path)

    # Check if file has macros
    if not has_macros(input_path):
        return build_response(
            "success",
            {
                "has_macros": False,
                "message": "No VBA macros found in file",
            },
            warnings=["File does not contain VBA macros"],
        )

    # Remove VBA
    result = _remove_vba(input_path, output_path)

    # Log audit
    audit = AuditTrail()
    audit.log(
        tool="xls_remove_macros",
        scope="macro:remove",
        target_file=input_path,
        file_version_hash=file_hash,
        actor_nonce=args.token[:32],
        operation_details={"output_path": str(output_path)},
        impact={"vba_removed": True},
        success=True,
        exit_code=0,
    )

    return build_response(
        "success",
        result,
        warnings=[
            "VBA macros have been permanently removed.",
            "This operation is IRREVERSIBLE.",
        ],
    )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
