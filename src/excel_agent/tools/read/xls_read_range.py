"""xls_read_range: Extract cell data from a range with chunked streaming."""

from __future__ import annotations

import json
import sys

from excel_agent.core.agent import ExcelAgent
from excel_agent.core.chunked_io import read_range_chunked, read_range_full
from excel_agent.core.serializers import RangeSerializer
from excel_agent.tools._tool_base import run_tool
from excel_agent.utils.cli_helpers import add_common_args, create_parser, validate_input_path
from excel_agent.utils.json_io import ExcelAgentEncoder, build_response


def _run() -> dict:
    parser = create_parser("Extract cell data from a range as JSON.")
    add_common_args(parser)
    parser.add_argument("--range", type=str, required=True, help="Cell range (e.g., A1:C10)")
    parser.add_argument(
        "--chunked", action="store_true", help="Stream as JSONL (one row per line)"
    )
    args = parser.parse_args()
    path = validate_input_path(args.input)

    with ExcelAgent(path, mode="r") as agent:
        wb = agent.workbook
        serializer = RangeSerializer(workbook=wb)
        coord = serializer.parse(args.range, default_sheet=args.sheet)

        sheet_name = coord.sheet or (args.sheet or wb.sheetnames[0])
        ws = wb[sheet_name]

        min_row = coord.min_row
        min_col = coord.min_col
        max_row = coord.max_row or ws.max_row or min_row
        max_col = coord.max_col or ws.max_column or min_col

        if args.chunked or args.format == "jsonl":
            # JSONL streaming mode — write directly to stdout, bypass normal return
            for chunk in read_range_chunked(ws, min_row, min_col, max_row, max_col):
                for row_data in chunk:
                    line = json.dumps(
                        {"values": row_data}, cls=ExcelAgentEncoder, ensure_ascii=False
                    )
                    sys.stdout.write(line + "\n")
                    sys.stdout.flush()
            sys.exit(0)

        # Normal JSON mode
        values = read_range_full(ws, min_row, min_col, max_row, max_col)
        return build_response(
            "success",
            {
                "range": args.range,
                "sheet": sheet_name,
                "rows": len(values),
                "cols": (max_col - min_col + 1),
                "values": values,
            },
            workbook_version=agent.version_hash,
        )


def main() -> None:
    run_tool(_run)


if __name__ == "__main__":
    main()
