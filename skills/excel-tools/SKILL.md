---
name: excel-tools
description: Create, modify, and manipulate .xlsx Excel files safely using the excel-agent-tools CLI suite. This skill should be used when users need to read, write, format, calculate, or export Excel workbooks programmatically. It provides governance-first workflows with HMAC-SHA256 token protection, formula integrity preservation, and headless operation without Microsoft Excel dependency. Use for tasks including data extraction, sheet manipulation, formula calculations, macro safety scanning, and format conversion.
license: MIT
allowed-tools:
  - bash
  - python
metadata:
  project-version: "1.0.0"
  total-tools: "53"
  calculation-tiers: "2"
  token-scopes: "7"
  coverage: "90%"
---

# Excel Tools Skill

Create, modify, and manipulate Excel (.xlsx/.xlsm) files safely using excel-agent-tools - a headless, governance-first CLI suite of 53 tools designed for AI agents.

## When to Use This Skill

Use this skill when:
- Reading or extracting data from Excel files
- Writing data to Excel workbooks (new or existing)
- Modifying sheet structure (add/delete/rename sheets, insert/delete rows/columns)
- Formatting cells (styles, conditional formatting, number formats)
- Calculating formulas (Tier 1 in-process or Tier 2 LibreOffice)
- Working with Excel objects (tables, charts, images, comments)
- Scanning for macro safety in .xlsm files
- Exporting to PDF, CSV, or JSON formats

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│  CLI Tools (53) → ExcelAgent → Core Libraries (openpyxl)  │
│                                                            │
│  Categories:                                               │
│  • Governance (6) - clone, validate, tokens               │
│  • Read (7) - extract data, metadata, formulas            │
│  • Write (4) - create, modify cell data                  │
│  • Structure (8) - sheets, rows, columns                  │
│  • Cells (4) - merge, unmerge, references               │
│  • Formulas (6) - calculate, errors, conversions        │
│  • Objects (5) - tables, charts, images                │
│  • Formatting (5) - styles, conditional formats         │
│  • Macros (5) - safety scan, VBA management              │
│  • Export (3) - PDF, CSV, JSON                           │
└─────────────────────────────────────────────────────────────┘
```

## Key Principles

1. **Clone-Before-Edit**: Always use `xls-clone-workbook` first; never mutate originals
2. **Token Protection**: Destructive ops require HMAC-SHA256 scoped tokens with TTL
3. **Formula Integrity**: Pre-flight dependency checks prevent #REF! errors
4. **JSON-Native**: All tools accept/return JSON; exit codes 0-5
5. **Headless**: No Excel/COM dependency; runs on Linux/macOS/Windows

## Core Workflow

### Step 1: Clone Source (Safety)

```bash
# Clone before any modifications
xls-clone-workbook --input original.xlsx --output-dir ./work/
# Returns: {"data": {"clone_path": "./work/original_20260409T143022_abc.xlsx"}}
```

### Step 2: Read/Extract Data

```bash
# Read range
xls-read-range --input ./work/original_*.xlsx --range A1:C10 --sheet Sheet1

# Get metadata
xls-get-workbook-metadata --input ./work/original_*.xlsx
```

### Step 3: Modify (if needed)

```bash
# Write data
xls-write-range --input ./work/original_*.xlsx --range F1 \
  --data '[["Header", "Value"], ["A", 100]]'

# Requires token for destructive ops
TOKEN=$(xls-approve-token --scope sheet:delete --file ./work/original_*.xlsx | jq -r '.data.token')
xls-delete-sheet --input ./work/original_*.xlsx --name "OldSheet" --token "$TOKEN"
```

### Step 4: Calculate (if formulas present)

```bash
# Auto Tier 1 → Tier 2 fallback
xls-recalculate --input ./work/original_*.xlsx --output ./work/original_*.xlsx
```

### Step 5: Validate & Export

```bash
# Validate integrity
xls-validate-workbook --input ./work/original_*.xlsx

# Export
xls-export-csv --input ./work/original_*.xlsx --outfile output.csv
xls-export-pdf --input ./work/original_*.xlsx --outfile output.pdf
```

## Token Scopes

| Scope | Risk | Operations |
|-------|------|------------|
| `sheet:delete` | High | Remove entire sheet |
| `sheet:rename` | Medium | Rename + update references |
| `range:delete` | High | Delete rows/columns |
| `formula:convert` | High | Formulas → values (irreversible) |
| `macro:remove` | Critical | Strip VBA (requires 2 tokens) |
| `macro:inject` | Critical | Inject VBA project |
| `structure:modify` | High | Batch structural changes |

## Exit Codes

| Code | Meaning | Action |
|------|---------|--------|
| 0 | Success | Parse JSON, proceed |
| 1 | Validation/Impact Denial | Fix input or acknowledge impact |
| 2 | File Not Found | Verify path |
| 3 | Lock Contention | Exponential backoff retry |
| 4 | Permission Denied | Generate new token |
| 5 | Internal Error | Alert operator |

## Important Constraints

- Export tools use `--outfile` NOT `--output` (avoids argparse conflict)
- LibreOffice required for PDF export and Tier 2 calculation
- `EXCEL_AGENT_SECRET` env var required for token operations
- Token TTL: 1-3600 seconds (default 300)
- Chunked mode returns JSONL not single JSON

## Referenced Resources

- `references/workflow-patterns.md` - Common patterns (clone-modify-export, etc.)
- `references/tool-reference.md` - All 53 tools with full JSON examples
- `references/troubleshooting.md` - Common issues and solutions
- `scripts/create_workbook.py` - Helper to create workbooks programmatically
- `scripts/batch_process.py` - Process multiple files
- `assets/template.xlsx` - Blank workbook template
- `assets/template_with_data.xlsx` - Sample workbook with formulas

## Installation

```bash
pip install excel-agent-tools

# Optional: LibreOffice for PDF/Tier 2 calc
# Ubuntu: sudo apt-get install -y libreoffice-calc
```

## Quick Examples

### Extract Data to JSON
```bash
xls-read-range --input data.xlsx --range A1:E100 --sheet Sales | jq '.data.values'
```

### Create New Workbook
```bash
xls-create-new --output report.xlsx --sheets "Summary,Data,Charts"
```

### Calculate Formulas
```bash
xls-recalculate --input report.xlsx --output report_calculated.xlsx
```

### Safe Sheet Deletion
```bash
TOKEN=$(xls-approve-token --scope sheet:delete --file report.xlsx --ttl 300 | jq -r '.data.token')
xls-delete-sheet --input report.xlsx --output report.xlsx --name "Draft" --token "$TOKEN" --acknowledge-impact
```

### Export to CSV
```bash
xls-export-csv --input report.xlsx --outfile report.csv --sheet "Summary"
```

## See Also

- Full documentation in project `/docs/` folder
- `CLAUDE.md` for complete architecture briefing
- `Project_Architecture_Document.md` for deep technical details
