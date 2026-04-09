# Excel Tools Complete Reference

All 53 tools with full JSON examples.

## Governance Tools (6)

### xls-clone-workbook
**Purpose**: Atomic copy to safe working directory.

**CLI**:
```bash
xls-clone-workbook --input source.xlsx [--output-dir ./work/]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "clone_path": "./work/source_20260409T143022_a3f7e2d1.xlsx",
    "source_hash": "sha256:abc...",
    "clone_hash": "sha256:abc...",
    "timestamp": "20260409T143022"
  }
}
```

---

### xls-validate-workbook
**Purpose**: OOXML compliance, broken refs, circular refs.

**CLI**:
```bash
xls-validate-workbook --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "valid": true,
    "errors": [],
    "warnings": ["Large image in B5"],
    "circular_refs": [],
    "broken_references": 0
  }
}
```

---

### xls-approve-token
**Purpose**: Generate HMAC-SHA256 scoped token.

**CLI**:
```bash
xls-approve-token --scope sheet:delete --file workbook.xlsx [--ttl 300]
```

**Scopes**: `sheet:delete`, `sheet:rename`, `range:delete`, `formula:convert`, `macro:remove`, `macro:inject`, `structure:modify`

**Output**:
```json
{
  "status": "success",
  "data": {
    "token": "eyJzY29wZSI6InNoZWV0OmRlbGV0ZSIs...",
    "scope": "sheet:delete",
    "expires_at": "2026-04-09T14:35:00Z",
    "file_hash": "sha256:abc..."
  }
}
```

---

### xls-version-hash
**Purpose**: Compute geometry hash for modification detection.

**CLI**:
```bash
xls-version-hash --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "geometry_hash": "sha256:abc...",
    "file_hash": "sha256:xyz..."
  }
}
```

---

### xls-lock-status
**Purpose**: Check if workbook is locked.

**CLI**:
```bash
xls-lock-status --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "locked": false,
    "lock_file_exists": false
  }
}
```

---

### xls-dependency-report
**Purpose**: Export formula dependency graph.

**CLI**:
```bash
xls-dependency-report --input workbook.xlsx [--sheet Sheet1]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "stats": {
      "total_cells": 500,
      "total_formulas": 50,
      "total_edges": 120
    },
    "graph": {
      "Sheet1!B1": ["Sheet1!C1"]
    },
    "circular_refs": []
  }
}
```

---

## Read Tools (7)

### xls-read-range
**Purpose**: Extract data as JSON.

**CLI**:
```bash
xls-read-range --input workbook.xlsx --range A1:C10 [--sheet Sheet1] [--chunked]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "values": [["Header", "Value"], ["A", 100]],
    "range": "A1:B2",
    "sheet": "Sheet1"
  }
}
```

---

### xls-get-sheet-names
**Purpose**: List all sheets.

**CLI**:
```bash
xls-get-sheet-names --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "sheets": [
      {"index": 0, "name": "Sheet1", "visibility": "visible"},
      {"index": 1, "name": "Data", "visibility": "hidden"}
    ]
  }
}
```

---

### xls-get-workbook-metadata
**Purpose**: High-level statistics.

**CLI**:
```bash
xls-get-workbook-metadata --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "sheet_count": 3,
    "total_formulas": 47,
    "named_ranges": ["SalesData"],
    "tables": ["Table1"],
    "has_macros": false,
    "file_size_bytes": 15234
  }
}
```

---

### xls-get-defined-names
**Purpose**: List named ranges.

**CLI**:
```bash
xls-get-defined-names --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "named_ranges": [
      {"name": "SalesData", "scope": "workbook", "refers_to": "Sheet3!$A$1:$B$5"}
    ]
  }
}
```

---

### xls-get-table-info
**Purpose**: List Excel Table objects.

**CLI**:
```bash
xls-get-table-info --input workbook.xlsx
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "tables": [
      {
        "name": "Table1",
        "sheet": "Sheet1",
        "range": "A1:D10",
        "columns": ["ID", "Name", "Value", "Total"]
      }
    ]
  }
}
```

---

### xls-get-cell-style
**Purpose**: Get style as JSON.

**CLI**:
```bash
xls-get-cell-style --input workbook.xlsx --cell A1 [--sheet Sheet1]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "font": {"name": "Arial", "size": 11, "bold": false},
    "fill": {"fgColor": "FFFFFFFF"},
    "border": {"top": null},
    "alignment": {"horizontal": "left"},
    "number_format": "General"
  }
}
```

---

### xls-get-formula
**Purpose**: Get formula from cell.

**CLI**:
```bash
xls-get-formula --input workbook.xlsx --cell A1 [--sheet Sheet1]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "cell": "A1",
    "formula": "=SUM(B1:B10)",
    "references": ["B1:B10"]
  }
}
```

---

## Write Tools (4)

### xls-create-new
**Purpose**: Create blank workbook.

**CLI**:
```bash
xls-create-new --output workbook.xlsx [--sheets "Sheet1,Sheet2"]
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "path": "workbook.xlsx",
    "sheets_created": ["Sheet1", "Sheet2"]
  }
}
```

---

### xls-create-from-template
**Purpose**: Clone template with substitution.

**CLI**:
```bash
xls-create-from-template --template template.xltx --output workbook.xlsx \
  --vars '{"company": "Acme", "year": "2026"}'
```

**Output**:
```json
{
  "status": "success",
  "data": {
    "path": "workbook.xlsx",
    "substitutions": 3
  }
}
```

---

### xls-write-range
**Purpose**: Write 2D array to range.

**CLI**:
```bash
xls-write-range --input workbook.xlsx --output workbook.xlsx --range A1 \
  --data '[["Name", "Value"], ["A", 100]]' [--sheet Sheet1]
```

**Output**:
```json
{
  "status": "success",
  "data": {"range_written": "A1:B2"},
  "impact": {"cells_modified": 4}
}
```

---

### xls-write-cell
**Purpose**: Write single cell.

**CLI**:
```bash
xls-write-cell --input workbook.xlsx --output workbook.xlsx --cell A1 \
  --value "=SUM(B1:B10)" [--type formula] [--sheet Sheet1]
```

**Types**: `auto`, `string`, `number`, `formula`, `date`, `boolean`

**Output**:
```json
{
  "status": "success",
  "data": {"cell": "A1", "type": "formula"},
  "impact": {"cells_modified": 1}
}
```

---

## Structure Tools (8) - ⚠️ Token Required

### xls-add-sheet
**CLI**: `xls-add-sheet --input X --output X --name "New" [--position 0]`

### xls-delete-sheet ⚠️
**Scope**: `sheet:delete`
**CLI**: `xls-delete-sheet --input X --output X --name "Sheet" --token T [--acknowledge-impact]`

**Denial Output**:
```json
{
  "status": "denied",
  "exit_code": 1,
  "guidance": "Run xls-update-references --updates '[...]' before retrying",
  "impact": {"broken_references": 7}
}
```

### xls-rename-sheet ⚠️
**Scope**: `sheet:rename`
**CLI**: `xls-rename-sheet --input X --output X --old "Old" --new "New" --token T`

### xls-insert-rows
**CLI**: `xls-insert-rows --input X --output X --sheet S --before-row 5 --count 3`

### xls-delete-rows ⚠️
**Scope**: `range:delete`
**CLI**: `xls-delete-rows --input X --output X --sheet S --start-row 5 --count 3 --token T`

### xls-insert-columns
**CLI**: `xls-insert-columns --input X --output X --sheet S --before-col C --count 2`

### xls-delete-columns ⚠️
**Scope**: `range:delete`
**CLI**: `xls-delete-columns --input X --output X --sheet S --start-col C --count 2 --token T`

### xls-move-sheet
**CLI**: `xls-move-sheet --input X --output X --name "Sheet" --position 0`

---

## Cells Tools (4)

### xls-merge-cells
**CLI**: `xls-merge-cells --input X --output X --range A1:C1`

### xls-unmerge-cells
**CLI**: `xls-unmerge-cells --input X --output X --range A1:C1`

### xls-delete-range ⚠️
**Scope**: `range:delete`
**CLI**: `xls-delete-range --input X --output X --range A1:C5 --shift up|left --token T`

### xls-update-references
**CLI**: `xls-update-references --input X --output X --updates '[{"old": "...", "new": "..."}]'`

---

## Formulas Tools (6)

### xls-set-formula
**CLI**: `xls-set-formula --input X --output X --cell A1 --formula "=SUM(B1:B10)"`

### xls-recalculate
**CLI**: `xls-recalculate --input X --output X [--tier 1|2]`

**Output**:
```json
{
  "status": "success",
  "data": {
    "formula_count": 47,
    "calculated_count": 47,
    "error_count": 0,
    "engine": "tier1_formulas",
    "recalc_time_ms": 45.2
  }
}
```

### xls-detect-errors
**CLI**: `xls-detect-errors --input X [--sheet S]`

**Output**:
```json
{
  "status": "success",
  "data": {
    "errors": [
      {"sheet": "Sheet1", "cell": "A1", "error": "#REF!", "formula": "=Sheet2!A1"}
    ]
  }
}
```

### xls-convert-to-values ⚠️
**Scope**: `formula:convert`
**CLI**: `xls-convert-to-values --input X --output X --range A1:C10 --token T`

### xls-copy-formula-down
**CLI**: `xls-copy-formula-down --input X --output X --cell A1 --count 10`

### xls-define-name
**CLI**: `xls-define-name --input X --output X --name "SalesData" --refers-to "Sheet1!A1:B10"`

---

## Objects Tools (5)

### xls-add-table
**CLI**: `xls-add-table --input X --output X --range A1:D10 --name "Table1" [--has-totals]`

### xls-add-chart
**CLI**: `xls-add-chart --input X --output X --type bar --data-range "A1:B10" --position "E1"`

### xls-add-image
**CLI**: `xls-add-image --input X --output X --image logo.png --cell A1 [--width 200]`

### xls-add-comment
**CLI**: `xls-add-comment --input X --output X --cell A1 --text "Review"`

### xls-set-data-validation
**CLI**: `xls-set-data-validation --input X --output X --range A1:A10 --type list --source '["Yes", "No"]'`

---

## Formatting Tools (5)

### xls-format-range
**CLI**: `xls-format-range --input X --output X --range A1:C10 --spec '{"font": {"bold": true}}'`

### xls-set-column-width
**CLI**: `xls-set-column-width --input X --output X --column A [--width 20|--auto-fit]`

### xls-freeze-panes
**CLI**: `xls-freeze-panes --input X --output X --row 2 [--column C]`

### xls-apply-conditional-formatting
**CLI**: `xls-apply-conditional-formatting --input X --output X --range A1:A100 --type colorscale --colors '["FF0000", "FFFF00", "00FF00"]'`

### xls-set-number-format
**CLI**: `xls-set-number-format --input X --output X --range A1:A10 --format '$#,##0.00'`

---

## Macros Tools (5) - ⚠️⚠️ Double Token

### xls-has-macros
**CLI**: `xls-has-macros --input file.xlsm`

### xls-inspect-macros
**CLI**: `xls-inspect-macros --input file.xlsm`

### xls-validate-macro-safety
**CLI**: `xls-validate-macro-safety --input file.xlsm`

**Output**:
```json
{
  "status": "success",
  "data": {
    "risk_level": "high",
    "auto_exec_triggers": ["AutoOpen"],
    "suspicious_keywords": ["Shell"]
  }
}
```

### xls-remove-macros ⚠️⚠️
**Scope**: `macro:remove` × 2
**CLI**: `xls-remove-macros --input file.xlsm --output file.xlsx --token T1 --token T2`

### xls-inject-vba-project ⚠️
**Scope**: `macro:inject`
**CLI**: `xls-inject-vba-project --input file.xlsx --output file.xlsm --vba-bin project.bin --token T`

---

## Export Tools (3)

### xls-export-pdf
**CLI**: `xls-export-pdf --input workbook.xlsx --outfile output.pdf [--recalc]`

**Note**: Uses `--outfile` not `--output` (avoids argparse conflict)

**Output**:
```json
{
  "status": "success",
  "data": {"output": "output.pdf", "pages": 3}
}
```

---

### xls-export-csv
**CLI**: `xls-export-csv --input workbook.xlsx --outfile output.csv [--sheet S] [--encoding utf-8]`

**Output**:
```json
{
  "status": "success",
  "data": {"output": "output.csv", "rows": 100}
}
```

---

### xls-export-json
**CLI**: `xls-export-json --input workbook.xlsx --outfile output.json [--orient records|values|columns]`

**Output**:
```json
{
  "status": "success",
  "data": {"output": "output.json", "records": 100}
}
```

---

## Format Notes

- **Dates**: Returned as ISO 8601 strings
- **Numbers**: Returned as JSON numbers
- **Formulas**: Returned as strings starting with `=`
- **Null**: Returned as `null` in JSON
- **Chunked**: Returns JSONL (one JSON per line)
- **Large**: Use `--chunked` for >100k rows
