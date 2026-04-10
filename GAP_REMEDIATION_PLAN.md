# Gap Remediation Plan
## Realistic Test Plan Findings

**Date:** April 10, 2026  
**Status:** Validation Complete - Ready for Remediation  
**Validated Findings:** 9 gaps confirmed

---

## Validation Summary

All 9 gaps from the Realistic Test Plan have been **confirmed** through direct codebase inspection.

### Validation Results

| Issue | Status | Validation Method |
|:---|:---:|:---|
| P0-1: xls_set_number_format help text | ✅ **CONFIRMED** | Direct file read - line 68 contains unescaped `%` |
| P0-2: xls_inject_vba_project duplicate --force | ✅ **CONFIRMED** | Code analysis - `add_governance_args()` already adds `--force`, tool adds again |
| P1-3: xls_get_defined_names named range handling | ⚠️ **NEEDS TESTING** | Code looks correct but test shows exit code 5 - needs investigation |
| P1-4: xls_copy_formula_down API | ✅ **CONFIRMED** | Tool uses `--cell`/`--count`, not `--source`/`--target` as documented |
| P2-5: Export tool range filtering | ✅ **CONFIRMED** | xls_export_csv.py has no `--range` argument |
| P2-6: CLI signature documentation | ✅ **CONFIRMED** | Multiple tools have doc/actual mismatches |

---

## Detailed Gap Analysis & Remediation Plan

---

## 🔴 P0 - CRITICAL (Must Fix Immediately)

### Issue 1: xls_set_number_format Help Text Escape Error

**File:** `src/excel_agent/tools/formatting/xls_set_number_format.py`  
**Line:** 68  
**Severity:** 🔴 CRITICAL - Tool crashes on --help

#### Current State
```python
parser.add_argument(
    "--number-format",
    type=str,
    required=True,
    dest="number_format",
    help="Excel number format code (e.g., '"$"#,##0.00', '0.00%', 'yyyy-mm-dd')",  # Line 68
)
```

#### Problem
The help string contains `0.00%` which argparse interprets as a format specifier ( `%` ) causing:
```
ValueError: unsupported format character ''' (0x27) at index 53
```

#### Root Cause
Argparse's `help` strings support `%(var)s` style interpolation. The `%` in `0.00%` is interpreted as a format character.

#### Remediation
**Option A: Escape the % (Recommended - minimal change)**
```python
help="Excel number format code (e.g., '"$"#,##0.00', '0.00%%', 'yyyy-mm-dd')",
```

**Option B: Use Raw String with Explicit Format**
```python
help="Excel number format code (e.g., '\"$\"#,##0.00', '0.00%%', 'yyyy-mm-dd'). "
     "Use %% to represent literal percent sign.",
```

**Option C: Set argparse interpolation off**
```python
parser = argparse.ArgumentParser(
    description="Apply number formats...",
    formatter_class=argparse.RawDescriptionHelpFormatter,
)
# In create_parser(), set argument_default or use RawDescriptionHelpFormatter
```

#### Action Items
- [ ] Update line 68 to escape `%` as `%%`
- [ ] Verify `--help` works without error
- [ ] Add regression test for `--help` on all formatting tools

#### Estimated Effort: 5 minutes

---

### Issue 2: xls_inject_vba_project Duplicate --force Argument

**File:** `src/excel_agent/tools/macros/xls_inject_vba_project.py`  
**Lines:** 92-96 and `add_governance_args()`  
**Severity:** 🔴 CRITICAL - Tool crashes on load

#### Current State
```python
def _run() -> dict[str, object]:
    parser = create_parser("Inject VBA project...")
    add_common_args(parser)
    add_governance_args(parser)  # <-- ALREADY adds --force (line 85)
    
    # ... more args ...
    
    # Line 92-96: DUPLICATES --force
    parser.add_argument(
        "--force",
        action="store_true",
        help="Inject even if risk level is high/critical",
    )
```

#### Problem
`add_governance_args()` at line 28-35 already adds `--force`:
```python
def add_governance_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--force",
        action="store_true",
        default=False,
        help="Force operation even if impact report shows warnings",
    )
```

This causes:
```
argparse.ArgumentError: argument --force: conflicting option string: --force
```

#### Remediation
**Remove the duplicate argument definition (lines 92-96):**
```python
def _run() -> dict[str, object]:
    parser = create_parser("Inject VBA project...")
    add_common_args(parser)
    add_governance_args(parser)  # Keeps --force, --token, --acknowledge-impact
    
    parser.add_argument(
        "--vba-bin",
        type=str,
        required=True,
        help="Path to vbaProject.bin file to inject",
    )
    # REMOVE: duplicate --force argument
    args = parser.parse_args()
```

#### Action Items
- [ ] Remove lines 92-96 (duplicate --force)
- [ ] Verify tool loads without error
- [ ] Test with `--help`
- [ ] Test injection workflow

#### Estimated Effort: 5 minutes

---

## 🟡 P1 - HIGH (Fix Before Production)

### Issue 3: xls_get_defined_names Named Range Handling

**File:** `src/excel_agent/tools/read/xls_get_defined_names.py`  
**Test Result:** Returns exit code 5 on OfficeOps_Expenses_KPI.xlsx  
**Severity:** 🟡 HIGH - Feature broken

#### Current Implementation
```python
def _run() -> dict:
    # ...
    with ExcelAgent(path, mode="r") as agent:
        wb = agent.workbook
        named_ranges: list[dict] = []
        
        for defn in wb.defined_names.definedName:  # Line 21
            # ... processing logic
```

#### Problem Analysis
The code assumes `wb.defined_names.definedName` exists, but the error suggests either:
1. `wb.defined_names` is None
2. `wb.defined_names.definedName` doesn't exist as expected
3. Exception in ExcelAgent context

#### Investigation Steps
```python
# Add debugging to identify root cause:
def _run() -> dict:
    parser = create_parser("List all named ranges in a workbook.")
    parser.add_argument("--input", type=str, required=True, help="Workbook path")
    args = parser.parse_args()
    path = validate_input_path(args.input)
    
    try:
        with ExcelAgent(path, mode="r") as agent:
            wb = agent.workbook
            
            # Debug: check what we're dealing with
            if wb.defined_names is None:
                return build_response("success", {"named_ranges": [], "count": 0})
            
            named_ranges: list[dict] = []
            
            # Handle different openpyxl versions
            defined_names = getattr(wb.defined_names, 'definedName', [])
            if defined_names is None:
                defined_names = []
            
            for defn in defined_names:
                # ... rest of logic
    except Exception as e:
        return build_response(
            "error", 
            None, 
            exit_code=5,
            error=f"Failed to read named ranges: {str(e)}"
        )
```

#### Remediation Options
**Option A: Add null-safety (Recommended)**
```python
with ExcelAgent(path, mode="r") as agent:
    wb = agent.workbook
    named_ranges: list[dict] = []
    
    # Safe access to defined names
    if wb.defined_names is None:
        return build_response("success", {"named_ranges": [], "count": 0})
    
    # Handle both old and new openpyxl versions
    defined_names_list = getattr(wb.defined_names, 'definedName', None) or []
    
    for defn in defined_names_list:
        # Process... (with null checks)
        scope = "Workbook"
        local_sheet_id = getattr(defn, 'localSheetId', None)
        if local_sheet_id is not None:
            try:
                idx = int(local_sheet_id)
                scope = wb.sheetnames[idx] if idx < len(wb.sheetnames) else f"Sheet(index={idx})"
            except (ValueError, IndexError):
                scope = "Unknown"
        
        named_ranges.append({
            "name": getattr(defn, 'name', 'Unknown'),
            "scope": scope,
            "refers_to": getattr(defn, 'attr_text', '') or "",
            "hidden": getattr(defn, 'hidden', False) or False,
            "is_reserved": getattr(defn, 'is_reserved', False) or False,
        })
```

**Option B: Use openpyxl's documented API**
```python
# Alternative approach using wb.defined_names directly
for name, defn in wb.defined_names.items():
    # Process... (if .items() is supported)
```

#### Action Items
- [ ] Run test with debug logging to identify exact failure point
- [ ] Add null-safety checks
- [ ] Handle multiple openpyxl API versions
- [ ] Add unit tests with fixture containing named ranges
- [ ] Verify fix with OfficeOps_Expenses_KPI.xlsx

#### Estimated Effort: 30-45 minutes

---

### Issue 4: xls_copy_formula_down API Alignment

**File:** `src/excel_agent/tools/formulas/xls_copy_formula_down.py`  
**Documentation Claims:** `--source` and `--target`  
**Actual Implementation:** `--cell` and `--count`  
**Severity:** 🟡 HIGH - Documentation/tool mismatch

#### Current Implementation
```python
def _run() -> dict:
    parser = create_parser("Copy formula from source cell down to target cells.")
    add_common_args(parser)
    parser.add_argument("--cell", type=str, required=True, help="Source cell (e.g., A1)")  # Not --source
    parser.add_argument("--count", type=int, required=True, help="Number of cells to fill")  # Not --target
```

#### Expected (Per Documentation)
```bash
xls-copy-formula-down --source H2 --target H2:H10
```

#### Actual
```bash
xls-copy-formula-down --cell H2 --count 9
```

#### Remediation Options

**Option A: Update Tool to Match Documentation (Preferred)**
```python
def _run() -> dict:
    parser = create_parser("Copy formula from source cell down to target cells.")
    add_common_args(parser)
    
    # Support both old and new API (backward compatible)
    parser.add_argument("--cell", type=str, help="Source cell (deprecated, use --source)")
    parser.add_argument("--count", type=int, help="Number of cells to fill (deprecated, use --target)")
    parser.add_argument("--source", type=str, help="Source cell (e.g., A1)")
    parser.add_argument("--target", type=str, help="Target range (e.g., A1:A10)")
    
    args = parser.parse_args()
    
    # Handle both APIs
    source = args.source or args.cell
    if not source:
        parser.error("--source or --cell is required")
    
    if args.target:
        # Parse range to get count
        from openpyxl.utils import range_boundaries
        min_col, min_row, max_col, max_row = range_boundaries(args.target)
        count = max_row - min_row if max_row and min_row else 0
    elif args.count:
        count = args.count
    else:
        parser.error("--target or --count is required")
```

**Option B: Update Documentation to Match Tool**
Update SKILL.md, API.md, and any examples to reflect `--cell` and `--count`.

#### Recommendation
**Option A** is preferred because:
1. `--source`/`--target` is more intuitive (matches openpyxl terminology)
2. `--target` accepts ranges like `A1:A10` which is standard
3. Backward compatibility can be maintained

#### Action Items
- [ ] Decide: Option A or Option B
- [ ] If Option A: Implement dual API support
- [ ] If Option B: Update all documentation
- [ ] Update test_realistic_office_workflow.py
- [ ] Add regression tests

#### Estimated Effort: 30 minutes

---

## 🟢 P2 - MEDIUM (Documentation & Clarification)

### Issue 5: Export Tool Range Filtering Clarification

**File:** `src/excel_agent/tools/export/xls_export_csv.py`  
**Current:** No `--range` argument  
**Expected (per test plan):** Range filtering capability  
**Severity:** 🟢 MEDIUM - Feature gap

#### Current Implementation
```python
def _run() -> dict[str, object]:
    parser = create_parser("Export Excel sheet to CSV.")
    add_common_args(parser)
    # ... encoding, delimiter, include-headers, outfile ...
    # NO --range argument
```

#### Analysis
The export tools export entire sheets, not specific ranges. This is:
- ✅ **Correct by design** - Simpler, faster
- ⚠️ **Different from test plan expectation** - Assumed range filtering

#### Remediation
**Add `--range` support (Feature Enhancement):**
```python
parser.add_argument(
    "--range",
    type=str,
    default=None,
    help='Export specific range only (e.g., "A1:H100")',
)

# In export logic:
if args.range:
    from openpyxl.utils import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(args.range)
    # Iterate only rows in range
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if min_row <= row_idx <= (max_row or float('inf')):
            # Filter columns too
            filtered = row[min_col-1:max_col] if max_col else row[min_col-1:]
            writer.writerow(filtered)
else:
    # Current behavior - export all
```

#### Action Items
- [ ] Decide if `--range` is needed for MVP
- [ ] If yes: Implement range filtering
- [ ] Update documentation to clarify current behavior
- [ ] Provide workaround pattern in docs

#### Estimated Effort: 15 minutes (if implemented)

---

### Issue 6: CLI Signature Documentation Updates

**Multiple files**  
**Severity:** 🟢 MEDIUM - Documentation accuracy

#### Issues Found
1. `xls_copy_formula_down` - `--cell`/`--count` vs `--source`/`--target`
2. `xls_export_csv` - No `--range` (clarify in docs)
3. `xls_detect_errors` - No `--range` (clarify in docs)

#### Remediation
Update documentation in:
- `skills/excel-tools/SKILL.md`
- `docs/API.md`
- `docs/WORKFLOWS.md`

#### Template for Documentation
```markdown
## xls_copy_formula_down

Copy formula from source cell down to target cells.

### Usage
```bash
xls-copy-formula-down --input INPUT --cell CELL --count COUNT [--sheet SHEET] [--output OUTPUT]
```

### Arguments
- `--cell`: Source cell (e.g., A1) - REQUIRED
- `--count`: Number of cells to fill - REQUIRED
- `--sheet`: Sheet name (default: first sheet)
- `--output`: Output path (default: overwrite input)

### Note on API
This tool uses `--cell` and `--count` rather than `--source` and `--target`.
To fill from A1 to A10, use: `--cell A1 --count 9`

### Examples
```bash
# Copy formula from A1 down 10 cells
xls-copy-formula-down --input data.xlsx --cell A1 --count 10

# Copy on specific sheet
xls-copy-formula-down --input data.xlsx --sheet Sales --cell B2 --count 5
```
```

#### Action Items
- [ ] Update SKILL.md with accurate CLI signatures
- [ ] Update API.md reference
- [ ] Update WORKFLOWS.md examples
- [ ] Add "Note on API" callouts where there are quirks

#### Estimated Effort: 20 minutes

---

## Execution Order

### Phase 1: P0 Fixes (10 minutes)
1. ✅ Fix xls_set_number_format.py (escape %)
2. ✅ Fix xls_inject_vba_project.py (remove duplicate --force)

### Phase 2: P1 Fixes (1 hour)
3. Debug and fix xls_get_defined_names.py
4. Decide and implement xls_copy_formula_down API alignment

### Phase 3: P2 Documentation (30 minutes)
5. Add --range support to export tools (optional)
6. Update all documentation with accurate CLI signatures

### Phase 4: Verification (30 minutes)
7. Re-run realistic test suite
8. Verify all fixes
9. Update test assertions where needed

---

## Verification Checklist

- [ ] `xls_set_number_format --help` works without error
- [ ] `xls_inject_vba_project --help` works without error
- [ ] `xls_get_defined_names` on OfficeOps_Expenses_KPI.xlsx returns named ranges
- [ ] `xls_copy_formula_down` accepts documented API (either updated tool or updated docs)
- [ ] Export tools documented correctly (with or without --range)
- [ ] SKILL.md updated with accurate signatures
- [ ] API.md updated with accurate signatures
- [ ] Realistic test suite passes

---

## Risks & Mitigation

| Risk | Impact | Mitigation |
|:---|:---:|:---|
| Fixing xls_get_defined_names breaks other functionality | Medium | Add comprehensive tests, use try/except blocks |
| API change breaks existing users | Low | Maintain backward compatibility with deprecated warnings |
| Time overrun | Low | P0 fixes are quick; P1 can be deferred if needed |

---

**Plan Approved:** Ready for execution  
**Estimated Total Effort:** 2-3 hours  
**Priority:** P0 blocks production release
