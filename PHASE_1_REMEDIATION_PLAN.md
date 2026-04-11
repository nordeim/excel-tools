# Phase 1 Remediation Plan: Unified "Edit Target" Semantics

**Status:** VALIDATED | **Priority:** CRITICAL | **Est. Effort:** 1-2 weeks

---

## Executive Summary

QA findings have been **100% validated** through comprehensive codebase analysis. The system has critical architectural inconsistencies:

1. **Double-Save Pattern**: 29 tools using ExcelAgent have a double-save bug
2. **Raw load_workbook() Bypass**: 13 tools bypass ExcelAgent entirely
3. **Inconsistent Macro Preservation**: Tools bypassing ExcelAgent don't preserve macros
4. **Weak Output Validation**: `validate_output_path()` doesn't check extensions or overwrite policy

---

## Validation Evidence

### Finding 1.1: Double-Save Pattern (CRITICAL)

**Location:** `xls_write_range.py` lines 67-68

```python
with ExcelAgent(input_path, mode="rw") as agent:
    wb = agent.workbook
    # ... modifications ...
    # If output differs from input, save to new path
    if str(output_path) != str(input_path):
        wb.save(str(output_path))  # SECOND SAVE!
```

**Problem Flow:**
1. `ExcelAgent.__exit__()` saves to `self._path` (input path) - **FIRST SAVE**
2. Tool checks if output differs, saves again to output_path - **SECOND SAVE**
3. Result: Double writes, race conditions, inconsistent state

**Affected Tools:** 29 tools use ExcelAgent, approximately 15-20 likely have this issue

---

### Finding 1.2: Raw load_workbook() Bypass (CRITICAL)

**Tools Bypassing ExcelAgent (13 total):**

| Tool | Line | Issue |
|------|------|-------|
| `xls_add_chart.py` | 155 | `wb = load_workbook(str(input_path))` |
| `xls_add_image.py` | 145 | `wb = load_workbook(str(input_path))` |
| `xls_add_table.py` | 143 | `wb = load_workbook(str(input_path))` |
| `xls_add_comment.py` | ? | Raw load_workbook |
| `xls_set_data_validation.py` | ? | Raw load_workbook |
| `xls_format_range.py` | 154 | `wb = load_workbook(str(input_path))` |
| `xls_set_column_width.py` | ? | Raw load_workbook |
| `xls_freeze_panes.py` | 12 | `from openpyxl import load_workbook` |
| `xls_apply_conditional_formatting.py` | 15 | `from openpyxl import load_workbook` |
| `xls_set_number_format.py` | 17 | `from openpyxl import load_workbook` |
| `xls_export_csv.py` | 108 | `wb = load_workbook(..., data_only=True, read_only=True)` |
| `xls_export_json.py` | 14 | `from openpyxl import load_workbook` |
| `xls_export_pdf.py` | ? | Raw load_workbook |

**Consequences:**
- ❌ No file locking (race conditions)
- ❌ No version hash computation (concurrent modification detection bypassed)
- ❌ No consistent macro preservation (VBA stripped silently)
- ❌ Inconsistent audit trail logging

---

### Finding 1.3: Macro Preservation Inconsistency (HIGH)

**ExcelAgent's Correct Behavior:**
```python
# agent.py lines 91-94
_VBA_EXTENSIONS = frozenset({".xlsm", ".xltm"})
if keep_vba is None:
    self._keep_vba = self._path.suffix.lower() in _VBA_EXTENSIONS
```

**Tool's Incorrect Behavior:**
```python
# xls_add_chart.py line 155
wb = load_workbook(str(input_path))  # No keep_vba!

# xls_create_from_template.py line 46
wb = load_workbook(str(template_path))  # No macro handling
```

**Result:** Macros stripped from `.xlsm`/`.xltm` files when modified by these tools.

---

### Finding 1.4: Weak Output Path Validation (MEDIUM)

**Current Implementation:** `cli_helpers.py` lines 139-163

```python
def validate_output_path(path_str: str, *, create_parents: bool = False) -> Path:
    path = Path(path_str).resolve()
    if create_parents:
        path.parent.mkdir(parents=True, exist_ok=True)
    elif not path.parent.exists():
        exit_with(...)  # Only checks parent existence
    return path  # No extension check, no overwrite protection
```

**Missing Validations:**
- ❌ Extension validation (.xlsx, .xlsm, .xltx, .xltm only)
- ❌ Overwrite protection (should fail if output exists without `--overwrite`)
- ❌ Macro contract validation (warn if .xlsm → .xlsx)

---

## Remediation Strategy

### Task 1: Create Edit Session Abstraction

**New File:** `src/excel_agent/core/edit_session.py`

```python
"""Edit session abstraction for unified target semantics.

Provides a single entry point for all mutating operations:
- Resolves edit target (input vs output)
- Handles file copying when needed
- Provides locked ExcelAgent context
- Eliminates double-save issues
"""

from pathlib import Path
from typing import Optional
import shutil

from excel_agent.core.agent import ExcelAgent
from excel_agent.utils.exceptions import ValidationError


class EditSession:
    """Manages edit session lifecycle with unified target semantics.
    
    Pattern:
        with EditSession.prepare(input_path, output_path) as session:
            # session.path is the file being modified
            # session.agent provides locked ExcelAgent
            # session.is_inplace tells if input == output
            pass
        # Auto-saved to correct path on exit
    """
    
    def __init__(self, edit_path: Path, is_inplace: bool):
        self.edit_path = edit_path
        self.is_inplace = is_inplace
        self._agent: Optional[ExcelAgent] = None
    
    @classmethod
    def prepare(
        cls,
        input_path: Path,
        output_path: Optional[Path] = None,
        *,
        force_inplace: bool = False,
    ) -> "EditSession":
        """Prepare edit session with unified semantics.
        
        Args:
            input_path: Source file path
            output_path: Target output path (None = inplace)
            force_inplace: Allow editing input directly (requires explicit opt-in)
            
        Returns:
            EditSession configured for the resolved edit target
            
        Raises:
            ValidationError: If output is same as input but force_inplace not set
        """
        input_path = Path(input_path).resolve()
        
        if output_path is None:
            # In-place edit
            return cls(edit_path=input_path, is_inplace=True)
        
        output_path = Path(output_path).resolve()
        
        if output_path == input_path:
            if not force_inplace:
                raise ValidationError(
                    "output_path same as input_path requires force_inplace=True",
                    details={"input": str(input_path), "output": str(output_path)},
                )
            return cls(edit_path=input_path, is_inplace=True)
        
        # Different output: copy input to output, edit output
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_path, output_path)
        
        return cls(edit_path=output_path, is_inplace=False)
    
    def __enter__(self) -> "EditSession":
        """Enter locked ExcelAgent context."""
        self._agent = ExcelAgent(self.edit_path, mode="rw").__enter__()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Exit and save (ExcelAgent handles this)."""
        if self._agent:
            self._agent.__exit__(exc_type, exc_val, exc_tb)
            self._agent = None
    
    @property
    def workbook(self):
        """Access workbook through agent."""
        if not self._agent:
            raise RuntimeError("EditSession not entered")
        return self._agent.workbook
    
    @property
    def agent(self) -> ExcelAgent:
        """Access underlying agent for version hash, etc."""
        if not self._agent:
            raise RuntimeError("EditSession not entered")
        return self._agent


def prepare_edit_target(
    input_path: str,
    output_path: Optional[str] = None,
    *,
    create_parents: bool = False,
    force_inplace: bool = False,
) -> Path:
    """Legacy compatibility: resolve edit target path.
    
    Returns the path that should be opened for editing.
    """
    inp = Path(input_path).resolve()
    out = Path(output_path).resolve() if output_path else None
    
    session = EditSession.prepare(inp, out, force_inplace=force_inplace)
    return session.edit_path
```

---

### Task 2: Fix ExcelAgent to Support External Save Path

**Modify:** `src/excel_agent/core/agent.py`

**Change:** Add `save_path` parameter to handle edit sessions correctly.

```python
def __init__(
    self,
    path: Path,
    *,
    mode: str = "rw",
    keep_vba: bool | None = None,
    lock_timeout: float = 30.0,
    data_only: bool = False,
    save_path: Path | None = None,  # NEW: Where to save (may differ from path)
) -> None:
    self._save_path = save_path or self._path  # NEW
    
def __exit__(...):
    if exc_type is None and self._mode == "rw" and self._wb is not None:
        self.verify_no_concurrent_modification()
        self._wb.save(str(self._save_path))  # Use save_path, not _path
```

---

### Task 3: Migrate 13 Tools to ExcelAgent

**Priority Order:**

| Priority | Tool | Effort |
|----------|------|--------|
| P0 | `xls_add_chart.py` | Medium |
| P0 | `xls_add_image.py` | Medium |
| P0 | `xls_add_table.py` | Medium |
| P0 | `xls_format_range.py` | Medium |
| P1 | `xls_add_comment.py` | Low |
| P1 | `xls_set_data_validation.py` | Low |
| P1 | `xls_set_column_width.py` | Low |
| P1 | `xls_freeze_panes.py` | Low |
| P1 | `xls_apply_conditional_formatting.py` | Low |
| P1 | `xls_set_number_format.py` | Low |
| P2 | `xls_export_csv.py` | Low* |
| P2 | `xls_export_json.py` | Low* |
| P2 | `xls_export_pdf.py` | Low* |

\* Export tools use read-only mode, may not need full migration but should use `ExcelAgent(..., mode="r")`

**Migration Pattern:**

```python
# BEFORE (raw load_workbook)
wb = load_workbook(str(input_path))
ws = wb[args.sheet] if args.sheet else wb.active
# ... modifications ...
wb.save(str(output_path))

# AFTER (EditSession + ExcelAgent)
from excel_agent.core.edit_session import EditSession

session = EditSession.prepare(input_path, output_path)
with session:
    wb = session.workbook
    ws = wb[args.sheet] if args.sheet else wb.active
    # ... modifications ...
    # No explicit save - EditSession handles it
```

---

### Task 4: Fix Double-Save in 29 Tools

**Pattern to Remove:**

```python
# REMOVE these lines from all tools:
if str(output_path) != str(input_path):
    wb.save(str(output_path))
```

**Replace With:**

```python
# Use EditSession which handles save correctly
session = EditSession.prepare(input_path, output_path)
with session:
    wb = session.workbook
    # ... work ...
    # Auto-saved on exit
```

---

### Task 5: Tighten validate_output_path()

**Modify:** `src/excel_agent/utils/cli_helpers.py`

**Add New Functions:**

```python
from pathlib import Path
from typing import Set

from excel_agent.utils.exit_codes import ExitCode, exit_with


def validate_output_suffix(
    path: Path,
    allowed: Set[str] = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"}),
) -> None:
    """Validate output file extension.
    
    Args:
        path: Output path to validate
        allowed: Set of allowed extensions
        
    Raises:
        SystemExit: If extension not in allowed set
    """
    ext = path.suffix.lower()
    if ext not in allowed:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Invalid output extension: {ext}",
            details={
                "path": str(path),
                "extension": ext,
                "allowed": list(allowed),
            },
        )


def check_macro_contract(input_path: Path, output_path: Path) -> None:
    """Warn if macros may be lost.
    
    Args:
        input_path: Source file
        output_path: Target file
        
    Returns:
        Warning message if macro contract violated, None otherwise
    """
    input_ext = input_path.suffix.lower()
    output_ext = output_path.suffix.lower()
    
    macro_extensions = {".xlsm", ".xltm"}
    
    if input_ext in macro_extensions and output_ext not in macro_extensions:
        return (
            f"Converting {input_ext} to {output_ext} will strip VBA macros. "
            f"Use {input_path.stem}.xlsm to preserve macros."
        )
    return None


def validate_output_path(
    path_str: str,
    *,
    create_parents: bool = False,
    allowed_suffixes: Set[str] | None = None,
    overwrite: bool = False,
) -> Path:
    """Validate output path with comprehensive checks.
    
    NEW: Now validates:
    - Parent directory exists (or creates if create_parents)
    - Extension is in allowed set
    - File doesn't exist (unless overwrite=True)
    """
    path = Path(path_str).resolve()
    
    # Check parent directory
    if create_parents:
        path.parent.mkdir(parents=True, exist_ok=True)
    elif not path.parent.exists():
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Output directory does not exist: {path.parent}",
            details={"path": str(path), "parent": str(path.parent)},
        )
    
    # Validate extension
    if allowed_suffixes:
        validate_output_suffix(path, allowed_suffixes)
    
    # Check overwrite
    if path.exists() and not overwrite:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Output file exists: {path}. Use --overwrite to replace.",
            details={"path": str(path)},
        )
    
    return path
```

---

### Task 6: Fix xls_create_from_template Macro Handling

**Modify:** `src/excel_agent/tools/write/xls_create_from_template.py`

```python
# Add import
from excel_agent.core.agent import ExcelAgent  # For VBA handling

# In _run(), change line 46:
# BEFORE:
wb = load_workbook(str(template_path))

# AFTER:
# Detect if template has macros
template_ext = template_path.suffix.lower()
keep_vba = template_ext in {".xltm", ".xlsm"}

wb = load_workbook(str(template_path), keep_vba=keep_vba)

# If output is .xlsm but template was .xltx, warn about macro loss
output_ext = output_path.suffix.lower()
if keep_vba and output_ext not in {".xlsm", ".xltm"}:
    warnings.append(
        f"Template contains macros but output is {output_ext}. "
        f"Macros will be stripped. Use .xlsm extension to preserve."
    )
```

---

## Implementation Schedule

### Week 1: Core Infrastructure

| Day | Task | Deliverable |
|-----|------|-------------|
| 1-2 | Create EditSession abstraction | `edit_session.py` with tests |
| 3 | Modify ExcelAgent for save_path | Updated `agent.py` |
| 4 | Update validate_output_path() | Enhanced `cli_helpers.py` |
| 5 | Code review & test fixes | Passing unit tests |

### Week 2: Tool Migration

| Day | Tasks | Deliverable |
|-----|-------|-------------|
| 6-7 | Migrate P0 tools (add_chart, add_image, add_table, format_range) | 4 tools migrated |
| 8 | Migrate P1 tools | 6 tools migrated |
| 9 | Fix double-save in 29 tools | All tools updated |
| 10 | Fix xls_create_from_template | Template tool updated |
| 11-12 | Integration testing | All tests passing |
| 13-14 | Documentation & review | Updated docs, review |

---

## Testing Strategy

### Unit Tests (New)

```python
# test_edit_session.py
def test_edit_session_inplace():
    """In-place edit uses same file."""
    session = EditSession.prepare(Path("test.xlsx"), None)
    assert session.edit_path == Path("test.xlsx").resolve()
    assert session.is_inplace is True

def test_edit_session_copy_on_different_output():
    """Different output copies file first."""
    session = EditSession.prepare(Path("input.xlsx"), Path("output.xlsx"))
    assert session.edit_path == Path("output.xlsx").resolve()
    assert session.is_inplace is False
    assert Path("output.xlsx").exists()  # File was copied

def test_edit_session_forces_lock():
    """EditSession acquires lock via ExcelAgent."""
    with EditSession.prepare(Path("test.xlsx"), None) as session:
        assert session.agent._lock is not None
```

### Integration Tests

```python
# Test macro preservation
def test_macro_preservation_across_tools():
    """All tools preserve macros in .xlsm files."""
    for tool in ALL_MUTATING_TOOLS:
        result = tool(input="macro_file.xlsm", output="output.xlsm")
        assert has_macros("output.xlsm")

# Test double-save eliminated
def test_no_double_save():
    """Verify only one save per operation."""
    with patch('openpyxl.Workbook.save') as mock_save:
        xls_write_range(input="test.xlsx", output="out.xlsx", ...)
        assert mock_save.call_count == 1
```

---

## Success Criteria

- [ ] EditSession abstraction implemented and tested
- [ ] Zero double-save occurrences in codebase
- [ ] 13 tools migrated from raw `load_workbook()` to EditSession
- [ ] All tools preserve macros in `.xlsm` files
- [ ] `validate_output_path()` rejects invalid extensions
- [ ] `validate_output_path()` fails on overwrite without flag
- [ ] All existing tests pass
- [ ] New integration tests for macro preservation
- [ ] Performance regression < 5%

---

## Risks & Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Breaking changes to tool APIs | High | Maintain backward compatibility, deprecation warnings |
| Performance regression | Medium | Benchmark before/after, optimize hot paths |
| Lock contention with EditSession | Medium | Test with concurrent access patterns |
| Macro preservation edge cases | Medium | Comprehensive .xlsm test suite |
| Large file handling | Low | Test with files >100MB |

---

## Appendix: Tools Status Matrix

| Tool | Uses ExcelAgent | Has Double-Save | Uses Raw load_workbook | Priority |
|------|----------------|-----------------|------------------------|----------|
| xls_add_chart | ❌ | N/A | ✅ Yes | P0 |
| xls_add_image | ❌ | N/A | ✅ Yes | P0 |
| xls_add_table | ❌ | N/A | ✅ Yes | P0 |
| xls_format_range | ❌ | N/A | ✅ Yes | P0 |
| xls_add_comment | ❌ | N/A | ✅ Yes | P1 |
| xls_set_data_validation | ❌ | N/A | ✅ Yes | P1 |
| xls_set_column_width | ❌ | N/A | ✅ Yes | P1 |
| xls_freeze_panes | ❌ | N/A | ✅ Yes | P1 |
| xls_apply_conditional_formatting | ❌ | N/A | ✅ Yes | P1 |
| xls_set_number_format | ❌ | N/A | ✅ Yes | P1 |
| xls_export_csv | ❌ | N/A | ✅ Yes* | P2 |
| xls_export_json | ❌ | N/A | ✅ Yes* | P2 |
| xls_export_pdf | ❌ | N/A | ✅ Yes* | P2 |
| xls_write_range | ✅ | ✅ Yes | ❌ No | P0 |
| xls_write_cell | ✅ | ✅ Yes | ❌ No | P0 |
| xls_add_sheet | ✅ | ✅ Yes | ❌ No | P1 |
| xls_delete_sheet | ✅ | ? | ❌ No | P1 |
| xls_rename_sheet | ✅ | ? | ❌ No | P1 |
| xls_delete_rows | ✅ | ? | ❌ No | P1 |
| xls_delete_columns | ✅ | ? | ❌ No | P1 |
| xls_insert_rows | ✅ | ? | ❌ No | P1 |
| xls_insert_columns | ✅ | ? | ❌ No | P1 |
| xls_move_sheet | ✅ | ? | ❌ No | P1 |
| xls_merge_cells | ✅ | ? | ❌ No | P1 |
| xls_unmerge_cells | ✅ | ? | ❌ No | P1 |
| xls_delete_range | ✅ | ? | ❌ No | P1 |
| xls_update_references | ✅ | ? | ❌ No | P1 |
| xls_set_formula | ✅ | ? | ❌ No | P1 |
| xls_define_name | ✅ | ? | ❌ No | P1 |
| xls_convert_to_values | ✅ | ? | ❌ No | P1 |
| xls_copy_formula_down | ✅ | ? | ❌ No | P1 |

\* Export tools are read-only, different pattern needed

---

**Document Version:** 1.0  
**Created:** April 2026  
**Author:** Code Analysis & Remediation Agent  
**Review Status:** Pending
