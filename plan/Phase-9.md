Now planning Phase 9 per the Master Execution Plan.

## Phase 1: ANALYZE — Macro Safety Requirements

From the Master Execution Plan and opus-3 research synthesis:

**Phase 9 scope**: 5 macro tools, 3 days, 7 files
- `xls_has_macros` — Boolean detection
- `xls_inspect_macros` — Extract source code, metadata
- `xls_validate_macro_safety` — Security analysis via oletools
- `xls_remove_macros` — Strip VBA project (token-gated)
- `xls_inject_vba_project` — Inject pre-scanned.bin (token-gated)

**Critical constraints from research**:
1. oletools 0.60.2 is inactive for 12 months but battle-tested by Palo Alto, Cuckoo, VirusTotal. Must isolate behind `MacroAnalyzer` Protocol
2. vbaProject.bin is opaque OLE COM binary. Cannot construct from scratch. Must extract from trusted source and re-inject
3. `detect_autoexec` identifies auto-run triggers (Workbook_Open, Auto_Open, etc.)
4. `detect_suspicious` flags malware patterns (CreateObject, Shell, registry writes)
5. Injection MUST be preceded by `scan_risk()` — never inject unscanned binary

**Governance integration**:
- Scopes: `macro:remove` and `macro:inject` (from Master Plan)
- Both require HMAC-SHA256 tokens with TTL
- All operations logged to audit trail, but macro source never logged (secret protection)
- Pre-flight check: validate.xlsm extension, verify vbaProject.bin exists

**Technical dependencies**:
- Phase 3: TokenManager and AuditTrail must be complete
- Phase 1: ExcelAgent context manager for safe file handling
- oletools provides: OleFileIO, VBA_Parser, detect_autoexec, detect_suspicious

## Phase 2: PLAN — Phase 9 Detailed Execution Roadmap

### Research Summary (Pre-Implementation Validation)

**oletools API surface (v0.60.2 confirmed)**:
- `from oletools.olevba import VBA_Parser, detect_autoexec, detect_suspicious`
- `VBA_Parser(filename)` → `vbaparser.detect_vba_macros()` returns bool
- `vbaparser.extract_macros()` → yields (filename, stream_path, vba_filename, vba_code)
- `vbaparser.analyze_macros()` → returns list of (type, keyword, description)
- `detect_autoexec(vba_code)` → list of (keyword, description) tuples
- `detect_suspicious(vba_code)` → list of (keyword, description) tuples
- Supports both.xlsm (OpenXML) and.xls (OLE) formats
- XLM macro detection via `vbaparser.xlm_macros` attribute

**vbaProject.bin handling**:
- Location in.xlsm: `xl/vbaProject.bin`
- openpyxl preserves it when `keep_vba=True` in `load_workbook()`
- Extraction: unzip.xlsm, copy `xl/vbaProject.bin`
- Injection: unzip target, replace bin, re-zip, ensure `[Content_Types].xml` includes VBA content type

**Security patterns to detect**:
- Auto-exec: `Auto_Open`, `Workbook_Open`, `Document_Open`, `Auto_Close`
- Suspicious: `CreateObject`, `Shell`, `WScript.Shell`, `Environ`, `URLDownloadToFile`, registry access, `Chr(` obfuscation

### File Structure

| # | Path | Purpose | Lines |
|---|---|---|---|
| 1 | `src/excel_agent/core/macro_handler.py` | MacroAnalyzer Protocol + Oletools implementation | ~350 |
| 2 | `src/excel_agent/tools/macros/__init__.py` | Package init | 5 |
| 3 | `src/excel_agent/tools/macros/xls_has_macros.py` | Boolean detection tool | ~80 |
| 4 | `src/excel_agent/tools/macros/xls_inspect_macros.py` | Extract source + metadata | ~150 |
| 5 | `src/excel_agent/tools/macros/xls_validate_macro_safety.py` | Security analysis | ~180 |
| 6 | `src/excel_agent/tools/macros/xls_remove_macros.py` | Strip VBA (token-gated) | ~120 |
| 7 | `src/excel_agent/tools/macros/xls_inject_vba_project.py` | Inject.bin (token-gated, pre-scan) | ~200 |
| 8 | `tests/unit/test_macro_handler.py` | Unit tests for MacroAnalyzer | ~250 |
| 9 | `tests/unit/test_macro_tools.py` | Unit tests for 5 tools | ~300 |
| 10 | `tests/integration/test_macro_workflow.py` | E2E macro workflows | ~200 |
| 11 | `tests/fixtures/macro_test.xlsm` | Generated test file with macros | — |

### Implementation Details

**File 1: macro_handler.py — The Abstraction Layer**

```python
from typing import Protocol, runtime_checkable
from dataclasses import dataclass
from pathlib import Path

@dataclass
class MacroModule:
    stream_path: str
    vba_filename: str
    code: str
    line_count: int

@dataclass
class AutoExecTrigger:
    keyword: str
    description: str
    module: str

@dataclass
class SuspiciousKeyword:
    keyword: str
    description: str
    module: str
    line_number: int | None

@dataclass
class RiskReport:
    has_macros: bool
    has_auto_exec: bool
    has_suspicious: bool
    auto_exec_triggers: list[AutoExecTrigger]
    suspicious_keywords: list[SuspiciousKeyword]
    risk_level: str # "low" | "medium" | "high" | "critical"
    xlm_macros_detected: bool

@runtime_checkable
class MacroAnalyzer(Protocol):
    def has_macros(self, path: Path) -> bool:...
    def extract_modules(self, path: Path) -> list[MacroModule]:...
    def detect_auto_exec(self, path: Path) -> list[AutoExecTrigger]:...
    def detect_suspicious(self, path: Path) -> list[SuspiciousKeyword]:...
    def scan_risk(self, path: Path) -> RiskReport:...
    def extract_vba_project_bin(self, path: Path, output_dir: Path) -> Path:...
    def inject_vba_project_bin(self, target: Path, vba_bin: Path, output: Path) -> None:...

class OletoolsMacroAnalyzer:
    """Production implementation using oletools 0.60.2"""
    # Implements all Protocol methods
    # Wraps VBA_Parser, detect_autoexec, detect_suspicious
    # Handles both.xlsm and.xls
    # Catches oletools exceptions, converts to ExcelAgentError
```

**File 3: xls_has_macros.py**
- Input: workbook path
- Output: `{"has_macros": bool, "vba_project_present": bool, "xlm_macros": bool}`
- Uses `MacroAnalyzer.has_macros()`
- No token required (read-only)
- Exit code 0 always (detection, not validation)

**File 4: xls_inspect_macros.py**
- Input: workbook path, optional `--include-code` flag
- Output: List of modules with metadata, optionally full source
- Security: If `--include-code` is used, output is NOT written to audit trail (secret protection)
- Returns line counts, module names, stream paths
- Warns if XLM macros detected (legacy Excel 4.0 macros)

**File 5: xls_validate_macro_safety.py**
- Input: workbook path or.bin path
- Output: Full RiskReport as JSON
- Calls `scan_risk()` which runs:
  1. `detect_autoexec` on each module
  2. `detect_suspicious` on each module
  3. Checks for XLM macros
  4. Assigns risk_level:
     - critical: auto_exec + suspicious
     - high: suspicious only
     - medium: auto_exec only
     - low: macros present but no triggers
- Exit code 0 if low/medium, exit code 1 if high/critical (validation failure)

**File 6: xls_remove_macros.py**
- Requires: `macro:remove` token
- Pre-flight: Check has_macros, if false return early success
- Implementation:
  1. Validate token via TokenManager
  2. Load workbook with openpyxl `keep_vba=False` (strips vbaProject.bin)
  3. Save to output path
  4. Verify removal via `has_macros()` on output
  5. Audit log: operation, file hash before/after, token_id (never log code)
- Impact: `{macros_removed: int, modules_deleted: list[str]}`
- Exit code 0 on success, 4 on bad token, 1 on validation failure

**File 7: xls_inject_vba_project.py**
- Requires: `macro:inject` token
- Hard pre-condition: MUST run `xls_validate_macro_safety` on the.bin FIRST
- Input: target workbook, vba_project.bin path, output path
- Implementation:
  1. Validate token
  2. Run `scan_risk()` on the.bin file (extract to temp, analyze)
  3. If risk_level is high/critical, deny unless `--force` flag present
  4. If denied, return exit code 1 with guidance: "Run xls_validate_macro_safety --bin <path> to review risks"
  5. If approved: unzip target.xlsm, replace xl/vbaProject.bin, update [Content_Types].xml if needed, re-zip
  6. Verify injection via `has_macros()` on output
  7. Audit log: operation, source bin hash, target hash before/after, risk_level
- Security: Never accept.bin from untrusted source without scan
- Exit codes: 0 success, 1 risk denial, 4 bad token

### Testing Strategy

**Unit tests** (`test_macro_handler.py`):
- Mock VBA_Parser to test Protocol without real.xlsm files
- Test risk_level assignment logic
- Test XLM detection
- Test exception handling for corrupt files

**Unit tests** (`test_macro_tools.py`):
- Test has_macros on clean and macro-enabled workbooks
- Test inspect with and without --include-code
- Test validate returns correct risk levels
- Test remove without token fails (exit 4)
- Test remove with token succeeds
- Test inject without pre-scan fails
- Test inject with high-risk bin requires --force

**Integration tests** (`test_macro_workflow.py`):
- Full workflow: create.xlsm with macro → inspect → validate → remove → verify clean → inject → verify restored
- Test audit trail contains entries but no source code
- Test token expiration

### Exit Criteria Checklist

| # | Criterion | Validation Method |
|---|---|---|
| 1 | MacroAnalyzer Protocol defined with 6 methods | Type check |
| 2 | OletoolsMacroAnalyzer implements Protocol | isinstance check |
| 3 | `has_macros` detects.xlsm and.xls | Test with fixtures |
| 4 | `extract_modules` returns code and metadata | Verify line counts |
| 5 | `detect_auto_exec` finds Workbook_Open | Test with auto-exec fixture |
| 6 | `detect_suspicious` finds Shell/CreateObject | Test with suspicious fixture |
| 7 | `scan_risk` assigns 4 risk levels correctly | Unit tests for each level |
| 8 | XLM macro detection works | Test with.xls containing XLM |
| 9 | `xls_has_macros` exit 0 always | Integration test |
| 10 | `xls_inspect_macros` respects --include-code | Verify output |
| 11 | Audit trail excludes source code | Check.jsonl file |
| 12 | `xls_validate_macro_safety` exit 1 on high/critical | Test |
| 13 | `xls_remove_macros` requires token | Test without token → exit 4 |
| 14 | Remove strips vbaProject.bin completely | Verify has_macros false after |
| 15 | `xls_inject_vba_project` requires token | Test without token → exit 4 |
| 16 | Inject runs pre-scan automatically | Verify scan_risk called |
| 17 | Inject denies high-risk without --force | Test |
| 18 | Inject succeeds with --force | Test |
| 19 | All operations logged to audit | Check audit entries |
| 20 | No oletools exceptions leak to stdout | Test with corrupt file |

### Key Design Decisions

1. **Protocol abstraction is mandatory** — oletools is inactive. If a CVE emerges or a fork appears, swapping implementations requires zero tool changes. The Protocol is the contract.

2. **Pre-scan on inject is non-negotiable** — The blueprint's Q1 answer: "Never inject an unscanned macro binary." This is enforced in code, not documentation. The tool calls `scan_risk()` internally before touching the target workbook.

3. **Source code never enters audit trail** — Macros may contain passwords, API keys, or proprietary logic. Audit logs record hashes and metadata only. The `--include-code` flag on inspect triggers a warning that output should not be persisted.

4. **Risk levels drive automation** — AI agents can programmatically decide: low/medium → proceed, high → require human review, critical → block. The JSON response includes `risk_level` and `suggestion` fields for this.

5. **XLM macros treated as high risk** — Excel 4.0 macros are obsolete, heavily used by malware, and invisible to most users. Detection automatically elevates risk_level to at least "high".

6. **Dual format support** — Both.xlsm (OpenXML) and.xls (OLE) are supported because legacy workbooks still circulate in enterprises. VBA_Parser handles both transparently.

### Dependencies and Risks

**Must be complete before Phase 9**:
- Phase 3: TokenManager with `macro:remove` and `macro:inject` scopes
- Phase 3: AuditTrail with `log_operation()` method
- Phase 1: ExcelAgent for safe file handling

**Risks**:
- oletools may fail on heavily obfuscated VBA. Mitigation: catch exceptions, return "scan_failed" risk_level, require manual review
- Injecting.bin into.xlsx (not.xlsm) will corrupt file. Mitigation: validate target extension before unzip
- Concurrent inject operations could corrupt shared temp dirs. Mitigation: use unique temp directories per operation

### Effort Estimate
- macro_handler.py: 1 day (Protocol + oletools wrapper + error handling)
- 5 tools: 1.5 days (each ~100 lines, mostly glue code)
- Tests: 0.5 days (fixtures + unit + integration)

Total: 3 days, aligns with Master Plan.

## Phase 3: VALIDATE — Confirmation Checkpoint

Before I generate the full Phase 9 implementation code, confirm:

1. **Risk level thresholds**: Current design blocks high/critical by default. Should medium risk also require `--force`, or is auto-exec without suspicious keywords acceptable for automated injection

2. **XLM handling**: Should XLM macros be automatically removed during `xls_remove_macros`, or treated separately. Current plan removes all VBA including XLM

3. **.bin validation**: Should `xls_inject_vba_project` accept only.bin files extracted by our own `extract_vba_project_bin`, or any.bin. Current plan accepts any but always scans first

4. **Audit detail**: Should we log the list of detected suspicious keywords in audit trail (metadata only, not code), or only risk_level. Current plan logs keywords for forensics

