$ ls -l Test-plan.md skills/excel-tools/*
-rw-rw-r-- 1 pete pete 11209 Apr 10 06:52 Test-plan.md
-rw-rw-r-- 1 pete pete  7023 Apr  9 22:42 skills/excel-tools/SKILL.md
-rw-rw-r-- 1 pete pete    70 Apr  9 22:48 skills/excel-tools/requirements.txt

skills/excel-tools/assets:
total 16
-rw-rw-r-- 1 pete pete 4784 Apr  9 22:48 template.xlsx
-rw-rw-r-- 1 pete pete 6031 Apr  9 22:50 template_with_data.xlsx

skills/excel-tools/references:
total 28
-rw-rw-r-- 1 pete pete 11940 Apr  9 22:45 tool-reference.md
-rw-rw-r-- 1 pete pete  6407 Apr  9 22:46 troubleshooting.md
-rw-rw-r-- 1 pete pete  5845 Apr  9 22:43 workflow-patterns.md

skills/excel-tools/scripts:
total 8
-rw-rw-r-- 1 pete pete 2957 Apr  9 22:48 batch_process.py
-rw-rw-r-- 1 pete pete 2337 Apr  9 22:47 create_workbook.py

---

# 🔍 Comprehensive Review & Critique: `SKILL.md`, References, Scripts & `Test-plan.md`

## Executive Summary

The skill documentation suite (`SKILL.md`, `tool-reference.md`, `troubleshooting.md`, `workflow-patterns.md`), helper scripts (`batch_process.py`, `create_workbook.py`), and `Test-plan.md` form a professional, well-structured body of work. The documentation is clearly written with good examples and proper cross-referencing. However, **validation against the actual source code reveals 4 critical missing implementations** that undermine the core "53 tools" claim, along with several inconsistencies and code quality concerns that should be addressed before production deployment.

---

## 1. SKILL.md — Review & Critique

### ✅ Strengths
- **Clear trigger scope**: The "When to Use This Skill" section accurately describes the project's capabilities
- **Architecture diagram**: The ASCII diagram correctly summarizes the 10 tool categories
- **5 Key Principles**: Clone-Before-Edit, Token Protection, Formula Integrity, JSON-Native, Headless — all confirmed in source code
- **Core Workflow**: The 5-step pipeline (Clone → Read → Modify → Calculate → Export) is accurate and well-demonstrated
- **Token Scopes table**: All 7 scopes (`sheet:delete`, `sheet:rename`, `range:delete`, `formula:convert`, `macro:remove`, `macro:inject`, `structure:modify`) match `token_manager.py` lines 32–42 exactly
- **Exit Codes table**: Codes 0–5 match `exit_codes.py` lines 27–35 exactly
- **Installation**: Simple and correct (`pip install excel-agent-tools`)
- **Referenced Resources**: Good cross-linking to the references/ and scripts/ directories

### 🔴 Critical Issues

**C1: "53 tools" claim is misleading**
The SKILL.md frontmatter metadata states `total-tools: "53"` and the description says "53 tools designed for AI agents." While 53 console_scripts are registered in `pyproject.toml`, **4 tool source files are missing**:
- `xls-detect-errors` — no `src/excel_agent/tools/formulas/xls_detect_errors.py`
- `xls-convert-to-values` — no corresponding source file
- `xls-copy-formula-down` — no corresponding source file
- `xls-define-name` — no corresponding source file

**Only 49 of 53 tools are actually implementable.** This should be corrected to `total-tools: "49"` with a note about the 4 pending implementations.

**C2: Quick Examples reference non-existent tool behaviors**
- The "Safe Sheet Deletion" example (line 178–181) shows `xls-delete-sheet --output report.xlsx` — this implies the tool accepts `--output`, but many structure tools may not have this parameter consistently.
- The "Calculate Formulas" example (line 172–175) implies `xls-recalculate` works end-to-end, but its underlying formula engine depends on the `formulas` library which has known limitations with complex cross-sheet references.

### 🟡 Minor Issues

**M1: Metadata `coverage: ">90%"` is unverifiable**
No coverage badge, coverage report, or CI configuration exists to substantiate this claim. It should either reference a specific coverage artifact or be removed.

**M2: `allowed-tools` list is incomplete**
The frontmatter lists only `bash` and `python`. Given that the tool also provides Python helper scripts and requires LibreOffice for Tier 2, the allowed-tools metadata could be more descriptive.

**M3: "See Also" section references non-existent docs**
Line 190 references `CLAUDE.md` and `Project_Architecture_Document.md` which exist at the repo root, but line 190 also says "Full documentation in project `/docs/` folder" — the `docs/` directory does not exist in the repository.

---

## 2. `tool-reference.md` — Review & Critique

### ✅ Strengths
- **Complete catalog**: All 53 tools are listed with CLI syntax and expected JSON output
- **Consistent format**: Each tool follows Purpose → CLI → Output structure
- **Governance warnings**: Token-required tools are clearly marked with ⚠️
- **Export `--outfile` caveat**: Correctly documents the argparse conflict reason
- **Format Notes section**: Useful reference for date/number/formula handling

### 🔴 Critical Issues

**C1: JSON response envelope inconsistency**
The tool reference shows responses with `"status": "success"` at the top level but no `"exit_code"` field. However, the actual `build_response()` in `json_io.py` (line 47) includes `exit_code`, `timestamp`, `workbook_version` in every response. The documented JSON schema is incomplete — it should show the full envelope.

**C2: 4 missing tools documented as if they work**
The following tools are fully documented with CLI syntax and expected output, but their source files don't exist:
- `xls-detect-errors` (lines 484–496)
- `xls-convert-to-values` (lines 499–501)
- `xls-copy-formula-down` (lines 503–504)
- `xls-define-name` (lines 506–507)

These entries should be clearly marked as ⚠️ **NOT YET IMPLEMENTED** or removed until implemented.

**C3: `xls-set-formula` output not documented**
The tool reference shows CLI syntax for `xls-set-formula` but no expected JSON output, unlike most other tools.

### 🟡 Minor Issues

**M1: Structure tools section header says "⚠️ Token Required"**
This implies ALL structure tools require tokens, but `xls-add-sheet`, `xls-insert-rows`, `xls-insert-columns`, and `xls-move-sheet` do NOT require tokens — only `xls-delete-sheet`, `xls-rename-sheet`, `xls-delete-rows`, and `xls-delete-columns` do. The header is misleading.

**M2: `xls-get-formula` output format differs from `xls-set-formula`**
The `xls-get-formula` output shows `"references": ["B1:B10"]` — this implies automatic reference extraction, but the actual implementation may not parse formula references into an array.

---

## 3. `troubleshooting.md` — Review & Critique

### ✅ Strengths
- **Comprehensive coverage**: 13 issues documented with symptoms, causes, and solutions
- **Accurate advice**: The `--outfile` vs `--output` confusion, chunked JSONL parsing, and token validation issues are all real and correctly explained
- **Practical solutions**: Each issue includes copy-pasteable bash commands
- **Debugging tips section**: Useful `jq` patterns and verification commands
- **Python version check**: Correctly notes `>= 3.12` requirement

### 🟡 Minor Issues

**M1: File Lock retry loop has a syntax issue**
Line 17: `for i in 0.5 1 2 4` — bash `for` loops iterate over words, so `$i` will be `0.5`, `1`, `2`, `4` (not fractions). The `sleep $i` will work, but the variable naming suggests seconds, and `sleep 0.5` works on most Linux systems. However, the "exponential backoff" label is misleading — the values `0.5, 1, 2, 4` are exponential, but the loop structure doesn't demonstrate a standard backoff pattern (no jitter, no max retries check).

**M2: "Debugging Tips → Enable Verbose Output" section is a dead end**
Lines 280–284 acknowledge that most tools don't have verbose mode and suggest checking JSON warnings. This is honest but not helpful — consider documenting `PYTHONVERBOSE=1` or `python -v` as alternatives for deeper debugging.

**M3: References `docs/API.md` which doesn't exist**
Line 324: "Review API docs: `docs/API.md`" — this file is not present in the repository.

**M4: References `src/excel_agent/utils/exit_codes.py`**
Line 325 — this path is correct for the source repo, but if a user installed via pip, they won't have access to the source file. Should reference the SKILL.md exit codes table instead.

---

## 4. `workflow-patterns.md` — Review & Critique

### ✅ Strengths
- **8 well-chosen patterns** covering the most common agent workflows
- **Pattern 3 (Safe Structural Edit)** is particularly well-designed, showing the denial-with-guidance loop
- **Pattern 6 (Macro Safety Audit)** correctly shows the double-token requirement
- **Error Handling Pattern** with the Python `subprocess` example is practical and reusable
- **Python Integration Pattern** provides a clean `run_tool()` wrapper

### 🔴 Critical Issues

**C1: Python Integration Pattern has a bug**
Lines 193–218: The `run_tool()` wrapper does `json.loads(result.stdout)` unconditionally (line 204), even when `result.returncode != 0`. If a tool fails with a malformed JSON (which shouldn't happen per the design, but could in edge cases), this will raise an unhandled `JSONDecodeError`. The Error Handling Pattern (lines 164–189) correctly shows checking `returncode` first, but the Integration Pattern contradicts this best practice.

**C2: Pattern 1 uses `--output "$CLONE"` for `xls-write-range`**
Line 18: `xls-write-range --input "$CLONE" --output "$CLONE"` — while this works for in-place modification, the SKILL.md Step 3 example (line 84) does NOT include `--output`. This inconsistency could confuse agents.

### 🟡 Minor Issues

**M1: Pattern 4 (Batch Processing) doesn't handle errors per-file**
The bash loop processes each file but has no error handling — if one file fails, it continues to the next without logging which failed or why.

**M2: No pattern for "Create from scratch with formulas"**
The most common use case for a new workbook is likely creating one with formulas, but no pattern demonstrates `xls-set-formula` after creation.

---

## 5. `batch_process.py` — Review & Critique

### ✅ Strengths
- Clean argparse interface with `--input`, `--output`, `--operations`
- Correctly clones before processing (respects clone-before-edit principle)
- Supports both directory and glob pattern inputs
- Prints JSON results for machine-readable output

### 🟡 Issues

**M1: `run_tool()` doesn't check return codes**
Line 20–25: `run_tool()` calls `subprocess.run()` but never checks `result.returncode`. If a tool exits with code 1–5, the stdout may still contain valid JSON with error information, but the caller won't know the operation failed unless it checks the JSON `status` field.

**M2: Potential `--output` double-write issue**
Lines 57–58: `params["input"] = clone_path` and `params["output"] = clone_path` — this works for tools that support `--output`, but some tools (like `xls-delete-sheet` in the SKILL.md examples) may not accept `--output`. The script would pass an unexpected argument.

**M3: Operations JSON structure is undocumented**
The `--operations` parameter expects `[{"tool": "...", "params": {...}}]` but this format isn't documented anywhere — not in the script's `--help`, not in SKILL.md, not in README.md.

**M4: No dry-run mode**
For a batch processing tool, a `--dry-run` flag that shows what would happen without executing would be valuable.

---

## 6. `create_workbook.py` — Review & Critique

### ✅ Strengths
- Simple, focused script for a single purpose
- Correctly chains `xls-create-new` → `xls-write-range`
- Properly handles the `--sheets` comma-separated format

### 🟡 Issues

**M1: Error handling checks stderr but tools write to stdout**
Line 37: `return {"status": "error", "error": result.stderr}` — but excel-agent-tools write error JSON to stdout (by design), not stderr. This means error messages will be lost.

**M2: No `--recalculate` option**
After writing data, it would be natural to recalculate formulas, but the script doesn't offer this option.

---

## 7. `Test-plan.md` — Review & Critique

### ✅ Strengths
- **Part 1 (Production Readiness)** is well-reasoned and correctly categorizes Phase 14 items as optional
- **Part 2 (E2E Test Plan)** is professionally structured with 5 comprehensive scenarios covering all tool categories
- **Security & Edge Cases section** (Section 4) is thorough — covers token expiry, nonce replay, scope mismatch, concurrent locks, path traversal, malformed JSON, missing LibreOffice
- **Performance SLAs** (Section 5) are specific and measurable
- **QA Pass/Fail Criteria** (Section 6) are rigorous and well-defined
- **Execution Strategy** (Section 7) is practical for CI integration

### 🔴 Critical Issues

**C1: Scenario C references 3 non-existent tools**
Scenario C ("Formula Engine & Error Recovery Pipeline") tests 6 formula tools, but 3 of them have no source files:
- `xls-detect-errors` (Step 3, line 83)
- `xls-copy-formula-down` (Step 4, line 84)
- `xls-convert-to-values` (Step 5, line 86)

**These test cases CANNOT execute.** This means the QA pass/fail criterion #1 ("All 53 CLI entry points execute without ModuleNotFoundError") will **definitely fail**.

**C2: Scenario B references `xls-update-references` which needs validation**
Step 3 (line 69): `xls-update-references --updates '[...]'` — while the source file exists, the `--updates` JSON format isn't fully documented. The Test-plan should specify the exact JSON structure.

**C3: Performance benchmarks depend on non-existent files**
The SLA table references `large_dataset.xlsx` (500k rows) and `complex_formulas.xlsx` (10 sheets, 1k+ formulas) as test fixtures. These don't exist, and the benchmark scripts referenced in the Master Execution Plan also don't exist.

**D1: Inappropriate URL at end of document**
Line 163: `# https://chat.qwen.ai/s/2503e7d5-e7b7-4b82-99f1-15a01453b0b1?fev=0.2.36` — This is a chat session URL, not a documentation reference. It should be removed for professionalism.

### 🟡 Minor Issues

**M1: Performance SLA for Large Write is questionable**
Line 135: "100k row JSON array via `xls-write-range` | `< 5.0s`" and the assertion says "`impact.cells_modified == 1M`" — but 100k rows × 10 columns = 1M cells, not 100k rows alone. The measurement and assertion should be consistent.

**M2: Test environment prerequisites may be incomplete**
The prerequisite list mentions `sample.xlsx`, `macros.xlsm`, `large_dataset.xlsx`, `template.xltx` as fixtures but these don't exist in the repository. The `conftest.py` generates some fixtures programmatically, but not `large_dataset.xlsx` (500k rows) or `template.xltx`.

**M3: No negative test for invalid token format**
The security section tests expiry, replay, and scope mismatch, but doesn't test what happens with a completely invalid/malformed token string (e.g., random bytes, empty string, extremely long string).

**M4: No test for extremely large cell values or deeply nested formulas**
The performance section focuses on large datasets but doesn't test edge cases like a single cell with a formula referencing 10k+ cells, or a cell containing 50k characters of text.

---

## 8. `requirements.txt` — Review & Critique

Single line: `excel-agent-tools>=1.0.0`

This is correct for the skill's scripts that wrap the CLI tools, but it's worth noting that:
- The version constraint `>=1.0.0` means the scripts will work with any future major version, which could break if the CLI interface changes
- There's no upper bound pinning (e.g., `<2.0.0`)

---

## 9. Consolidated Findings Summary

| Category | Severity | Count |
|----------|----------|-------|
| Missing tool implementations | 🔴 Critical | 4 |
| Documentation propagates false claims | 🔴 Critical | 3 |
| Test scenarios reference non-existent tools | 🔴 Critical | 3 |
| JSON schema inconsistencies | 🟡 Medium | 3 |
| Script bugs / missing error handling | 🟡 Medium | 5 |
| Missing docs/ files referenced | 🟡 Medium | 5 |
| Unlisted `requests` dependency | 🟡 Medium | 1 |
| Minor inaccuracies | 🟢 Low | 6 |

### Priority Recommendations

1. **[P0]** Implement the 4 missing formula tools or update all documentation to reflect 49 tools
2. **[P0]** Remove the chat URL from `Test-plan.md` line 163
3. **[P1]** Fix `batch_process.py` error handling to check return codes
4. **[P1]** Fix `create_workbook.py` to read errors from stdout, not stderr
5. **[P1]** Add `requests` to `pyproject.toml` dependencies or make the import conditional
6. **[P1]** Mark the 4 missing tools as ⚠️ NOT YET IMPLEMENTED in `tool-reference.md`
7. **[P2]** Create the 5 missing `docs/*.md` files or remove references to them
8. **[P2]** Create performance benchmark scripts and test fixtures
9. **[P2]** Fix `workflow-patterns.md` Python Integration Pattern to check return codes before parsing JSON
10. **[P2]** Update `SKILL.md` metadata to `total-tools: "49"` and add a note about pending implementations

