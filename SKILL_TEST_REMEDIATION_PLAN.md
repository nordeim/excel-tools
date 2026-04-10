# QA Remediation Plan: Review_Report_Skill_and_Test.md

**Validation Date:** April 10, 2026  
**Status:** 6 Critical/High Issues Confirmed - Remediation Required  

---

## Executive Summary

Meticulous validation of QA feedback against the actual codebase has revealed **6 confirmed issues** requiring remediation. The 4 missing formula tools were already implemented in the previous QA remediation, but documentation and script bugs remain.

---

## ✅ Validated Findings

### Issue 1: P0 - Chat URL in Test-plan.md

**Status:** ✅ **CONFIRMED**  
**File:** `Test-plan.md:163`  
**Finding:** Inappropriate chat URL present

```bash
$ grep -n "chat.qwen" Test-plan.md
163:# https://chat.qwen.ai/s/2503e7d5-e7b7-4b82-99f1-15a01453b0b1?fev=0.2.36
```

**Impact:** Unprofessional, should be removed  
**Fix:** Delete line 163

---

### Issue 2: P1 - batch_process.py Return Code Checking

**Status:** ✅ **CONFIRMED**  
**File:** `skills/excel-tools/scripts/batch_process.py:20-25`

**Current Code:**
```python
def run_tool(tool: str, **kwargs) -> dict:
    """Run an excel-agent tool."""
    cmd = [f"xls-{tool}"]
    for key, value in kwargs.items():
        cmd.extend([f"--{key.replace('_', '-')}", str(value)])
    result = subprocess.run(cmd, capture_output=True, text=True)
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError:
        return {"status": "error", "error": result.stderr or result.stdout}
```

**Problem:** Never checks `result.returncode`. Tool can exit with error code but still return valid JSON.

**Fix:** Add return code checking before JSON parsing

---

### Issue 3: P1 - create_workbook.py Error Source

**Status:** ✅ **CONFIRMED**  
**File:** `skills/excel-tools/scripts/create_workbook.py:37`

**Current Code:**
```python
if result.returncode != 0:
    return {"status": "error", "error": result.stderr}  # WRONG: Should read stdout
```

**Problem:** excel-agent-tools write JSON errors to stdout, not stderr.

**Fix:** Read from `result.stdout` instead of `result.stderr`

---

### Issue 4: P2 - Missing requests Dependency

**Status:** ✅ **CONFIRMED**

**Problem:** `requests` library not explicitly declared, may be required by:
- oletools (for URL downloads)
- Future webhook audit backends

**Fix:** Add `requests>=2.32.0` to pyproject.toml dependencies

---

### Issue 5: P2 - SKILL.md Metadata Inaccurate

**Status:** ✅ **CONFIRMED**  
**File:** `skills/excel-tools/SKILL.md`

**Current:**
```yaml
total-tools: "53"
coverage: ">90%"
```

**Problem:** Documentation claims 53 tools, but:
- 53 entry points registered
- All 53 now implemented (post-remediation)
- Coverage claim unverified

**Fix:** 
- Verify all 53 tools implemented ✅ (Done in previous remediation)
- Update coverage claim to be verifiable or remove

---

### Issue 6: P3 - Missing docs/API.md Reference

**Status:** ✅ **CONFIRMED**

**Finding:** `troubleshooting.md` references `docs/API.md` which doesn't exist at that path (it's in root `docs/API.md`).

**Fix:** Update path references

---

## 📋 Remediation Tasks

### Sprint 1: Critical Fixes (Day 1)

| Task | File | Action | Time |
|------|------|--------|------|
| 1.1 | Test-plan.md | Remove chat URL (line 163) | 5 min |
| 1.2 | batch_process.py | Add return code checking | 15 min |
| 1.3 | create_workbook.py | Fix stderr→stdout error reading | 10 min |
| 1.4 | pyproject.toml | Add requests dependency | 5 min |

**Sprint 1 Total:** 35 minutes

### Sprint 2: Documentation Updates (Day 1-2)

| Task | File | Action | Time |
|------|------|--------|------|
| 2.1 | SKILL.md | Update metadata to be accurate | 15 min |
| 2.2 | tool-reference.md | Mark all 53 tools as implemented | 10 min |
| 2.3 | troubleshooting.md | Fix docs/API.md path | 5 min |
| 2.4 | Test-plan.md | Update to reflect all tools implemented | 20 min |

**Sprint 2 Total:** 50 minutes

---

## 🔧 Detailed Fix Instructions

### Fix 1.1: Remove Chat URL

**File:** Test-plan.md:163

```diff
- # https://chat.qwen.ai/s/2503e7d5-e7b7-4b82-99f1-15a01453b0b1?fev=0.2.36
+ # End of document
```

### Fix 1.2: batch_process.py Return Code

**File:** skills/excel-tools/scripts/batch_process.py

```python
# BEFORE (lines 20-25):
def run_tool(tool: str, **kwargs) -> dict:
    cmd = [f"xls-{tool}"]
    for key, value in kwargs.items():
        cmd.extend([f"--{key.replace('_', '-')}", str(value)])
    result = subprocess.run(cmd, capture_output=True, text=True)
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError:
        return {"status": "error", "error": result.stderr or result.stdout}

# AFTER:
def run_tool(tool: str, **kwargs) -> dict:
    cmd = [f"xls-{tool}"]
    for key, value in kwargs.items():
        cmd.extend([f"--{key.replace('_', '-')}", str(value)])
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    # Check return code first
    if result.returncode != 0:
        try:
            error_data = json.loads(result.stdout)
            return error_data
        except json.JSONDecodeError:
            return {
                "status": "error",
                "exit_code": result.returncode,
                "error": result.stderr or result.stdout or f"Tool exited with code {result.returncode}"
            }
    
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError:
        return {"status": "error", "error": f"Invalid JSON output: {result.stdout[:200]}"}
```

### Fix 1.3: create_workbook.py Error Source

**File:** skills/excel-tools/scripts/create_workbook.py:37

```python
# BEFORE:
if result.returncode != 0:
    return {"status": "error", "error": result.stderr}  # WRONG

# AFTER:
if result.returncode != 0:
    # excel-agent-tools writes errors to stdout as JSON
    try:
        error_data = json.loads(result.stdout)
        return error_data
    except json.JSONDecodeError:
        return {"status": "error", "error": result.stdout or result.stderr}
```

### Fix 1.4: Add requests Dependency

**File:** pyproject.toml

```toml
[project]
dependencies = [
    "openpyxl>=3.1.5",
    "defusedxml>=0.7.1",
    "oletools>=0.60.2",
    "formulas[excel]>=1.3.4",
    "pandas>=3.0.0",
    "jsonschema>=4.26.0",
    "requests>=2.32.0",  # NEW: For oletools URL handling
]
```

### Fix 2.1: SKILL.md Metadata

**File:** skills/excel-tools/SKILL.md

```yaml
# BEFORE:
total-tools: "53"
coverage: ">90%"

# AFTER:
total-tools: "53"  # All tools now implemented (Phase 14 complete)
# Coverage measured via pytest --cov; check CI badge for current value
```

---

## ✅ Validation Checklist

### Pre-Merge

- [ ] Test-plan.md chat URL removed
- [ ] batch_process.py checks return codes
- [ ] create_workbook.py reads errors from stdout
- [ ] requests added to pyproject.toml
- [ ] SKILL.md metadata updated
- [ ] All 53 tools verified implemented
- [ ] Script tests pass
- [ ] Code review completed

---

**Plan Created:** April 10, 2026  
**Next Action:** Execute fixes
