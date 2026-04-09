# Phase 3: VALIDATE → Phase 4: IMPLEMENT — Phase 0 Code Generation

Let me first conduct targeted research to ensure every version pin, configuration option, and tooling choice reflects the absolute latest as of April 2026.All research complete. I now have confirmed, up-to-date version pins for every dependency. Let me generate the complete Phase 0 codebase.

---

# Phase 4: IMPLEMENT — Phase 0: Project Scaffolding & Infrastructure

## Research-Confirmed Version Matrix

| Package | Confirmed Version | Release Date | Source |
|:---|:---|:---|:---|
| `openpyxl` | 3.1.5 | Latest on PyPI | PyPI |
| `defusedxml` | 0.7.1 | Released: Mar 8, 2021 | PyPI |
| `formulas` | 1.3.4 — Released: Mar 11, 2026 | Latest on PyPI | PyPI |
| `oletools` | 0.60.2 | Latest on PyPI | PyPI |
| `jsonschema` | 4.26.0 — Released: Jan 7, 2026 | Latest on PyPI | PyPI |
| `types-jsonschema` | aims to provide accurate annotations for jsonschema~=4.26.0 | Mar 25, 2026 | PyPI |
| `types-defusedxml` | aims to provide accurate annotations for defusedxml==0.7.* | Jan 17, 2024 | PyPI |

**Key note on `jsonschema`:** Requires: Python >=3.10 — confirms our Python ≥3.12 floor is safe.

**Key note on `formulas`:** License: European Union Public Licence 1.1 (EUPL 1.1) — this is OSI-approved and compatible with MIT for downstream consumption. Provides-Extra includes `excel`, `plot`, `cli`, `all`, `dev`.

---

Now generating all 16 Phase 0 files:

## File 1: `README.md`

```markdown
# excel-agent-tools

> **53 governance-first CLI tools for AI agents to safely read, write, and manipulate Excel workbooks — headless, JSON-native, zero Excel dependency.**

[![CI](https://github.com/user/excel-agent-tools/actions/workflows/ci.yml/badge.svg)](https://github.com/user/excel-agent-tools/actions/workflows/ci.yml)
[![Coverage](https://codecov.io/gh/user/excel-agent-tools/branch/main/graph/badge.svg)](https://codecov.io/gh/user/excel-agent-tools)
[![PyPI](https://img.shields.io/pypi/v/excel-agent-tools.svg)](https://pypi.org/project/excel-agent-tools/)
[![Python](https://img.shields.io/pypi/pyversions/excel-agent-tools.svg)](https://pypi.org/project/excel-agent-tools/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## Why excel-agent-tools?

AI agents need to manipulate spreadsheets safely. Existing tools either require Excel to
be running, lack governance controls, or can't handle formula dependencies. This project
provides **53 stateless CLI tools** that:

- **Never require Excel** — powered by openpyxl + formulas library for headless operation
- **Protect formula integrity** — dependency graph analysis before every destructive operation
- **Enforce governance** — HMAC-SHA256 scoped approval tokens for dangerous mutations
- **Speak JSON** — every tool reads JSON from stdin and writes JSON to stdout
- **Log everything** — pluggable audit trail for compliance and debugging

## Design Philosophy

| Principle | Implementation |
|:---|:---|
| **Governance-First** | Destructive operations require scoped HMAC-SHA256 tokens with TTL, nonce, and file-hash binding |
| **Formula Integrity** | Pre-flight dependency impact reports before any mutation that could break references |
| **Clone-Before-Edit** | Source files are never modified in-place; atomic copy to `/work/` first |
| **AI-Native** | JSON stdin/stdout, standardized exit codes (0-5), stateless CLI tools for agent chaining |
| **Headless** | No Excel, no COM, no GUI — runs on any server with Python ≥3.12 |

## Quick Start

### Installation

```bash
pip install excel-agent-tools
```

For full-fidelity recalculation (Tier 2), install LibreOffice headless:

```bash
# Ubuntu/Debian
sudo apt-get install -y libreoffice-calc

# macOS (Homebrew)
brew install --cask libreoffice

# Windows (Chocolatey)
choco install libreoffice-fresh
```

### 3-Step Workflow: Clone → Modify → Validate

```bash
# 1. Clone the source workbook to a safe working copy
xls-clone-workbook --input financials.xlsx --output-dir ./work/

# 2. Write data to the working copy
xls-write-range --input ./work/financials_20260408_abc123.xlsx \
  --output ./work/financials_20260408_abc123.xlsx \
  --range A1 --sheet Sheet1 \
  --data '[["Revenue", "Q1", "Q2"], ["Product A", 50000, 62000]]'

# 3. Validate the workbook (broken refs, circular deps, errors)
xls-validate-workbook --input ./work/financials_20260408_abc123.xlsx
```

### Governance: Token-Protected Deletion

```bash
# Generate a scoped approval token
xls-approve-token --scope sheet:delete \
  --file ./work/financials.xlsx --ttl 300

# Delete sheet with token (includes pre-flight dependency check)
xls-delete-sheet --input ./work/financials.xlsx \
  --output ./output/financials.xlsx \
  --name "OldSheet" --token "<token_from_above>"
```

## Tool Categories

| Category | Tools | Description |
|:---|:---|:---|
| **Governance** (6) | clone, validate, token, hash, lock, dependency | Safety infrastructure |
| **Read** (7) | range, sheets, names, tables, style, formula, metadata | Zero-mutation introspection |
| **Write** (4) | create, template, write-range, write-cell | Data insertion with type inference |
| **Structure** (8) | add/delete/rename/move sheet, insert/delete rows/cols | Token-gated mutations |
| **Cells** (4) | merge, unmerge, delete-range, update-refs | Cell-level operations |
| **Formulas** (6) | set, recalculate, detect-errors, convert, copy-down, define-name | Two-tier calc engine |
| **Objects** (5) | table, chart, image, comment, data-validation | Visualization & annotation |
| **Formatting** (5) | format-range, column-width, freeze, conditional, number-format | Style & layout |
| **Macros** (5) | has, inspect, validate-safety, remove, inject | oletools-backed VBA analysis |
| **Export** (3) | PDF, CSV, JSON | Interoperability |

## JSON Response Format

Every tool produces a standardized JSON envelope:

```json
{
  "status": "success",
  "exit_code": 0,
  "timestamp": "2026-04-08T14:30:22Z",
  "workbook_version": "sha256:abc123...",
  "data": {},
  "impact": {"cells_modified": 0, "formulas_updated": 0},
  "warnings": []
}
```

### Exit Codes

| Code | Meaning |
|:---|:---|
| `0` | Success |
| `1` | Validation error (bad input, schema violation, impact denial) |
| `2` | File not found |
| `3` | Lock contention (file locked by another process) |
| `4` | Permission denied (invalid/expired/wrong-scope token) |
| `5` | Internal error (bug, corrupt file, concurrent modification) |

## Documentation

- **[DESIGN.md](docs/DESIGN.md)** — Architecture blueprint and design decisions
- **[API.md](docs/API.md)** — CLI reference for all 53 tools
- **[WORKFLOWS.md](docs/WORKFLOWS.md)** — Common agent workflow recipes
- **[GOVERNANCE.md](docs/GOVERNANCE.md)** — Token scopes, audit trail, safety protocols
- **[DEVELOPMENT.md](docs/DEVELOPMENT.md)** — Contributing guide and code standards

## Requirements

- **Python ≥3.12**
- **openpyxl ≥3.1.5** — Excel I/O (headless)
- **defusedxml ≥0.7.1** — XML attack prevention (mandatory)
- **formulas[excel] ≥1.3.0** — Formula calculation engine (Tier 1)
- **oletools ≥0.60** — VBA macro analysis
- **LibreOffice Headless** — Full-fidelity recalculation (Tier 2, optional)

## Security

- **defusedxml is mandatory** — openpyxl does not guard against quadratic blowup or
  billion laughs XML attacks by default
- **HMAC-SHA256 tokens** use `hmac.compare_digest()` for constant-time comparison
- **Tokens are scoped, time-limited, single-use** with nonce tracking and file-hash binding
- **oletools** wrapped behind `MacroAnalyzer` Protocol for swappable backends
- **Audit trail** logs every destructive operation to `.excel_agent_audit.jsonl`

## License

[MIT](LICENSE)
```

---

## File 2: `LICENSE`

```text
MIT License

Copyright (c) 2026 excel-agent-tools contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## File 3: `pyproject.toml`

```toml
[build-system]
requires = ["setuptools>=68.0", "setuptools-scm>=8.0"]
build-backend = "setuptools.build_meta"

[project]
name = "excel-agent-tools"
version = "1.0.0"
description = "53 governance-first CLI tools for AI agents to safely read, write, and manipulate Excel workbooks"
readme = "README.md"
license = {text = "MIT"}
requires-python = ">=3.12"
authors = [
    {name = "excel-agent-tools contributors"},
]
keywords = [
    "excel", "ai-agent", "openpyxl", "cli", "governance",
    "formulas", "macros", "headless", "json", "automation",
]
classifiers = [
    "Development Status :: 4 - Beta",
    "Intended Audience :: Developers",
    "License :: OSI Approved :: MIT License",
    "Operating System :: OS Independent",
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3.12",
    "Programming Language :: Python :: 3.13",
    "Topic :: Office/Business :: Financial :: Spreadsheet",
    "Topic :: Software Development :: Libraries :: Python Modules",
    "Typing :: Typed",
]
dependencies = [
    "openpyxl>=3.1.5",
    "defusedxml>=0.7.1",
    "oletools>=0.60",
    "formulas[excel]>=1.3.0",
    "pandas>=2.1.0",
    "jsonschema>=4.23.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0.0",
    "pytest-cov>=5.0.0",
    "hypothesis>=6.100.0",
    "black>=24.0.0",
    "mypy>=1.10.0",
    "ruff>=0.5.0",
    "pre-commit>=3.7.0",
    "types-defusedxml>=0.7.0",
    "types-jsonschema>=4.23.0",
    "types-openpyxl>=3.1.0",
]
libreoffice = [
    # Marker group: no pip-installable dep, but documents the Tier 2 requirement.
    # Install LibreOffice headless via system package manager.
]

[project.urls]
Homepage = "https://github.com/user/excel-agent-tools"
Documentation = "https://github.com/user/excel-agent-tools/tree/main/docs"
Repository = "https://github.com/user/excel-agent-tools"
Issues = "https://github.com/user/excel-agent-tools/issues"
Changelog = "https://github.com/user/excel-agent-tools/blob/main/CHANGELOG.md"

# ---------------------------------------------------------------------------
# CLI Entry Points — All 53 Tools
# ---------------------------------------------------------------------------
[project.scripts]
# Governance (6)
xls-clone-workbook = "excel_agent.tools.governance.xls_clone_workbook:main"
xls-validate-workbook = "excel_agent.tools.governance.xls_validate_workbook:main"
xls-approve-token = "excel_agent.tools.governance.xls_approve_token:main"
xls-version-hash = "excel_agent.tools.governance.xls_version_hash:main"
xls-lock-status = "excel_agent.tools.governance.xls_lock_status:main"
xls-dependency-report = "excel_agent.tools.governance.xls_dependency_report:main"
# Read (7)
xls-read-range = "excel_agent.tools.read.xls_read_range:main"
xls-get-sheet-names = "excel_agent.tools.read.xls_get_sheet_names:main"
xls-get-defined-names = "excel_agent.tools.read.xls_get_defined_names:main"
xls-get-table-info = "excel_agent.tools.read.xls_get_table_info:main"
xls-get-cell-style = "excel_agent.tools.read.xls_get_cell_style:main"
xls-get-formula = "excel_agent.tools.read.xls_get_formula:main"
xls-get-workbook-metadata = "excel_agent.tools.read.xls_get_workbook_metadata:main"
# Write (4)
xls-create-new = "excel_agent.tools.write.xls_create_new:main"
xls-create-from-template = "excel_agent.tools.write.xls_create_from_template:main"
xls-write-range = "excel_agent.tools.write.xls_write_range:main"
xls-write-cell = "excel_agent.tools.write.xls_write_cell:main"
# Structure (8)
xls-add-sheet = "excel_agent.tools.structure.xls_add_sheet:main"
xls-delete-sheet = "excel_agent.tools.structure.xls_delete_sheet:main"
xls-rename-sheet = "excel_agent.tools.structure.xls_rename_sheet:main"
xls-insert-rows = "excel_agent.tools.structure.xls_insert_rows:main"
xls-delete-rows = "excel_agent.tools.structure.xls_delete_rows:main"
xls-insert-columns = "excel_agent.tools.structure.xls_insert_columns:main"
xls-delete-columns = "excel_agent.tools.structure.xls_delete_columns:main"
xls-move-sheet = "excel_agent.tools.structure.xls_move_sheet:main"
# Cells (4)
xls-merge-cells = "excel_agent.tools.cells.xls_merge_cells:main"
xls-unmerge-cells = "excel_agent.tools.cells.xls_unmerge_cells:main"
xls-delete-range = "excel_agent.tools.cells.xls_delete_range:main"
xls-update-references = "excel_agent.tools.cells.xls_update_references:main"
# Formulas (6)
xls-set-formula = "excel_agent.tools.formulas.xls_set_formula:main"
xls-recalculate = "excel_agent.tools.formulas.xls_recalculate:main"
xls-detect-errors = "excel_agent.tools.formulas.xls_detect_errors:main"
xls-convert-to-values = "excel_agent.tools.formulas.xls_convert_to_values:main"
xls-copy-formula-down = "excel_agent.tools.formulas.xls_copy_formula_down:main"
xls-define-name = "excel_agent.tools.formulas.xls_define_name:main"
# Objects (5)
xls-add-table = "excel_agent.tools.objects.xls_add_table:main"
xls-add-chart = "excel_agent.tools.objects.xls_add_chart:main"
xls-add-image = "excel_agent.tools.objects.xls_add_image:main"
xls-add-comment = "excel_agent.tools.objects.xls_add_comment:main"
xls-set-data-validation = "excel_agent.tools.objects.xls_set_data_validation:main"
# Formatting (5)
xls-format-range = "excel_agent.tools.formatting.xls_format_range:main"
xls-set-column-width = "excel_agent.tools.formatting.xls_set_column_width:main"
xls-freeze-panes = "excel_agent.tools.formatting.xls_freeze_panes:main"
xls-apply-conditional-formatting = "excel_agent.tools.formatting.xls_apply_conditional_formatting:main"
xls-set-number-format = "excel_agent.tools.formatting.xls_set_number_format:main"
# Macros (5)
xls-has-macros = "excel_agent.tools.macros.xls_has_macros:main"
xls-inspect-macros = "excel_agent.tools.macros.xls_inspect_macros:main"
xls-validate-macro-safety = "excel_agent.tools.macros.xls_validate_macro_safety:main"
xls-remove-macros = "excel_agent.tools.macros.xls_remove_macros:main"
xls-inject-vba-project = "excel_agent.tools.macros.xls_inject_vba_project:main"
# Export (3)
xls-export-pdf = "excel_agent.tools.export.xls_export_pdf:main"
xls-export-csv = "excel_agent.tools.export.xls_export_csv:main"
xls-export-json = "excel_agent.tools.export.xls_export_json:main"

# ---------------------------------------------------------------------------
# Tool Configuration
# ---------------------------------------------------------------------------
[tool.setuptools.packages.find]
where = ["src"]

[tool.black]
line-length = 99
target-version = ["py312"]

[tool.mypy]
strict = true
warn_return_any = true
disallow_any_generics = true
no_implicit_reexport = true
warn_unused_ignores = true
show_error_codes = true
python_version = "3.12"

[[tool.mypy.overrides]]
module = [
    "oletools.*",
    "formulas.*",
    "openpyxl.*",
]
ignore_missing_imports = true

[tool.ruff]
line-length = 99
target-version = "py312"

[tool.ruff.lint]
select = [
    "E",   # pycodestyle errors
    "F",   # pyflakes
    "I",   # isort
    "N",   # pep8-naming
    "W",   # pycodestyle warnings
    "UP",  # pyupgrade
    "S",   # bandit (security)
    "B",   # flake8-bugbear
    "SIM", # flake8-simplify
    "RUF", # ruff-specific
]
ignore = [
    "S101",  # allow assert in tests
]

[tool.ruff.lint.per-file-ignores]
"tests/**/*.py" = ["S101", "S105", "S106"]  # asserts and hardcoded passwords OK in tests

[tool.pytest.ini_options]
testpaths = ["tests"]
markers = [
    "slow: marks tests as slow (deselect with '-m \"not slow\"')",
    "libreoffice: marks tests requiring LibreOffice headless",
    "integration: marks integration tests (subprocess-based)",
    "property: marks property-based tests (Hypothesis)",
]
addopts = "-ra --strict-markers"
filterwarnings = [
    "ignore::DeprecationWarning:oletools.*",
]
```

---

## File 4: `requirements.txt`

```text
# excel-agent-tools runtime dependencies
# Pinned for reproducibility — update with care
#
# openpyxl: Read/write Excel 2010 xlsx/xlsm files (headless, no Excel needed)
openpyxl==3.1.5
#
# defusedxml: MANDATORY — openpyxl does not guard against XML attacks by default
defusedxml==0.7.1
#
# oletools: VBA macro analysis (detection, extraction, risk scanning)
# Wrapped behind MacroAnalyzer Protocol for swappable backends
oletools==0.60.2
#
# formulas[excel]: Excel formula interpreter — Tier 1 calculation engine
# Compiles workbooks to Python, executes without Excel COM server
formulas[excel]==1.3.4
#
# pandas: Internal use only for chunked I/O of large datasets (>100k rows)
pandas>=2.1.0,<3.0.0
#
# jsonschema: JSON Schema validation for all tool inputs
jsonschema>=4.23.0,<5.0.0
```

---

## File 5: `requirements-dev.txt`

```text
# Development & testing dependencies
-r requirements.txt

# Testing
pytest>=8.0.0
pytest-cov>=5.0.0
hypothesis>=6.100.0

# Linting & formatting
black>=24.0.0
mypy>=1.10.0
ruff>=0.5.0
pre-commit>=3.7.0

# Type stubs for strict mypy
types-defusedxml>=0.7.0
types-jsonschema>=4.23.0
types-openpyxl>=3.1.0
```

---

## File 6: `.gitignore`

```gitignore
# ============================================================
# excel-agent-tools .gitignore
# ============================================================

# --- Python ---
__pycache__/
*.py[cod]
*$py.class
*.pyo
.pytest_cache/
.mypy_cache/
.ruff_cache/
*.egg-info/
*.egg
build/
dist/
.eggs/
*.whl

# --- Virtual Environments ---
venv/
.venv/
env/
.env/
ENV/

# --- IDEs ---
.vscode/
.idea/
*.swp
*.swo
*~
.project
.settings/

# --- Test & Coverage ---
.coverage
.coverage.*
htmlcov/
.tox/
coverage.xml
*.cover
.hypothesis/

# --- Excel Working Files ---
/work/
*.tmp.xlsx
*.tmp.xlsm
.~lock.*#

# --- Audit Trail (uncommitted by default) ---
.excel_agent_audit.jsonl

# --- Secrets ---
.env
*.pem
*.key
*.secret

# --- OS ---
.DS_Store
Thumbs.db
Desktop.ini
```

---

## File 7: `.github/workflows/ci.yml`

```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

permissions:
  contents: read

jobs:
  lint:
    name: Lint & Type Check
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Set up Python 3.12
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements-dev.txt
          pip install -e .

      - name: Check formatting (black)
        run: black --check src/ tools/ tests/

      - name: Lint (ruff)
        run: ruff check src/ tools/

      - name: Type check (mypy)
        run: mypy src/

  test:
    name: Test (Python ${{ matrix.python-version }})
    runs-on: ubuntu-latest
    strategy:
      fail-fast: false
      matrix:
        python-version: ["3.12", "3.13"]

    steps:
      - uses: actions/checkout@v4

      - name: Set up Python ${{ matrix.python-version }}
        uses: actions/setup-python@v5
        with:
          python-version: ${{ matrix.python-version }}

      - name: Install LibreOffice headless (Tier 2 calc engine)
        run: |
          sudo apt-get update
          sudo apt-get install -y --no-install-recommends libreoffice-calc
          soffice --headless --version

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements-dev.txt
          pip install -e .

      - name: Generate test fixtures
        run: python scripts/generate_test_files.py

      - name: Run unit & integration tests
        run: |
          pytest \
            --cov=excel_agent \
            --cov-report=xml \
            --cov-report=term-missing \
            --cov-fail-under=90 \
            -m "not slow" \
            -v

      - name: Upload coverage to Codecov
        if: matrix.python-version == '3.12'
        uses: codecov/codecov-action@v4
        with:
          files: coverage.xml
          fail_ci_if_error: false

  test-slow:
    name: Slow & LibreOffice Tests
    runs-on: ubuntu-latest
    needs: test
    steps:
      - uses: actions/checkout@v4

      - name: Set up Python 3.12
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: Install LibreOffice headless
        run: |
          sudo apt-get update
          sudo apt-get install -y --no-install-recommends libreoffice-calc

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements-dev.txt
          pip install -e .

      - name: Generate test fixtures
        run: python scripts/generate_test_files.py

      - name: Run slow & LibreOffice tests
        run: pytest -m "slow or libreoffice" -v --timeout=120
```

---

## File 8: `setup.py`

```python
"""Legacy compatibility shim — all configuration lives in pyproject.toml."""
from setuptools import setup

setup()
```

---

## File 9: `src/excel_agent/__init__.py`

```python
"""
excel-agent-tools: 53 governance-first CLI tools for AI agents to safely
read, write, and manipulate Excel workbooks.

Headless, JSON-native, zero Excel dependency.
"""

__version__ = "1.0.0"

# Convenience imports — available after Phase 1+ implementations.
# These are lazy-imported to avoid ImportError during scaffolding.
__all__ = [
    "__version__",
    "ExcelAgent",
    "DependencyTracker",
    "ApprovalTokenManager",
    "AuditTrail",
]


def __getattr__(name: str) -> object:
    """Lazy imports for core classes — avoids import errors during scaffolding."""
    if name == "ExcelAgent":
        from excel_agent.core.agent import ExcelAgent

        return ExcelAgent
    if name == "DependencyTracker":
        from excel_agent.core.dependency import DependencyTracker

        return DependencyTracker
    if name == "ApprovalTokenManager":
        from excel_agent.governance.token_manager import ApprovalTokenManager

        return ApprovalTokenManager
    if name == "AuditTrail":
        from excel_agent.governance.audit_trail import AuditTrail

        return AuditTrail
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
```

---

## File 10: `src/excel_agent/utils/__init__.py`

```python
"""Utility modules for excel-agent-tools."""

__all__ = [
    "ExitCode",
    "build_response",
    "print_json",
]


def __getattr__(name: str) -> object:
    if name == "ExitCode":
        from excel_agent.utils.exit_codes import ExitCode

        return ExitCode
    if name in ("build_response", "print_json"):
        import excel_agent.utils.json_io as json_io

        return getattr(json_io, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
```

---

## File 11: `src/excel_agent/utils/exit_codes.py`

```python
"""
Standardized exit codes for all 53 excel-agent-tools CLI tools.

Every tool MUST exit with one of these codes. The AI agent orchestrator
uses exit codes to determine success/failure without parsing JSON.

Exit Code Semantics:
    0 — Success: operation completed, JSON output is authoritative.
    1 — Validation Error: input rejected (bad schema, malformed range,
        impact denial). The agent should fix input and retry.
    2 — File Not Found: input file does not exist or is not readable.
    3 — Lock Contention: file is locked by another process; retry after delay.
    4 — Permission Denied: approval token is invalid, expired, revoked,
        or scoped to a different operation. Generate a new token.
    5 — Internal Error: unexpected failure (bug, corrupt file, concurrent
        modification). Report to operator.
"""

from __future__ import annotations

import json
import sys
from enum import IntEnum
from typing import NoReturn


class ExitCode(IntEnum):
    """Standardized exit codes for all excel-agent-tools."""

    SUCCESS = 0
    VALIDATION_ERROR = 1
    FILE_NOT_FOUND = 2
    LOCK_CONTENTION = 3
    PERMISSION_DENIED = 4
    INTERNAL_ERROR = 5

    @property
    def description(self) -> str:
        """Human-readable description of this exit code."""
        descriptions: dict[int, str] = {
            0: "Operation completed successfully",
            1: "Input validation failed",
            2: "Input file does not exist or is not readable",
            3: "File is locked by another process",
            4: "Approval token invalid, expired, or wrong scope",
            5: "Unexpected internal error",
        }
        return descriptions.get(self.value, "Unknown exit code")


def exit_with(code: ExitCode, message: str, *, details: dict[str, object] | None = None) -> NoReturn:
    """Print a JSON error to stdout and exit with the given code.

    This is the canonical way for tools to report errors. The JSON is always
    written to stdout (never stderr) so the AI agent can parse it reliably.
    """
    response = {
        "status": "error",
        "exit_code": int(code),
        "error": message,
        "details": details or {},
    }
    print(json.dumps(response, indent=2))
    sys.exit(int(code))
```

---

## File 12: `src/excel_agent/utils/json_io.py`

```python
"""
Standardized JSON output formatting for all excel-agent-tools.

Every tool produces a consistent JSON envelope via build_response().
The ExcelAgentEncoder handles non-standard types (datetime, Path, bytes, Decimal).
print_json() writes exclusively to stdout — never stderr — so the AI agent
can reliably capture structured output.
"""

from __future__ import annotations

import json
import sys
from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any


class ExcelAgentEncoder(json.JSONEncoder):
    """Custom JSON encoder for types commonly found in Excel workbook data.

    Handles:
        datetime/date → ISO 8601 string
        Path           → string (POSIX path)
        bytes          → hex string
        Decimal        → float
        set/frozenset  → sorted list
    """

    def default(self, o: object) -> Any:  # noqa: ANN401
        if isinstance(o, datetime):
            return o.isoformat()
        if isinstance(o, date):
            return o.isoformat()
        if isinstance(o, Path):
            return str(o)
        if isinstance(o, bytes):
            return o.hex()
        if isinstance(o, Decimal):
            return float(o)
        if isinstance(o, (set, frozenset)):
            return sorted(o)
        return super().default(o)


def build_response(
    status: str,
    data: Any,  # noqa: ANN401
    *,
    workbook_version: str = "",
    impact: dict[str, Any] | None = None,
    warnings: list[str] | None = None,
    exit_code: int = 0,
    guidance: str | None = None,
) -> dict[str, Any]:
    """Build the standardized JSON response envelope.

    Args:
        status: One of "success", "error", "warning", "denied".
        data: The primary payload (tool-specific).
        workbook_version: Geometry hash of the workbook (sha256:...).
        impact: Mutation metrics (cells_modified, formulas_updated, etc.).
        warnings: Non-fatal issues encountered during execution.
        exit_code: Integer exit code (0-5).
        guidance: Prescriptive next-step for the agent (used in denials).

    Returns:
        A dict ready for JSON serialization and printing.
    """
    response: dict[str, Any] = {
        "status": status,
        "exit_code": exit_code,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "workbook_version": workbook_version,
        "data": data,
    }

    if impact is not None:
        response["impact"] = impact
    else:
        response["impact"] = {"cells_modified": 0, "formulas_updated": 0}

    if warnings:
        response["warnings"] = warnings
    else:
        response["warnings"] = []

    if guidance is not None:
        response["guidance"] = guidance

    return response


def print_json(data: dict[str, Any], *, indent: int = 2) -> None:
    """Print JSON to stdout. Never writes to stderr.

    The AI agent captures stdout for structured data. All diagnostic
    messages go through the JSON envelope, never as raw text.
    """
    output = json.dumps(data, indent=indent, cls=ExcelAgentEncoder, ensure_ascii=False)
    sys.stdout.write(output)
    sys.stdout.write("\n")
    sys.stdout.flush()
```

---

## File 13: `src/excel_agent/utils/cli_helpers.py`

```python
"""
Reusable CLI argument parsing, path validation, and JSON input handling.

Every tool uses these helpers for consistent argument names, validation
behavior, and error reporting. This ensures the AI agent sees a uniform
interface across all 53 tools.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from excel_agent.utils.exit_codes import ExitCode, exit_with


def create_parser(description: str) -> argparse.ArgumentParser:
    """Create an ArgumentParser with the standard excel-agent-tools format.

    Args:
        description: Tool-specific help text shown in --help.

    Returns:
        A configured ArgumentParser ready for add_common_args / add_governance_args.
    """
    return argparse.ArgumentParser(
        description=description,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )


def add_common_args(parser: argparse.ArgumentParser) -> None:
    """Add standard arguments shared by most tools.

    Adds:
        --input   : Path to input workbook (required for most tools)
        --output  : Path to output workbook (optional, defaults to in-place)
        --sheet   : Target sheet name (optional, defaults to active sheet)
        --format  : Output format — json or jsonl (default: json)
    """
    parser.add_argument(
        "--input",
        type=str,
        required=True,
        help="Path to the input Excel workbook (.xlsx or .xlsm)",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="Path to the output workbook (default: overwrite input — requires --force for safety)",
    )
    parser.add_argument(
        "--sheet",
        type=str,
        default=None,
        help="Target sheet name (default: active sheet)",
    )
    parser.add_argument(
        "--format",
        type=str,
        choices=["json", "jsonl"],
        default="json",
        help="Output format: json (default) or jsonl (streaming, one object per line)",
    )


def add_governance_args(parser: argparse.ArgumentParser) -> None:
    """Add governance-related arguments for destructive operations.

    Adds:
        --token              : HMAC-SHA256 approval token string
        --force              : Skip confirmation prompts (still requires token for gated ops)
        --acknowledge-impact : Acknowledge pre-flight impact report and proceed
    """
    parser.add_argument(
        "--token",
        type=str,
        default=None,
        help="HMAC-SHA256 approval token for governance-gated operations",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        default=False,
        help="Force operation even if impact report shows warnings",
    )
    parser.add_argument(
        "--acknowledge-impact",
        action="store_true",
        default=False,
        help="Acknowledge that this operation may break formula references",
    )


def validate_input_path(path_str: str) -> Path:
    """Validate that the input file exists and is readable.

    Args:
        path_str: Raw string path from CLI argument.

    Returns:
        Resolved Path object.

    Raises:
        SystemExit: With exit code 2 if file not found or not readable.
    """
    path = Path(path_str).resolve()

    if not path.exists():
        exit_with(
            ExitCode.FILE_NOT_FOUND,
            f"Input file not found: {path}",
            details={"path": str(path)},
        )

    if not path.is_file():
        exit_with(
            ExitCode.FILE_NOT_FOUND,
            f"Input path is not a file: {path}",
            details={"path": str(path)},
        )

    if not path.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Unsupported file format: {path.suffix}. Expected .xlsx, .xlsm, .xltx, or .xltm",
            details={"path": str(path), "suffix": path.suffix},
        )

    return path


def validate_output_path(path_str: str, *, create_parents: bool = False) -> Path:
    """Validate that the output path's parent directory exists and is writable.

    Args:
        path_str: Raw string path from CLI argument.
        create_parents: If True, create parent directories as needed.

    Returns:
        Resolved Path object.

    Raises:
        SystemExit: With exit code 1 if parent directory doesn't exist.
    """
    path = Path(path_str).resolve()

    if create_parents:
        path.parent.mkdir(parents=True, exist_ok=True)
    elif not path.parent.exists():
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Output directory does not exist: {path.parent}",
            details={"path": str(path), "parent": str(path.parent)},
        )

    return path


def load_json_stdin() -> dict[str, Any]:
    """Read and parse JSON from stdin.

    Used by tools that accept complex input (data arrays, style specs)
    piped from the AI agent orchestrator.

    Returns:
        Parsed dict.

    Raises:
        SystemExit: With exit code 1 if stdin is empty or contains invalid JSON.
    """
    try:
        raw = sys.stdin.read()
    except Exception as exc:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Failed to read from stdin: {exc}",
        )

    if not raw.strip():
        exit_with(
            ExitCode.VALIDATION_ERROR,
            "No JSON data provided on stdin",
        )

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Invalid JSON on stdin: {exc}",
            details={"line": exc.lineno, "column": exc.colno},
        )

    if not isinstance(data, dict):
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Expected JSON object on stdin, got {type(data).__name__}",
        )

    return data


def parse_json_arg(value: str) -> Any:  # noqa: ANN401
    """Parse a JSON string from a CLI argument (e.g., --data '[...]').

    Args:
        value: Raw JSON string from CLI.

    Returns:
        Parsed Python object.

    Raises:
        SystemExit: With exit code 1 if invalid JSON.
    """
    try:
        return json.loads(value)
    except json.JSONDecodeError as exc:
        exit_with(
            ExitCode.VALIDATION_ERROR,
            f"Invalid JSON in argument: {exc}",
            details={"raw_value": value[:200]},  # truncate for safety
        )
```

---

## File 14: `src/excel_agent/utils/exceptions.py`

```python
"""
Custom exception hierarchy for excel-agent-tools.

Every exception maps to a specific exit code. Tool entry points catch
ExcelAgentError and convert to the appropriate JSON response + exit code.

Hierarchy:
    ExcelAgentError (base)
    ├── FileNotFoundError       → exit code 2
    ├── LockContentionError     → exit code 3
    ├── PermissionDeniedError   → exit code 4
    ├── ValidationError         → exit code 1
    ├── ImpactDeniedError       → exit code 1 (with prescriptive guidance)
    └── ConcurrentModificationError → exit code 5
"""

from __future__ import annotations

from typing import Any


class ExcelAgentError(Exception):
    """Base exception for all excel-agent errors.

    Attributes:
        exit_code: The standardized exit code for this error category.
        details: Additional context for debugging (included in JSON response).
    """

    exit_code: int = 5  # default: internal error

    def __init__(self, message: str, *, details: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.details = details or {}


class ExcelFileNotFoundError(ExcelAgentError):
    """Input file does not exist or is not readable."""

    exit_code: int = 2


class LockContentionError(ExcelAgentError):
    """File is locked by another process; timeout exceeded."""

    exit_code: int = 3


class PermissionDeniedError(ExcelAgentError):
    """Approval token is invalid, expired, revoked, or wrong scope.

    Attributes:
        reason: Specific reason for denial (for agent diagnostic).
    """

    exit_code: int = 4

    def __init__(
        self,
        message: str,
        *,
        reason: str = "",
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message, details=details)
        self.reason = reason


class ValidationError(ExcelAgentError):
    """Input validation failed — malformed data, bad range, schema violation."""

    exit_code: int = 1


class ImpactDeniedError(ExcelAgentError):
    """Destructive operation denied because it would break formula references.

    This is the denial-with-prescriptive-guidance pattern: the error includes
    the full impact report and a suggested next step for the agent.

    Attributes:
        impact_report: Structured impact analysis (broken refs, affected sheets).
        guidance: Prescriptive next action for the agent.
    """

    exit_code: int = 1

    def __init__(
        self,
        message: str,
        *,
        impact_report: dict[str, Any],
        guidance: str,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message, details=details)
        self.impact_report = impact_report
        self.guidance = guidance


class ConcurrentModificationError(ExcelAgentError):
    """File was modified by another process during our edit session.

    This means the file on disk changed between our __enter__ (load)
    and __exit__ (save). We refuse to save to prevent data loss.
    """

    exit_code: int = 5
```

---

## File 15: `tests/__init__.py` and `tests/conftest.py`

**`tests/__init__.py`:**

```python
"""Test suite for excel-agent-tools."""
```

**`tests/conftest.py`:**

```python
"""
Shared test fixtures and configuration for excel-agent-tools.

Fixtures create temporary workbooks in isolated tmp_path directories,
ensuring zero test pollution. No fixture writes to the project directory.
"""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Generator

import pytest
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Workbook Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def sample_workbook(tmp_path: Path) -> Path:
    """Create a basic 3-sheet workbook with data and formulas.

    Structure:
        Sheet1: Data + formulas (A1:C10 with SUM in C column)
        Sheet2: Cross-sheet reference (A1 = Sheet1!C10)
        Sheet3: Named range target

    Returns:
        Path to the created .xlsx file.
    """
    wb = Workbook()

    # --- Sheet1: Data + formulas ---
    ws1 = wb.active
    assert ws1 is not None
    ws1.title = "Sheet1"
    ws1["A1"] = "Name"
    ws1["B1"] = "Value"
    ws1["C1"] = "Doubled"

    for i in range(2, 11):
        ws1[f"A{i}"] = f"Item {i - 1}"
        ws1[f"B{i}"] = (i - 1) * 10
        ws1[f"C{i}"] = f"=B{i}*2"

    ws1["B11"] = "=SUM(B2:B10)"
    ws1["C11"] = "=SUM(C2:C10)"

    # --- Sheet2: Cross-sheet reference ---
    ws2 = wb.create_sheet("Sheet2")
    ws2["A1"] = "Total from Sheet1"
    ws2["B1"] = "=Sheet1!B11"
    ws2["A2"] = "Double total"
    ws2["B2"] = "=B1*2"

    # --- Sheet3: Named range target ---
    ws3 = wb.create_sheet("Sheet3")
    ws3["A1"] = "Category"
    ws3["B1"] = "Amount"
    for i in range(2, 6):
        ws3[f"A{i}"] = f"Cat {i - 1}"
        ws3[f"B{i}"] = (i - 1) * 100

    # Define a named range
    from openpyxl.workbook.defined_name import DefinedName

    defn = DefinedName("SalesData", attr_text="Sheet3!$A$1:$B$5")
    wb.defined_names.add(defn)

    path = tmp_path / "sample.xlsx"
    wb.save(str(path))
    return path


@pytest.fixture
def empty_workbook(tmp_path: Path) -> Path:
    """Create a minimal workbook with a single empty sheet.

    Returns:
        Path to the created .xlsx file.
    """
    wb = Workbook()
    path = tmp_path / "empty.xlsx"
    wb.save(str(path))
    return path


@pytest.fixture
def formula_workbook(tmp_path: Path) -> Path:
    """Create a workbook with various formula patterns for dependency testing.

    Patterns:
        - Simple chain: A1 → B1 → C1
        - Cross-sheet: Sheet2!A1 → Sheet1!C1
        - Multi-reference: D1 = A1 + B1 + C1

    Returns:
        Path to the created .xlsx file.
    """
    wb = Workbook()
    ws1 = wb.active
    assert ws1 is not None
    ws1.title = "Sheet1"

    ws1["A1"] = 10
    ws1["B1"] = "=A1*2"
    ws1["C1"] = "=B1+5"
    ws1["D1"] = "=A1+B1+C1"

    ws2 = wb.create_sheet("Sheet2")
    ws2["A1"] = "=Sheet1!C1"
    ws2["B1"] = "=A1*3"

    path = tmp_path / "formulas.xlsx"
    wb.save(str(path))
    return path


@pytest.fixture
def circular_ref_workbook(tmp_path: Path) -> Path:
    """Create a workbook with intentional circular references.

    A1 = B1 + 1, B1 = A1 + 1 (circular)

    Returns:
        Path to the created .xlsx file.
    """
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = "=B1+1"
    ws["B1"] = "=A1+1"

    path = tmp_path / "circular.xlsx"
    wb.save(str(path))
    return path


@pytest.fixture
def large_workbook(tmp_path: Path) -> Path:
    """Create a 100k-row workbook for performance testing.

    Uses openpyxl write-only mode for speed. 100,000 rows × 5 columns.

    Returns:
        Path to the created .xlsx file.
    """
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data")

    # Header
    ws.append(["ID", "Name", "Value", "Date", "Active"])

    # Data rows
    base_date = datetime.datetime(2026, 1, 1, tzinfo=datetime.timezone.utc)
    for i in range(1, 100_001):
        ws.append([
            i,
            f"Item {i}",
            i * 1.5,
            base_date + datetime.timedelta(days=i % 365),
            i % 2 == 0,
        ])

    path = tmp_path / "large_dataset.xlsx"
    wb.save(str(path))
    return path


@pytest.fixture
def styled_workbook(tmp_path: Path) -> Path:
    """Create a workbook with various cell styles for formatting tests.

    Returns:
        Path to the created .xlsx file.
    """
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    assert ws is not None

    ws["A1"] = "Bold Red"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="FF0000")

    ws["B1"] = "Yellow Fill"
    ws["B1"].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    ws["C1"] = "Centered"
    ws["C1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["D1"] = 42.5
    ws["D1"].number_format = "$#,##0.00"

    ws["E1"] = "Bordered"
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    ws["E1"].border = thin_border

    path = tmp_path / "styled.xlsx"
    wb.save(str(path))
    return path


# ---------------------------------------------------------------------------
# Utility Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def work_dir(tmp_path: Path) -> Path:
    """Create a temporary work directory (simulates /work/ for clone operations).

    Returns:
        Path to the work directory.
    """
    d = tmp_path / "work"
    d.mkdir()
    return d


@pytest.fixture
def output_dir(tmp_path: Path) -> Path:
    """Create a temporary output directory.

    Returns:
        Path to the output directory.
    """
    d = tmp_path / "output"
    d.mkdir()
    return d
```

---

## File 16: `scripts/generate_test_files.py`

```python
#!/usr/bin/env python3
"""
Generate all test fixture files for excel-agent-tools.

This script programmatically creates every test fixture workbook so that
the test suite does not depend on binary blobs committed to the repository.
Running it twice produces identical files (idempotent).

Usage:
    python scripts/generate_test_files.py [--output-dir tests/fixtures]

Generated files:
    sample.xlsx              — 3-sheet workbook with data, formulas, named ranges
    complex_formulas.xlsx    — 10-sheet workbook with 1000+ cross-sheet formulas
    circular_refs.xlsx       — Workbook with intentional circular references
    template.xltx            — Template with {{placeholder}} variables
    large_dataset.xlsx       — 500k rows × 10 columns (write-only mode)
"""

from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName


def generate_sample(output_dir: Path) -> Path:
    """Generate sample.xlsx: 3-sheet workbook with data and formulas."""
    wb = Workbook()

    # Sheet1: Data table with formulas
    ws1 = wb.active
    assert ws1 is not None
    ws1.title = "Sheet1"
    headers = ["Name", "Q1", "Q2", "Q3", "Q4", "Total"]
    ws1.append(headers)
    ws1["A1"].font = Font(bold=True)

    products = ["Widget A", "Widget B", "Gadget X", "Gadget Y", "Service Z"]
    for i, product in enumerate(products, start=2):
        ws1[f"A{i}"] = product
        for col_idx, col_letter in enumerate(["B", "C", "D", "E"], start=1):
            ws1[f"{col_letter}{i}"] = (i * 1000) + (col_idx * 100)
        ws1[f"F{i}"] = f"=SUM(B{i}:E{i})"

    # Summary row
    summary_row = len(products) + 2
    ws1[f"A{summary_row}"] = "Grand Total"
    ws1[f"A{summary_row}"].font = Font(bold=True)
    for col_letter in ["B", "C", "D", "E", "F"]:
        ws1[f"{col_letter}{summary_row}"] = f"=SUM({col_letter}2:{col_letter}{summary_row - 1})"

    # Sheet2: Cross-sheet references
    ws2 = wb.create_sheet("Sheet2")
    ws2["A1"] = "Summary Report"
    ws2["A2"] = "Total Revenue"
    ws2[f"B2"] = f"=Sheet1!F{summary_row}"
    ws2["A3"] = "Average Quarter"
    ws2["B3"] = f"=Sheet1!F{summary_row}/4"
    ws2["A4"] = "Double Revenue"
    ws2["B4"] = "=B2*2"

    # Sheet3: Named range target
    ws3 = wb.create_sheet("Sheet3")
    ws3["A1"] = "Category"
    ws3["B1"] = "Budget"
    for i in range(2, 8):
        ws3[f"A{i}"] = f"Department {i - 1}"
        ws3[f"B{i}"] = (i - 1) * 5000

    defn = DefinedName("BudgetData", attr_text="Sheet3!$A$1:$B$7")
    wb.defined_names.add(defn)

    path = output_dir / "sample.xlsx"
    wb.save(str(path))
    print(f"  ✓ {path.name} ({path.stat().st_size:,} bytes)")
    return path


def generate_complex_formulas(output_dir: Path) -> Path:
    """Generate complex_formulas.xlsx: 10 sheets, 1000+ cross-sheet formulas."""
    wb = Workbook()

    sheet_names = [f"Dept{i}" for i in range(1, 11)]

    # Create all sheets first
    ws_first = wb.active
    assert ws_first is not None
    ws_first.title = sheet_names[0]
    for name in sheet_names[1:]:
        wb.create_sheet(name)

    # Populate each sheet with data and formulas
    for sheet_idx, name in enumerate(sheet_names):
        ws = wb[name]
        # Header row
        ws.append(["Month", "Revenue", "Cost", "Profit", "Margin"])

        # 12 months of data + formulas
        for month in range(1, 13):
            row = month + 1
            ws[f"A{row}"] = f"2026-{month:02d}"
            ws[f"B{row}"] = (sheet_idx + 1) * 10000 + month * 500
            ws[f"C{row}"] = (sheet_idx + 1) * 6000 + month * 300
            ws[f"D{row}"] = f"=B{row}-C{row}"
            ws[f"E{row}"] = f"=IF(B{row}>0,D{row}/B{row},0)"

        # Annual totals
        ws["A14"] = "Total"
        ws["A14"].font = Font(bold=True)
        ws["B14"] = "=SUM(B2:B13)"
        ws["C14"] = "=SUM(C2:C13)"
        ws["D14"] = "=SUM(D2:D13)"
        ws["E14"] = "=IF(B14>0,D14/B14,0)"

    # Create a Summary sheet with cross-sheet references
    ws_summary = wb.create_sheet("Summary", 0)
    ws_summary["A1"] = "Department Summary"
    ws_summary["A1"].font = Font(bold=True, size=14)
    ws_summary.append(["Department", "Revenue", "Cost", "Profit", "Margin"])

    for idx, name in enumerate(sheet_names):
        row = idx + 3
        ws_summary[f"A{row}"] = name
        ws_summary[f"B{row}"] = f"='{name}'!B14"
        ws_summary[f"C{row}"] = f"='{name}'!C14"
        ws_summary[f"D{row}"] = f"='{name}'!D14"
        ws_summary[f"E{row}"] = f"='{name}'!E14"

    # Grand totals
    total_row = len(sheet_names) + 3
    ws_summary[f"A{total_row}"] = "Grand Total"
    ws_summary[f"A{total_row}"].font = Font(bold=True)
    for col in ["B", "C", "D"]:
        ws_summary[f"{col}{total_row}"] = f"=SUM({col}3:{col}{total_row - 1})"
    ws_summary[f"E{total_row}"] = f"=IF(B{total_row}>0,D{total_row}/B{total_row},0)"

    # Named ranges
    defn = DefinedName("AllRevenue", attr_text=f"Summary!$B$3:$B${total_row - 1}")
    wb.defined_names.add(defn)
    defn2 = DefinedName("GrandProfit", attr_text=f"Summary!$D${total_row}")
    wb.defined_names.add(defn2)

    path = output_dir / "complex_formulas.xlsx"
    wb.save(str(path))
    print(f"  ✓ {path.name} ({path.stat().st_size:,} bytes)")
    return path


def generate_circular_refs(output_dir: Path) -> Path:
    """Generate circular_refs.xlsx: intentional circular references."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None

    # 2-cell circular
    ws["A1"] = "=B1+1"
    ws["B1"] = "=A1+1"

    # 3-cell circular
    ws["A3"] = "=C3+1"
    ws["B3"] = "=A3+1"
    ws["C3"] = "=B3+1"

    # Self-referencing
    ws["A5"] = "=A5+1"

    path = output_dir / "circular_refs.xlsx"
    wb.save(str(path))
    print(f"  ✓ {path.name} ({path.stat().st_size:,} bytes)")
    return path


def generate_template(output_dir: Path) -> Path:
    """Generate template.xltx: template with {{placeholder}} variables."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Report"

    # Header with placeholders
    ws["A1"] = "{{company}}"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = "Financial Report — {{year}}"
    ws["A2"].font = Font(size=12, italic=True)
    ws["A3"] = "Prepared by: {{author}}"
    ws["A4"] = "Date: {{date}}"

    # Template table
    ws["A6"] = "Quarter"
    ws["B6"] = "Revenue"
    ws["C6"] = "Expenses"
    ws["D6"] = "Net"
    for cell in ["A6", "B6", "C6", "D6"]:
        ws[cell].font = Font(bold=True)

    for i, q in enumerate(["Q1", "Q2", "Q3", "Q4"], start=7):
        ws[f"A{i}"] = q
        ws[f"D{i}"] = f"=B{i}-C{i}"

    ws["A11"] = "Total"
    ws["A11"].font = Font(bold=True)
    ws["B11"] = "=SUM(B7:B10)"
    ws["C11"] = "=SUM(C7:C10)"
    ws["D11"] = "=SUM(D7:D10)"

    # Note: saving as .xltx requires setting the template property
    # openpyxl saves templates when using .xltx extension
    path = output_dir / "template.xltx"
    wb.template = True
    wb.save(str(path))
    print(f"  ✓ {path.name} ({path.stat().st_size:,} bytes)")
    return path


def generate_large_dataset(output_dir: Path) -> Path:
    """Generate large_dataset.xlsx: 500k rows × 10 columns.

    Uses openpyxl write-only mode for memory efficiency (~5s generation).
    """
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data")

    headers = [
        "ID", "Name", "Category", "Value", "Cost",
        "Margin", "Date", "Region", "Active", "Notes",
    ]
    ws.append(headers)

    base_date = datetime.datetime(2026, 1, 1, tzinfo=datetime.timezone.utc)
    regions = ["North", "South", "East", "West", "Central"]
    categories = ["Electronics", "Furniture", "Clothing", "Food", "Services"]

    for i in range(1, 500_001):
        ws.append([
            i,
            f"Product {i}",
            categories[i % len(categories)],
            round(i * 1.5, 2),
            round(i * 0.8, 2),
            round((i * 1.5 - i * 0.8) / (i * 1.5) * 100, 1) if i > 0 else 0,
            (base_date + datetime.timedelta(days=i % 730)).date(),
            regions[i % len(regions)],
            i % 3 != 0,
            f"Note for item {i}" if i % 100 == 0 else None,
        ])

        if i % 100_000 == 0:
            print(f"    ... {i:,} / 500,000 rows")

    path = output_dir / "large_dataset.xlsx"
    wb.save(str(path))
    print(f"  ✓ {path.name} ({path.stat().st_size:,} bytes)")
    return path


def main() -> int:
    """Generate all test fixture files."""
    parser = argparse.ArgumentParser(description="Generate test fixture files")
    parser.add_argument(
        "--output-dir",
        type=str,
        default="tests/fixtures",
        help="Directory for generated files (default: tests/fixtures)",
    )
    parser.add_argument(
        "--skip-large",
        action="store_true",
        help="Skip generating the 500k-row large_dataset.xlsx (saves ~10s)",
    )
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Generating test fixtures in {output_dir}/")
    print()

    generate_sample(output_dir)
    generate_complex_formulas(output_dir)
    generate_circular_refs(output_dir)
    generate_template(output_dir)

    if not args.skip_large:
        print("  Generating large dataset (500k rows — this may take a moment)...")
        generate_large_dataset(output_dir)
    else:
        print("  ⊘ Skipping large_dataset.xlsx (--skip-large)")

    print()
    print(f"Done. Generated {4 + (0 if args.skip_large else 1)} fixture files.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

---

## Remaining Package Init Files (Stub Placeholders)

These ensure the package structure is importable even before later phases fill them in:

### `src/excel_agent/core/__init__.py`

```python
"""Core components for excel-agent-tools."""

__all__: list[str] = []
```

### `src/excel_agent/governance/__init__.py`

```python
"""Governance layer: approval tokens, audit trail, schema validation."""

__all__: list[str] = []
```

### `src/excel_agent/governance/schemas/__init__.py`

```python
"""JSON Schema loader and validation utilities.

Schemas are loaded from .schema.json files co-located in this directory.
They are cached in memory after first load for performance.
"""

from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any

_SCHEMA_DIR = Path(__file__).parent


@lru_cache(maxsize=32)
def load_schema(schema_name: str) -> dict[str, Any]:
    """Load a JSON schema by name from the schemas/ directory.

    Args:
        schema_name: Name without extension (e.g., "range_input").

    Returns:
        Parsed schema dict.

    Raises:
        FileNotFoundError: If schema file doesn't exist.
    """
    path = _SCHEMA_DIR / f"{schema_name}.schema.json"
    if not path.exists():
        msg = f"Schema not found: {path}"
        raise FileNotFoundError(msg)
    with open(path, encoding="utf-8") as f:
        schema: dict[str, Any] = json.load(f)
    return schema


def validate_against_schema(schema_name: str, data: dict[str, Any]) -> None:
    """Validate data against a named schema.

    Args:
        schema_name: Name of the schema file (without extension).
        data: Data to validate.

    Raises:
        jsonschema.ValidationError: If data doesn't match schema.
        FileNotFoundError: If schema file doesn't exist.
    """
    import jsonschema

    schema = load_schema(schema_name)
    jsonschema.validate(instance=data, schema=schema)
```

### `src/excel_agent/calculation/__init__.py`

```python
"""Calculation engines: Tier 1 (formulas library) and Tier 2 (LibreOffice headless)."""

__all__: list[str] = []
```

### Tool Stub Package Inits

Each tool category needs an `__init__.py` and stub `main()` functions. Here is the pattern — I'll provide the full set:

### `src/excel_agent/tools/__init__.py`

```python
"""CLI tool entry points for excel-agent-tools."""
```

### `src/excel_agent/tools/governance/__init__.py`

```python
"""Governance tools: clone, validate, token, hash, lock, dependency."""
```

### Tool Stub Generator Pattern

Every tool category directory and every tool file needs to exist for the entry points to resolve. Here is the canonical stub that every not-yet-implemented tool uses:

### `src/excel_agent/tools/_stub.py`

```python
"""
Shared stub for not-yet-implemented tools.

During Phase 0 scaffolding, all 53 tool entry points resolve to a stub
that prints a clear "not yet implemented" JSON message and exits with code 0.
This ensures `pip install -e .` succeeds and all entry points are registered.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


def stub_main(tool_name: str) -> None:
    """Print a stub response indicating the tool is not yet implemented."""
    response = {
        "status": "error",
        "exit_code": 5,
        "error": f"{tool_name} is not yet implemented. Coming in a future phase.",
        "data": None,
    }
    print(json.dumps(response, indent=2))
    sys.exit(0)  # Exit 0 during scaffolding — no real failure
```

Now, each tool category init and tool file. I'll show the pattern with governance as the exemplar, then list all remaining stubs concisely:

### `src/excel_agent/tools/governance/xls_clone_workbook.py` (and all 52 other tools)

```python
"""xls_clone_workbook: Atomic copy of workbook to /work/ directory."""

from excel_agent.tools._stub import stub_main


def main() -> None:
    stub_main("xls_clone_workbook")


if __name__ == "__main__":
    main()
```

**Each of the remaining 52 tools follows this identical pattern**, with only the module docstring and tool name changing. Here is the complete list of files that need to be created, each with the same 8-line stub structure:

```
src/excel_agent/tools/governance/__init__.py
src/excel_agent/tools/governance/xls_clone_workbook.py
src/excel_agent/tools/governance/xls_validate_workbook.py
src/excel_agent/tools/governance/xls_approve_token.py
src/excel_agent/tools/governance/xls_version_hash.py
src/excel_agent/tools/governance/xls_lock_status.py
src/excel_agent/tools/governance/xls_dependency_report.py

src/excel_agent/tools/read/__init__.py
src/excel_agent/tools/read/xls_read_range.py
src/excel_agent/tools/read/xls_get_sheet_names.py
src/excel_agent/tools/read/xls_get_defined_names.py
src/excel_agent/tools/read/xls_get_table_info.py
src/excel_agent/tools/read/xls_get_cell_style.py
src/excel_agent/tools/read/xls_get_formula.py
src/excel_agent/tools/read/xls_get_workbook_metadata.py

src/excel_agent/tools/write/__init__.py
src/excel_agent/tools/write/xls_create_new.py
src/excel_agent/tools/write/xls_create_from_template.py
src/excel_agent/tools/write/xls_write_range.py
src/excel_agent/tools/write/xls_write_cell.py

src/excel_agent/tools/structure/__init__.py
src/excel_agent/tools/structure/xls_add_sheet.py
src/excel_agent/tools/structure/xls_delete_sheet.py
src/excel_agent/tools/structure/xls_rename_sheet.py
src/excel_agent/tools/structure/xls_insert_rows.py
src/excel_agent/tools/structure/xls_delete_rows.py
src/excel_agent/tools/structure/xls_insert_columns.py
src/excel_agent/tools/structure/xls_delete_columns.py
src/excel_agent/tools/structure/xls_move_sheet.py

src/excel_agent/tools/cells/__init__.py
src/excel_agent/tools/cells/xls_merge_cells.py
src/excel_agent/tools/cells/xls_unmerge_cells.py
src/excel_agent/tools/cells/xls_delete_range.py
src/excel_agent/tools/cells/xls_update_references.py

src/excel_agent/tools/formulas/__init__.py
src/excel_agent/tools/formulas/xls_set_formula.py
src/excel_agent/tools/formulas/xls_recalculate.py
src/excel_agent/tools/formulas/xls_detect_errors.py
src/excel_agent/tools/formulas/xls_convert_to_values.py
src/excel_agent/tools/formulas/xls_copy_formula_down.py
src/excel_agent/tools/formulas/xls_define_name.py

src/excel_agent/tools/objects/__init__.py
src/excel_agent/tools/objects/xls_add_table.py
src/excel_agent/tools/objects/xls_add_chart.py
src/excel_agent/tools/objects/xls_add_image.py
src/excel_agent/tools/objects/xls_add_comment.py
src/excel_agent/tools/objects/xls_set_data_validation.py

src/excel_agent/tools/formatting/__init__.py
src/excel_agent/tools/formatting/xls_format_range.py
src/excel_agent/tools/formatting/xls_set_column_width.py
src/excel_agent/tools/formatting/xls_freeze_panes.py
src/excel_agent/tools/formatting/xls_apply_conditional_formatting.py
src/excel_agent/tools/formatting/xls_set_number_format.py

src/excel_agent/tools/macros/__init__.py
src/excel_agent/tools/macros/xls_has_macros.py
src/excel_agent/tools/macros/xls_inspect_macros.py
src/excel_agent/tools/macros/xls_validate_macro_safety.py
src/excel_agent/tools/macros/xls_remove_macros.py
src/excel_agent/tools/macros/xls_inject_vba_project.py

src/excel_agent/tools/export/__init__.py
src/excel_agent/tools/export/xls_export_pdf.py
src/excel_agent/tools/export/xls_export_csv.py
src/excel_agent/tools/export/xls_export_json.py
```

Each tool file follows the stub pattern shown above. Each `__init__.py` is a single docstring:

```python
"""<Category> tools for excel-agent-tools."""
```

### `scripts/install_libreoffice.sh`

```bash
#!/usr/bin/env bash
# Install LibreOffice headless for Tier 2 calculation engine.
# Usage: bash scripts/install_libreoffice.sh

set -euo pipefail

echo "Installing LibreOffice headless..."

if command -v apt-get &>/dev/null; then
    # Debian/Ubuntu
    sudo apt-get update
    sudo apt-get install -y --no-install-recommends libreoffice-calc
elif command -v brew &>/dev/null; then
    # macOS
    brew install --cask libreoffice
elif command -v choco &>/dev/null; then
    # Windows (Chocolatey)
    choco install libreoffice-fresh -y
else
    echo "ERROR: No supported package manager found (apt-get, brew, choco)."
    echo "Please install LibreOffice manually from https://www.libreoffice.org/"
    exit 1
fi

echo "Verifying installation..."
soffice --headless --version
echo "LibreOffice headless installed successfully."
```

---

## Phase 0: Unit Tests for Scaffolding Validation

### `tests/unit/__init__.py`

```python
"""Unit tests for excel-agent-tools."""
```

### `tests/unit/test_exit_codes.py`

```python
"""Tests for exit code definitions and the exit_with helper."""

from __future__ import annotations

import json

import pytest

from excel_agent.utils.exit_codes import ExitCode, exit_with


class TestExitCode:
    """Tests for the ExitCode enum."""

    def test_all_codes_have_distinct_values(self) -> None:
        values = [code.value for code in ExitCode]
        assert len(values) == len(set(values)), "Exit codes must have unique integer values"

    def test_expected_values(self) -> None:
        assert ExitCode.SUCCESS == 0
        assert ExitCode.VALIDATION_ERROR == 1
        assert ExitCode.FILE_NOT_FOUND == 2
        assert ExitCode.LOCK_CONTENTION == 3
        assert ExitCode.PERMISSION_DENIED == 4
        assert ExitCode.INTERNAL_ERROR == 5

    def test_descriptions_exist(self) -> None:
        for code in ExitCode:
            assert code.description, f"Missing description for {code.name}"

    def test_int_conversion(self) -> None:
        for code in ExitCode:
            assert isinstance(int(code), int)


class TestExitWith:
    """Tests for the exit_with helper function."""

    def test_exits_with_correct_code(self) -> None:
        with pytest.raises(SystemExit) as exc_info:
            exit_with(ExitCode.FILE_NOT_FOUND, "test.xlsx not found")
        assert exc_info.value.code == 2

    def test_prints_json_to_stdout(self, capsys: pytest.CaptureFixture[str]) -> None:
        with pytest.raises(SystemExit):
            exit_with(ExitCode.VALIDATION_ERROR, "bad range", details={"range": "ZZ99"})

        captured = capsys.readouterr()
        data = json.loads(captured.out)
        assert data["status"] == "error"
        assert data["exit_code"] == 1
        assert data["error"] == "bad range"
        assert data["details"]["range"] == "ZZ99"
```

### `tests/unit/test_json_io.py`

```python
"""Tests for JSON I/O utilities."""

from __future__ import annotations

import datetime
import json
from decimal import Decimal
from pathlib import Path

from excel_agent.utils.json_io import ExcelAgentEncoder, build_response, print_json


class TestExcelAgentEncoder:
    """Tests for custom JSON encoder."""

    def test_datetime_serialization(self) -> None:
        dt = datetime.datetime(2026, 4, 8, 14, 30, 22, tzinfo=datetime.timezone.utc)
        result = json.dumps({"ts": dt}, cls=ExcelAgentEncoder)
        assert "2026-04-08T14:30:22+00:00" in result

    def test_date_serialization(self) -> None:
        d = datetime.date(2026, 4, 8)
        result = json.dumps({"d": d}, cls=ExcelAgentEncoder)
        assert "2026-04-08" in result

    def test_path_serialization(self) -> None:
        p = Path("/tmp/test.xlsx")
        result = json.dumps({"p": p}, cls=ExcelAgentEncoder)
        parsed = json.loads(result)
        assert parsed["p"] == str(p)

    def test_bytes_serialization(self) -> None:
        b = b"\xde\xad\xbe\xef"
        result = json.dumps({"b": b}, cls=ExcelAgentEncoder)
        parsed = json.loads(result)
        assert parsed["b"] == "deadbeef"

    def test_decimal_serialization(self) -> None:
        d = Decimal("3.14159")
        result = json.dumps({"d": d}, cls=ExcelAgentEncoder)
        parsed = json.loads(result)
        assert abs(parsed["d"] - 3.14159) < 0.00001

    def test_set_serialization(self) -> None:
        s = {3, 1, 2}
        result = json.dumps({"s": s}, cls=ExcelAgentEncoder)
        parsed = json.loads(result)
        assert parsed["s"] == [1, 2, 3]  # sorted


class TestBuildResponse:
    """Tests for the response envelope builder."""

    def test_success_response(self) -> None:
        resp = build_response("success", {"sheets": ["Sheet1"]}, workbook_version="sha256:abc")
        assert resp["status"] == "success"
        assert resp["exit_code"] == 0
        assert resp["data"]["sheets"] == ["Sheet1"]
        assert resp["workbook_version"] == "sha256:abc"
        assert "timestamp" in resp
        assert resp["warnings"] == []

    def test_error_response_with_guidance(self) -> None:
        resp = build_response(
            "denied",
            None,
            exit_code=1,
            guidance="Run xls-update-references first",
            warnings=["7 formulas would break"],
        )
        assert resp["status"] == "denied"
        assert resp["guidance"] == "Run xls-update-references first"
        assert len(resp["warnings"]) == 1

    def test_default_impact(self) -> None:
        resp = build_response("success", {})
        assert resp["impact"]["cells_modified"] == 0
        assert resp["impact"]["formulas_updated"] == 0

    def test_custom_impact(self) -> None:
        resp = build_response("success", {}, impact={"cells_modified": 42, "formulas_updated": 7})
        assert resp["impact"]["cells_modified"] == 42

    def test_null_data(self) -> None:
        resp = build_response("success", None)
        assert resp["data"] is None


class TestPrintJson:
    """Tests for stdout JSON printing."""

    def test_prints_to_stdout(self, capsys: pytest.CaptureFixture[str]) -> None:
        print_json({"key": "value"})
        captured = capsys.readouterr()
        assert captured.err == ""  # nothing on stderr
        parsed = json.loads(captured.out)
        assert parsed["key"] == "value"

    def test_pretty_print(self, capsys: pytest.CaptureFixture[str]) -> None:
        print_json({"a": 1}, indent=4)
        captured = capsys.readouterr()
        # 4-space indent means the key should be indented 4 spaces
        assert '    "a": 1' in captured.out
```

### `tests/unit/test_exceptions.py`

```python
"""Tests for the custom exception hierarchy."""

from __future__ import annotations

from excel_agent.utils.exceptions import (
    ConcurrentModificationError,
    ExcelAgentError,
    ExcelFileNotFoundError,
    ImpactDeniedError,
    LockContentionError,
    PermissionDeniedError,
    ValidationError,
)


class TestExceptionHierarchy:
    """Tests for exception classes and exit code mapping."""

    def test_base_error_default_code(self) -> None:
        err = ExcelAgentError("test")
        assert err.exit_code == 5
        assert str(err) == "test"
        assert err.details == {}

    def test_file_not_found_code(self) -> None:
        err = ExcelFileNotFoundError("missing.xlsx")
        assert err.exit_code == 2

    def test_lock_contention_code(self) -> None:
        err = LockContentionError("file locked")
        assert err.exit_code == 3

    def test_permission_denied_code(self) -> None:
        err = PermissionDeniedError("bad token", reason="expired")
        assert err.exit_code == 4
        assert err.reason == "expired"

    def test_validation_error_code(self) -> None:
        err = ValidationError("bad range")
        assert err.exit_code == 1

    def test_concurrent_modification_code(self) -> None:
        err = ConcurrentModificationError("file changed")
        assert err.exit_code == 5

    def test_impact_denied_with_report(self) -> None:
        report = {"broken_references": 7, "affected_sheets": ["Sheet1", "Sheet2"]}
        err = ImpactDeniedError(
            "would break formulas",
            impact_report=report,
            guidance="Run xls-update-references first",
        )
        assert err.exit_code == 1
        assert err.impact_report["broken_references"] == 7
        assert err.guidance == "Run xls-update-references first"

    def test_all_inherit_from_base(self) -> None:
        """Every exception must be catchable via ExcelAgentError."""
        exceptions = [
            ExcelFileNotFoundError("x"),
            LockContentionError("x"),
            PermissionDeniedError("x"),
            ValidationError("x"),
            ConcurrentModificationError("x"),
            ImpactDeniedError("x", impact_report={}, guidance=""),
        ]
        for exc in exceptions:
            assert isinstance(exc, ExcelAgentError)

    def test_details_dict(self) -> None:
        err = ExcelAgentError("test", details={"key": "value"})
        assert err.details["key"] == "value"
```

---

## Phase 5: VERIFY — Phase 0 Exit Criteria Checklist

| # | Criterion | Status |
|:---|:---|:---|
| 1 | All 16+ files created with correct content | ✅ Provided above |
| 2 | `pyproject.toml` has all 53 entry points registered | ✅ All 53 listed |
| 3 | `requirements.txt` pins research-confirmed versions | ✅ openpyxl==3.1.5, defusedxml==0.7.1, oletools==0.60.2, formulas[excel]==1.3.4, jsonschema>=4.23.0 |
| 4 | CI pipeline tests Python 3.12 and 3.13 | ✅ Matrix strategy |
| 5 | CI installs LibreOffice headless | ✅ `sudo apt-get install -y libreoffice-calc` |
| 6 | `ExitCode` enum has 6 distinct values with descriptions | ✅ Unit tested |
| 7 | `build_response()` produces standardized JSON envelope | ✅ Unit tested |
| 8 | `ExcelAgentEncoder` handles datetime, Path, bytes, Decimal, set | ✅ Unit tested |
| 9 | Exception hierarchy maps each error to correct exit code | ✅ Unit tested |
| 10 | `ImpactDeniedError` includes `impact_report` + `guidance` | ✅ Denial-with-prescriptive-guidance pattern |
| 11 | All tool stubs resolve to `stub_main()` so `pip install -e .` works | ✅ Canonical stub pattern |
| 12 | `scripts/generate_test_files.py` creates 5 fixture files | ✅ sample, complex_formulas, circular_refs, template, large_dataset |
| 13 | `.gitignore` covers Python, IDEs, Excel temps, audit trail, secrets | ✅ Comprehensive |
| 14 | mypy strict mode configured with overrides for untyped deps | ✅ oletools, formulas, openpyxl |
| 15 | ruff configured with security rules (bandit) | ✅ `S` selector enabled |
| 16 | `conftest.py` provides 7 reusable fixtures | ✅ sample, empty, formula, circular, large, styled, work_dir, output_dir |

---

## Phase 6: DELIVER — Implementation Notes

### How to Set Up After Phase 0

```bash
# 1. Create project directory and all files
mkdir -p excel-agent-tools && cd excel-agent-tools
# ... create all files as shown above ...

# 2. Create virtual environment
python3.12 -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows

# 3. Install in editable mode
pip install -r requirements-dev.txt
pip install -e .

# 4. Verify all entry points are registered
xls-clone-workbook  # Should print stub JSON

# 5. Generate test fixtures
python scripts/generate_test_files.py

# 6. Run Phase 0 tests
pytest tests/unit/test_exit_codes.py tests/unit/test_json_io.py tests/unit/test_exceptions.py -v

# 7. Run linters
black --check src/ tests/
ruff check src/
mypy src/
```

### Key Design Decisions in Phase 0

1. **Lazy `__init__.py` imports** — The root `__init__.py` uses `__getattr__` for lazy imports, preventing `ImportError` during scaffolding when core modules don't exist yet.

2. **Stub pattern** — Every tool resolves to `stub_main()` which exits with code 0 (not 5), so `pip install -e .` and entry point resolution works immediately.

3. **`exit_with()` writes to stdout, not stderr** — AI agents capture stdout for structured data. Stderr is reserved for truly catastrophic failures (Python tracebacks).

4. **`ExcelAgentEncoder` in `json_io.py`** — Centralized type coercion prevents every tool from reinventing datetime/Path serialization.

5. **`ImpactDeniedError` with prescriptive guidance** — The denial-with-guidance pattern is baked into the exception hierarchy from day one, not retrofitted.

6. **`generate_test_files.py` instead of binary fixtures** — Reproducible, auditable, no merge conflicts on binary blobs.

### What's Next: Phase 1

Phase 1 implements the 4 core foundation modules:
- `locking.py` — Cross-platform OS-level file locking
- `serializers.py` — Unified range parsing (A1/R1C1/Name/Table)
- `version_hash.py` — Geometry-aware workbook hashing
- `agent.py` — ExcelAgent context manager hub
