# Excel Tools Workflow Patterns

Common patterns for working with excel-agent-tools.

## Pattern 1: Clone-Modify-Export Pipeline

**Use Case**: Make changes to an existing workbook and export results.

```bash
# 1. Clone source
CLONE=$(xls-clone-workbook --input financials.xlsx --output-dir ./work/ | jq -r '.data.clone_path')

# 2. Read current state
xls-get-workbook-metadata --input "$CLONE"

# 3. Modify
xls-write-range --input "$CLONE" --output "$CLONE" --range A1 \
  --data '[["Q1", "Q2"], [100, 200]]'

# 4. Recalculate
xls-recalculate --input "$CLONE" --output "$CLONE"

# 5. Validate
xls-validate-workbook --input "$CLONE"

# 6. Export
xls-export-pdf --input "$CLONE" --outfile ./output/report.pdf
```

**Expected JSON Output**:
```json
{
  "status": "success",
  "data": {"output": "./output/report.pdf", "pages": 3}
}
```

## Pattern 2: Template Population

**Use Case**: Fill placeholders in template.

```bash
# Create from template with substitution
xls-create-from-template --template invoice.xltx --output invoice_001.xlsx \
  --vars '{"company": "Acme", "amount": "$500"}'
```

## Pattern 3: Safe Structural Edit

**Use Case**: Delete sheet that has formula references.

```bash
# 1. Check dependencies
xls-dependency-report --input workbook.xlsx | jq '.data.graph'

# 2. Generate token
TOKEN=$(xls-approve-token --scope sheet:delete --file workbook.xlsx | jq -r '.data.token')

# 3. Attempt deletion (may be denied)
xls-delete-sheet --input workbook.xlsx --output workbook.xlsx \
  --name "OldData" --token "$TOKEN"

# If denied, response contains guidance:
# {
#   "status": "denied",
#   "guidance": "Run xls-update-references --updates '[{\"old\": \"...\", \"new\": \"...\"}]'"
# }

# 4. Fix references per guidance
xls-update-references --input workbook.xlsx --output workbook.xlsx \
  --updates '[{"old": "OldData!A1", "new": "NewData!A1"}]'

# 5. Retry deletion with acknowledgment
xls-delete-sheet --input workbook.xlsx --output workbook.xlsx \
  --name "OldData" --token "$TOKEN" --acknowledge-impact
```

## Pattern 4: Batch Processing

**Use Case**: Process multiple files.

```bash
for file in ./data/*.xlsx; do
  # Clone
  clone=$(xls-clone-workbook --input "$file" --output-dir ./work/ | jq -r '.data.clone_path')
  
  # Process
  xls-recalculate --input "$clone" --output "$clone"
  
  # Export
  xls-export-csv --input "$clone" --outfile "./output/$(basename "$file" .xlsx).csv"
done
```

## Pattern 5: Large Dataset Streaming

**Use Case**: Read >100k rows efficiently.

```bash
# Chunked mode returns JSONL (one JSON per line)
xls-read-range --input large.xlsx --range A1:E100000 --chunked > output.jsonl

# Parse each chunk
while IFS= read -r line; do
  chunk=$(echo "$line" | jq '.')
  # Process chunk
done < output.jsonl
```

## Pattern 6: Macro Safety Audit

**Use Case**: Scan file for unsafe macros before processing.

```bash
# Check if macros exist
xls-has-macros --input report.xlsm | jq '.data.has_macros'

# If true, inspect
xls-inspect-macros --input report.xlsm

# Validate safety
SAFETY=$(xls-validate-macro-safety --input report.xlsm | jq -r '.data.risk_level')

# If high/critical risk, remove macros before processing
if [ "$SAFETY" = "high" ] || [ "$SAFETY" = "critical" ]; then
  TOKEN1=$(xls-approve-token --scope macro:remove --file report.xlsm | jq -r '.data.token')
  TOKEN2=$(xls-approve-token --scope macro:remove --file report.xlsm | jq -r '.data.token')
  xls-remove-macros --input report.xlsm --output report_clean.xlsx \
    --token "$TOKEN1" --token "$TOKEN2"
fi
```

## Pattern 7: Formula Error Detection

**Use Case**: Find and fix broken references.

```bash
# Detect errors
ERRORS=$(xls-detect-errors --input workbook.xlsx | jq '.data.errors')

# If errors found, fix by updating references
if [ $(echo "$ERRORS" | jq 'length') -gt 0 ]; then
  xls-update-references --input workbook.xlsx --output workbook.xlsx \
    --updates '[{"old": "Sheet1!#REF!", "new": "Sheet1!A1"}]'
fi
```

## Pattern 8: Conditional Formatting

**Use Case**: Add visual indicators.

```bash
# Color scale (green-yellow-red)
xls-apply-conditional-formatting --input report.xlsx --range A1:A100 \
  --type colorscale --colors '["00FF00", "FFFF00", "FF0000"]'

# Data bars
xls-apply-conditional-formatting --input report.xlsx --range B1:B100 \
  --type databar --color "638EC6"
```

## Error Handling Pattern

**Always check exit codes before parsing JSON**:

```python
import subprocess
import json

result = subprocess.run(
    ["xls-read-range", "--input", "data.xlsx", "--range", "A1"],
    capture_output=True,
    text=True
)

if result.returncode == 0:
    data = json.loads(result.stdout)
    values = data["data"]["values"]
elif result.returncode == 1:
    print("Validation error - check input")
elif result.returncode == 2:
    print("File not found")
elif result.returncode == 3:
    print("File locked - retry with backoff")
elif result.returncode == 4:
    print("Permission denied - generate new token")
elif result.returncode == 5:
    print("Internal error - check traceback")
```

## Python Integration Pattern

```python
import subprocess
import json
from pathlib import Path

def run_tool(tool: str, **kwargs) -> dict:
    """Run an excel-agent tool and return parsed JSON."""
    cmd = [f"xls-{tool}"]
    for key, value in kwargs.items():
        cmd.extend([f"--{key.replace('_', '-')}", str(value)])
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    data = json.loads(result.stdout)
    
    if result.returncode != 0:
        raise RuntimeError(f"Tool failed: {data.get('error', 'unknown')}")
    
    return data

# Usage
clone = run_tool("clone-workbook", input="data.xlsx", output_dir="./work/")
clone_path = clone["data"]["clone_path"]

meta = run_tool("get-workbook-metadata", input=clone_path)
print(f"Sheets: {meta['data']['sheet_count']}")
```
