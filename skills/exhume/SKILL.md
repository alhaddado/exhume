---
name: exhume
description: Use when working with Excel .xlsx files - inspecting structure, reading cell data, extracting embedded OLE objects (packaged text files, documents), or converting spreadsheets to JSON/CSV. Triggers on xlsx, spreadsheet, embedded objects, OLE, Excel extraction, workbook introspection.
---

# Exhume - Excel Introspection & Extraction

## Overview

Exhume is a CLI tool and Python library for inspecting Excel files and extracting embedded OLE objects. JSON output by default, built for AI agents.

## Installation

Exhume must be installed before use. If the `exhume` command is not available:

```bash
cd /Users/omar/Projects/cli/exhume
pip install .
```

Verify: `exhume --version`

## When to Use

- User provides or references an `.xlsx` file
- Need to inspect spreadsheet structure (sheets, dimensions, merged cells)
- Need to read cell values, formulas, or types from a spreadsheet
- Need to extract embedded files (OLE objects) from an Excel file
- Need to convert a spreadsheet to JSON or CSV
- Need to list what objects are embedded in an Excel file with their positions

## Quick Reference

| Task | Command |
|------|---------|
| Workbook structure | `exhume info file.xlsx --pretty` |
| Cell data (one sheet) | `exhume cells file.xlsx --sheet Sheet1 --pretty` |
| Cell data (all sheets) | `exhume cells file.xlsx --pretty` |
| List embedded objects | `exhume objects file.xlsx --pretty` |
| Extract objects to disk | `exhume extract file.xlsx --out ./output` |
| Extract organized by row | `exhume extract file.xlsx --out ./output --by-row` |
| Full JSON export | `exhume convert file.xlsx --format json --out ./output` |
| Full CSV export | `exhume convert file.xlsx --format csv --out ./output` |

### Global Flags

- `--output json|text` -- default `json` (machine-readable), `text` for human display
- `--pretty` -- indent JSON output
- `--quiet` -- suppress progress messages
- `--sheet <name|index>` -- target a specific sheet

## Key Commands

### `exhume objects` - List embedded objects with context

Returns every embedded OLE object with its filename, position (sheet/row/column), size, and neighboring cell values. This is the command to use when you need to understand what's embedded in a spreadsheet and where.

The `neighborCells` field gives context from the same row -- useful for understanding what each embedded object relates to.

### `exhume extract` - Extract embedded objects to disk

Writes all embedded objects to a directory. Use `--flat` (default) for all files in one directory, or `--by-row` for `row-<N>/` subdirectories.

Handles filename deduplication automatically (appends `_2`, `_3`, etc.).

### `exhume convert` - Full export

Produces a structured directory:
```
output/
  metadata.json          # Workbook-level info
  sheets/<Name>/
    data.json (or .csv)  # Cell data
    metadata.json        # Sheet dimensions, merged cells
  objects/
    manifest.json        # All embedded objects with metadata
    files/               # Extracted files
```

## Python Library Usage

```python
from exhume import Workbook

wb = Workbook("file.xlsx")
info = wb.info()                      # WorkbookInfo
cells = wb.cells("Sheet1")           # List[CellInfo]
objects = wb.list_objects()           # List[EmbeddedObject]
wb.extract_objects("./output")       # Extract to disk
```

## Common Patterns

**Find objects from a specific row:**
```bash
exhume objects file.xlsx --pretty | python3 -c "
import json, sys
data = json.load(sys.stdin)
for obj in data['objects']:
    if obj['row'] == 48:
        print(obj['filename'], obj['sizeBytes'])
"
```

**Pipe object list to downstream processing:**
```bash
exhume objects file.xlsx | jq '.objects[] | {filename, row, neighborCells}'
```
