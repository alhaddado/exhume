# Exhume

Excel introspection and embedded object extraction CLI, built for AI agents.

Exhume reads `.xlsx` files and lets you inspect their structure, extract embedded OLE objects (like text files packaged inside spreadsheets), and convert everything to machine-readable formats.

## Install

```bash
git clone <repo-url>
cd exhume
pip install .

# Or for development:
pip install -e ".[dev]"
```

## Claude Code Skill

Exhume ships with a Claude Code skill so Claude can automatically use it when working with Excel files. Claude Code discovers skills from `SKILL.md` files placed in specific directories.

**For all your projects** (personal skill):

```bash
mkdir -p ~/.claude/skills/exhume
cp skills/exhume/SKILL.md ~/.claude/skills/exhume/SKILL.md
```

**For a specific project** (project skill):

```bash
mkdir -p /path/to/your/project/.claude/skills/exhume
cp skills/exhume/SKILL.md /path/to/your/project/.claude/skills/exhume/SKILL.md
```

No additional configuration needed -- Claude Code auto-discovers skills from these locations. Once installed, Claude will automatically invoke exhume when it encounters `.xlsx` files or needs to extract embedded objects.

## CLI Usage

All commands default to JSON output (agent-friendly). Add `--output text` for human-readable output.

### Inspect workbook structure

```bash
exhume info spreadsheet.xlsx --pretty
```

Example output:

```json
{
  "filename": "vendor-catalog-q1.xlsx",
  "sheetCount": 3,
  "sheets": [
    {
      "name": "Products",
      "index": 0,
      "dimensions": "A1:F120",
      "maxRow": 120,
      "maxColumn": 6,
      "mergedCells": []
    },
    {
      "name": "Pricing",
      "index": 1,
      "dimensions": "A1:D45",
      "maxRow": 45,
      "maxColumn": 4,
      "mergedCells": ["A1:D1"]
    }
  ],
  "namedRanges": []
}
```

### Read cell data

```bash
# All sheets
exhume cells spreadsheet.xlsx --pretty

# Specific sheet
exhume cells spreadsheet.xlsx --sheet "Products" --pretty
```

Example output:

```json
[
  {
    "coordinate": "A1",
    "value": "SKU",
    "dataType": "string",
    "formula": null,
    "numberFormat": "General"
  },
  {
    "coordinate": "B1",
    "value": "Description",
    "dataType": "string",
    "formula": null,
    "numberFormat": "General"
  },
  {
    "coordinate": "C1",
    "value": "Unit Price",
    "dataType": "number",
    "formula": null,
    "numberFormat": "#,##0.00"
  }
]
```

### List embedded objects

```bash
exhume objects spreadsheet.xlsx --pretty
```

Returns metadata for every embedded OLE object: filename, sheet, row, column, size, and neighboring cell values for context.

Example output:

```json
{
  "file": "vendor-catalog-q1.xlsx",
  "totalObjects": 24,
  "objects": [
    {
      "index": 1,
      "filename": "product-spec-A100.txt",
      "sheet": "Products",
      "row": 4,
      "column": 5,
      "shapeId": 1028,
      "oleSource": "oleObject1.bin",
      "sizeBytes": 12480,
      "progId": "Packager Shell Object",
      "neighborCells": {
        "A4": "A100",
        "B4": "Wireless Router",
        "C4": 49.99,
        "D4": "2025-03-10 00:00:00",
        "F4": "Active"
      }
    }
  ]
}
```

### Extract embedded objects to disk

```bash
# Flat output (default)
exhume extract spreadsheet.xlsx --out ./extracted

# Organized by row
exhume extract spreadsheet.xlsx --out ./extracted --by-row
```

Example output:

```json
{
  "outputDir": "./extracted",
  "totalExtracted": 24,
  "objects": [
    {
      "index": 1,
      "filename": "product-spec-A100.txt",
      "sheet": "Products",
      "row": 4,
      "column": 5,
      "shapeId": 1028,
      "oleSource": "oleObject1.bin",
      "sizeBytes": 12480,
      "progId": "Packager Shell Object",
      "neighborCells": {
        "A4": "A100",
        "B4": "Wireless Router",
        "C4": 49.99
      }
    }
  ]
}
```

### Full export (data + metadata + objects)

```bash
# JSON format
exhume convert spreadsheet.xlsx --format json --out ./output

# CSV format
exhume convert spreadsheet.xlsx --format csv --out ./output
```

Example output:

```json
{"outputDir": "./output", "format": "json", "status": "complete"}
```

Produces a directory structure:

```
output/
  metadata.json
  sheets/
    Products/
      data.json (or data.csv)
      metadata.json
  objects/
    manifest.json
    files/
      product-spec-A100.txt
      product-spec-B200.txt
      ...
```

### Global flags

| Flag | Default | Description |
|------|---------|-------------|
| `--output json\|text` | `json` | Output format |
| `--sheet <name\|index>` | all | Target specific sheet |
| `--pretty` | off | Indented JSON |
| `--quiet` | off | Suppress progress output |

## Library Usage

```python
from exhume import Workbook

wb = Workbook("spreadsheet.xlsx")

# Inspect structure
info = wb.info()
print(info.sheet_count, info.sheets[0].name)

# Read cells
cells = wb.cells("Sheet1")
for cell in cells:
    print(cell.coordinate, cell.value, cell.data_type)

# List embedded objects with metadata
objects = wb.list_objects()
for obj in objects:
    print(obj.filename, obj.row, obj.neighbor_cells)

# Extract objects to disk
wb.extract_objects("./output", by_row=True)
```

## How it works

Excel `.xlsx` files are ZIP archives. Embedded objects (like `.txt` files) are stored as OLE compound documents inside `xl/embeddings/`. Exhume navigates four layers of indirection to map them:

1. VML drawings (`vmlDrawing*.vml`) -- shape anchors with row/column positions
2. Drawing relationships (`_rels/`) -- shape-to-image mappings
3. Worksheet relationships (`_rels/`) -- rId-to-oleObject.bin mappings
4. OLE compound documents -- `Ole10Native` stream containing the actual file content

## Development

```bash
# Install dev dependencies
pip install -e ".[dev]"

# Run tests
pytest -v

# Run with coverage
pytest --cov=exhume -v
```

## Dependencies

- [click](https://click.palletsprojects.com/) -- CLI framework
- [openpyxl](https://openpyxl.readthedocs.io/) -- .xlsx reading
- [olefile](https://olefile.readthedocs.io/) -- OLE compound document parsing
