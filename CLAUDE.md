# CLAUDE.md

## Project Overview

Exhume is a Python CLI tool and library for inspecting Excel (.xlsx) files and extracting embedded OLE objects. Designed for AI agent consumption with JSON-first output.

## Architecture

Layered library + CLI:

- `exhume/models.py` -- Dataclasses (CellInfo, SheetInfo, WorkbookInfo, EmbeddedObject) with camelCase `to_dict()` serialization
- `exhume/inspector.py` -- Sheet/cell introspection via openpyxl
- `exhume/extractor.py` -- OLE extraction pipeline (VML parsing, worksheet rel mapping, Ole10Native parsing)
- `exhume/workbook.py` -- Facade class (`Workbook`) tying inspector + extractor
- `exhume/output.py` -- JSON/text output formatting
- `exhume/converter.py` -- Full export to CSV/JSON directory structure
- `exhume/cli.py` -- Click CLI entry point with subcommands

## Common Commands

```bash
# Run tests
.venv/bin/pytest -v

# Run specific test file
.venv/bin/pytest tests/test_extractor.py -v

# Run with coverage
.venv/bin/pytest --cov=exhume -v

# Install in dev mode
pip install -e ".[dev]"

# Test CLI
.venv/bin/exhume --help
.venv/bin/exhume info <file> --pretty
```

## Key Patterns

- All models use snake_case fields internally, serialize to camelCase via `_to_camel()` in `models.py`
- `_to_serializable()` in `output.py` recursively converts nested models/lists/dicts for JSON output
- The OLE extraction pipeline in `extractor.py` handles 4 layers of indirection: VML shapes -> drawing rels -> worksheet rels -> oleObject .bin files
- `Ole10Native` binary format: 4-byte size + 2-byte flags + null-terminated filename + null-terminated src_path + 4-byte reserved + null-terminated temp_path + 4-byte data_size + content
- Test fixtures in `tests/conftest.py`: `tmp_xlsx` creates a minimal xlsx, `ole10native_data` creates synthetic OLE stream data

## Dependencies

- Python 3.10+
- click (CLI), openpyxl (.xlsx reading), olefile (OLE parsing)
- pytest, pytest-cov (dev)

## Notes

- The venv is at `.venv/` -- use `.venv/bin/pytest` and `.venv/bin/exhume` directly
- JSON output is the default for all CLI commands (optimized for AI agents)
- The `--output text` flag switches to human-readable format
