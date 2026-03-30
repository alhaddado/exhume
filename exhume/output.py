"""Output formatting for CLI: JSON and human-readable text."""

from __future__ import annotations

import json
from typing import Any

from exhume.models import CellInfo, EmbeddedObject, SheetInfo, WorkbookInfo


def _to_serializable(obj: Any) -> Any:
    if hasattr(obj, "to_dict"):
        return obj.to_dict()
    if isinstance(obj, list):
        return [_to_serializable(v) for v in obj]
    if isinstance(obj, dict):
        return {k: _to_serializable(v) for k, v in obj.items()}
    return obj


def format_json(data: Any, pretty: bool = False) -> str:
    """Serialize data to JSON string."""
    serializable = _to_serializable(data)
    indent = 2 if pretty else None
    return json.dumps(serializable, indent=indent, default=str, ensure_ascii=False)


def format_text(data: Any) -> str:
    """Serialize data to human-readable text."""
    if isinstance(data, WorkbookInfo):
        return _format_workbook_text(data)
    if isinstance(data, list) and data:
        if isinstance(data[0], CellInfo):
            return _format_cells_text(data)
        if isinstance(data[0], EmbeddedObject):
            return _format_objects_text(data)
        if isinstance(data[0], SheetInfo):
            return _format_sheets_text(data)
    if isinstance(data, list) and not data:
        return "No results."
    return str(data)


def _format_workbook_text(info: WorkbookInfo) -> str:
    lines = [
        f"File: {info.filename}",
        f"Sheets: {info.sheet_count}",
        "",
    ]
    for s in info.sheets:
        merged = f", merged: {', '.join(s.merged_cells)}" if s.merged_cells else ""
        lines.append(f"  [{s.index}] {s.name}  {s.dimensions}  ({s.max_row} rows x {s.max_column} cols{merged})")
    if info.named_ranges:
        lines.append("")
        lines.append("Named ranges:")
        for nr in info.named_ranges:
            lines.append(f"  {nr}")
    return "\n".join(lines)


def _format_cells_text(cells: list[CellInfo]) -> str:
    lines = []
    for c in cells:
        formula_part = f"  formula={c.formula}" if c.formula else ""
        lines.append(f"  {c.coordinate:<6} {str(c.value):<30} ({c.data_type}){formula_part}")
    return "\n".join(lines)


def _format_objects_text(objects: list[EmbeddedObject]) -> str:
    lines = []
    for obj in objects:
        lines.append(f"  {obj.index:>3}. {obj.filename:<50} sheet={obj.sheet} row={obj.row} ({obj.size_bytes} bytes)")
    return "\n".join(lines)


def _format_sheets_text(sheets: list[SheetInfo]) -> str:
    lines = []
    for s in sheets:
        lines.append(f"  [{s.index}] {s.name}  {s.dimensions}")
    return "\n".join(lines)
