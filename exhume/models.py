"""Data models for Exhume. All dataclasses with camelCase JSON serialization."""

from __future__ import annotations

from dataclasses import dataclass, fields
from typing import Any


def _to_camel(name: str) -> str:
    parts = name.split("_")
    return parts[0] + "".join(p.capitalize() for p in parts[1:])


def _serialize(obj: Any) -> Any:
    if hasattr(obj, "to_dict"):
        return obj.to_dict()
    if isinstance(obj, list):
        return [_serialize(v) for v in obj]
    if isinstance(obj, dict):
        return {k: _serialize(v) for k, v in obj.items()}
    return obj


@dataclass
class CellInfo:
    coordinate: str
    value: Any
    data_type: str
    formula: str | None
    number_format: str

    def to_dict(self) -> dict[str, Any]:
        return {_to_camel(f.name): _serialize(getattr(self, f.name)) for f in fields(self)}


@dataclass
class SheetInfo:
    name: str
    index: int
    dimensions: str
    max_row: int
    max_column: int
    merged_cells: list[str]

    def to_dict(self) -> dict[str, Any]:
        return {_to_camel(f.name): _serialize(getattr(self, f.name)) for f in fields(self)}


@dataclass
class WorkbookInfo:
    filename: str
    sheet_count: int
    sheets: list[SheetInfo]
    named_ranges: list[str]

    def to_dict(self) -> dict[str, Any]:
        return {_to_camel(f.name): _serialize(getattr(self, f.name)) for f in fields(self)}


@dataclass
class EmbeddedObject:
    index: int
    filename: str
    sheet: str
    row: int
    column: int
    shape_id: int
    ole_source: str
    size_bytes: int
    prog_id: str
    neighbor_cells: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return {_to_camel(f.name): _serialize(getattr(self, f.name)) for f in fields(self)}
