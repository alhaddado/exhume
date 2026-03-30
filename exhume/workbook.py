"""Workbook facade -- main entry point for library and CLI."""

from __future__ import annotations

from pathlib import Path
from typing import Union

from exhume.inspector import inspect_workbook, inspect_sheet_cells
from exhume.extractor import extract_objects_metadata, extract_objects_to_disk
from exhume.models import WorkbookInfo, CellInfo, EmbeddedObject


class Workbook:
    """Facade for introspecting and extracting from an Excel workbook."""

    def __init__(self, path: Union[str, Path]):
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(f"File not found: {self.path}")

    def info(self) -> WorkbookInfo:
        """Return structural info about the workbook."""
        return inspect_workbook(self.path)

    def sheet_names(self) -> list[str]:
        """Return list of sheet names."""
        return [s.name for s in self.info().sheets]

    def cells(self, sheet_ref: Union[str, int]) -> list[CellInfo]:
        """Return all non-empty cells in a sheet (by name or index)."""
        return inspect_sheet_cells(self.path, sheet_ref)

    def list_objects(self, sheet_name: str | None = None) -> list[EmbeddedObject]:
        """List all embedded OLE objects with metadata."""
        return extract_objects_metadata(self.path, sheet_name)

    def extract_objects(
        self,
        out_dir: Union[str, Path],
        by_row: bool = False,
        clean: bool = True,
    ) -> list[EmbeddedObject]:
        """Extract embedded objects to disk."""
        return extract_objects_to_disk(self.path, out_dir, by_row=by_row, clean=clean)
