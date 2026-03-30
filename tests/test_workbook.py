import pytest

from exhume import Workbook
from exhume.models import WorkbookInfo, SheetInfo, CellInfo


class TestWorkbook:
    def test_info(self, tmp_xlsx):
        wb = Workbook(tmp_xlsx)
        info = wb.info()
        assert isinstance(info, WorkbookInfo)
        assert info.sheet_count == 2

    def test_sheet_names(self, tmp_xlsx):
        wb = Workbook(tmp_xlsx)
        assert wb.sheet_names() == ["Sheet1", "Sheet2"]

    def test_sheet_cells(self, tmp_xlsx):
        wb = Workbook(tmp_xlsx)
        cells = wb.cells("Sheet1")
        assert len(cells) > 0
        by_coord = {c.coordinate: c for c in cells}
        assert by_coord["A1"].value == "ID"

    def test_sheet_cells_by_index(self, tmp_xlsx):
        wb = Workbook(tmp_xlsx)
        cells = wb.cells(0)
        by_coord = {c.coordinate: c for c in cells}
        assert by_coord["A1"].value == "ID"

    def test_list_objects_empty(self, tmp_xlsx):
        wb = Workbook(tmp_xlsx)
        assert wb.list_objects() == []

    def test_invalid_file(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            Workbook(tmp_path / "nonexistent.xlsx")

    def test_str_path(self, tmp_xlsx):
        wb = Workbook(str(tmp_xlsx))
        info = wb.info()
        assert info.sheet_count == 2
