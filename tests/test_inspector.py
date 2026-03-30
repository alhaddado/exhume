from exhume.inspector import inspect_workbook, inspect_sheet_cells


class TestInspectWorkbook:
    def test_returns_workbook_info(self, tmp_xlsx):
        info = inspect_workbook(tmp_xlsx)
        assert info.filename == "test.xlsx"
        assert info.sheet_count == 2

    def test_sheet_names(self, tmp_xlsx):
        info = inspect_workbook(tmp_xlsx)
        names = [s.name for s in info.sheets]
        assert names == ["Sheet1", "Sheet2"]

    def test_sheet_dimensions(self, tmp_xlsx):
        info = inspect_workbook(tmp_xlsx)
        sheet1 = info.sheets[0]
        assert sheet1.max_row >= 3
        assert sheet1.max_column >= 3

    def test_merged_cells(self, tmp_xlsx):
        info = inspect_workbook(tmp_xlsx)
        sheet1 = info.sheets[0]
        assert "A5:C5" in sheet1.merged_cells


class TestInspectSheetCells:
    def test_returns_cell_list(self, tmp_xlsx):
        cells = inspect_sheet_cells(tmp_xlsx, "Sheet1")
        assert len(cells) > 0

    def test_cell_values(self, tmp_xlsx):
        cells = inspect_sheet_cells(tmp_xlsx, "Sheet1")
        by_coord = {c.coordinate: c for c in cells}
        assert by_coord["A1"].value == "ID"
        assert by_coord["A2"].value == 1
        assert by_coord["C2"].value == 10.5

    def test_cell_types(self, tmp_xlsx):
        cells = inspect_sheet_cells(tmp_xlsx, "Sheet1")
        by_coord = {c.coordinate: c for c in cells}
        assert by_coord["A1"].data_type == "string"
        assert by_coord["A2"].data_type == "number"

    def test_cell_formula(self, tmp_xlsx):
        cells = inspect_sheet_cells(tmp_xlsx, "Sheet1")
        by_coord = {c.coordinate: c for c in cells}
        assert by_coord["D3"].formula == "=C2+C3"

    def test_sheet_by_index(self, tmp_xlsx):
        cells = inspect_sheet_cells(tmp_xlsx, 0)
        by_coord = {c.coordinate: c for c in cells}
        assert by_coord["A1"].value == "ID"

    def test_sheet_not_found(self, tmp_xlsx):
        import pytest
        with pytest.raises(ValueError, match="Sheet .* not found"):
            inspect_sheet_cells(tmp_xlsx, "NonExistent")
