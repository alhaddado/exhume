from exhume.models import CellInfo, SheetInfo, WorkbookInfo, EmbeddedObject


class TestCellInfo:
    def test_to_dict(self):
        cell = CellInfo(
            coordinate="A1",
            value="hello",
            data_type="string",
            formula=None,
            number_format="General",
        )
        d = cell.to_dict()
        assert d == {
            "coordinate": "A1",
            "value": "hello",
            "dataType": "string",
            "formula": None,
            "numberFormat": "General",
        }

    def test_to_dict_with_formula(self):
        cell = CellInfo(
            coordinate="C3",
            value=30.5,
            data_type="number",
            formula="=C2+C3",
            number_format="0.00",
        )
        d = cell.to_dict()
        assert d["formula"] == "=C2+C3"
        assert d["value"] == 30.5


class TestSheetInfo:
    def test_to_dict(self):
        sheet = SheetInfo(
            name="Sheet1",
            index=0,
            dimensions="A1:C10",
            max_row=10,
            max_column=3,
            merged_cells=["A5:C5"],
        )
        d = sheet.to_dict()
        assert d["name"] == "Sheet1"
        assert d["index"] == 0
        assert d["maxRow"] == 10
        assert d["maxColumn"] == 3
        assert d["mergedCells"] == ["A5:C5"]


class TestWorkbookInfo:
    def test_to_dict(self):
        sheet = SheetInfo(
            name="Sheet1", index=0, dimensions="A1:C10",
            max_row=10, max_column=3, merged_cells=[],
        )
        info = WorkbookInfo(
            filename="test.xlsx",
            sheet_count=1,
            sheets=[sheet],
            named_ranges=[],
        )
        d = info.to_dict()
        assert d["filename"] == "test.xlsx"
        assert d["sheetCount"] == 1
        assert len(d["sheets"]) == 1
        assert d["sheets"][0]["name"] == "Sheet1"


class TestEmbeddedObject:
    def test_to_dict(self):
        obj = EmbeddedObject(
            index=1,
            filename="test.txt",
            sheet="Sheet1",
            row=48,
            column=5,
            shape_id=1232,
            ole_source="oleObject183.bin",
            size_bytes=2328,
            prog_id="Packager Shell Object",
            neighbor_cells={"A48": "48", "B48": "App/Site API"},
        )
        d = obj.to_dict()
        assert d["index"] == 1
        assert d["filename"] == "test.txt"
        assert d["row"] == 48
        assert d["shapeId"] == 1232
        assert d["oleSource"] == "oleObject183.bin"
        assert d["neighborCells"]["A48"] == "48"
