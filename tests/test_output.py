import json

from exhume.models import SheetInfo, WorkbookInfo, CellInfo, EmbeddedObject
from exhume.output import format_json, format_text


class TestFormatJson:
    def test_formats_workbook_info(self):
        info = WorkbookInfo(
            filename="test.xlsx", sheet_count=1,
            sheets=[SheetInfo(name="Sheet1", index=0, dimensions="A1:C3", max_row=3, max_column=3, merged_cells=[])],
            named_ranges=[],
        )
        result = format_json(info)
        parsed = json.loads(result)
        assert parsed["filename"] == "test.xlsx"

    def test_formats_list(self):
        cells = [
            CellInfo(coordinate="A1", value="hi", data_type="string", formula=None, number_format="General"),
        ]
        result = format_json(cells)
        parsed = json.loads(result)
        assert isinstance(parsed, list)
        assert parsed[0]["coordinate"] == "A1"

    def test_pretty_mode(self):
        info = WorkbookInfo(filename="t.xlsx", sheet_count=0, sheets=[], named_ranges=[])
        compact = format_json(info, pretty=False)
        pretty = format_json(info, pretty=True)
        assert "\n" not in compact
        assert "\n" in pretty


class TestFormatText:
    def test_workbook_info(self):
        info = WorkbookInfo(
            filename="test.xlsx", sheet_count=2,
            sheets=[
                SheetInfo(name="Sheet1", index=0, dimensions="A1:C3", max_row=3, max_column=3, merged_cells=["A5:C5"]),
                SheetInfo(name="Sheet2", index=1, dimensions="A1:A1", max_row=1, max_column=1, merged_cells=[]),
            ],
            named_ranges=[],
        )
        result = format_text(info)
        assert "test.xlsx" in result
        assert "Sheet1" in result
        assert "Sheet2" in result

    def test_cell_list(self):
        cells = [
            CellInfo(coordinate="A1", value="ID", data_type="string", formula=None, number_format="General"),
            CellInfo(coordinate="B1", value=42, data_type="number", formula=None, number_format="0"),
        ]
        result = format_text(cells)
        assert "A1" in result
        assert "ID" in result

    def test_embedded_objects(self):
        objs = [
            EmbeddedObject(
                index=1, filename="test.txt", sheet="Sheet1", row=5, column=3,
                shape_id=100, ole_source="oleObject1.bin", size_bytes=1024,
                prog_id="Packager Shell Object", neighbor_cells={"A5": "hello"},
            )
        ]
        result = format_text(objs)
        assert "test.txt" in result
        assert "Sheet1" in result
