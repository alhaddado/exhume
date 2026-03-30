import csv
import json

from exhume.converter import convert_workbook


class TestConvertJson:
    def test_creates_output_structure(self, tmp_xlsx, tmp_path):
        out_dir = tmp_path / "output"
        convert_workbook(tmp_xlsx, out_dir, fmt="json")

        assert (out_dir / "metadata.json").exists()
        assert (out_dir / "sheets" / "Sheet1" / "data.json").exists()
        assert (out_dir / "sheets" / "Sheet1" / "metadata.json").exists()
        assert (out_dir / "sheets" / "Sheet2" / "data.json").exists()

    def test_workbook_metadata(self, tmp_xlsx, tmp_path):
        out_dir = tmp_path / "output"
        convert_workbook(tmp_xlsx, out_dir, fmt="json")

        meta = json.loads((out_dir / "metadata.json").read_text())
        assert meta["filename"] == "test.xlsx"
        assert meta["sheetCount"] == 2

    def test_sheet_data_json(self, tmp_xlsx, tmp_path):
        out_dir = tmp_path / "output"
        convert_workbook(tmp_xlsx, out_dir, fmt="json")

        data = json.loads((out_dir / "sheets" / "Sheet1" / "data.json").read_text())
        assert "headers" in data
        assert "rows" in data
        assert len(data["rows"]) > 0

    def test_sheet_metadata(self, tmp_xlsx, tmp_path):
        out_dir = tmp_path / "output"
        convert_workbook(tmp_xlsx, out_dir, fmt="json")

        meta = json.loads((out_dir / "sheets" / "Sheet1" / "metadata.json").read_text())
        assert meta["name"] == "Sheet1"
        assert meta["maxRow"] >= 3

    def test_objects_manifest(self, tmp_xlsx, tmp_path):
        out_dir = tmp_path / "output"
        convert_workbook(tmp_xlsx, out_dir, fmt="json")

        manifest = json.loads((out_dir / "objects" / "manifest.json").read_text())
        assert "objects" in manifest
        assert manifest["totalObjects"] == 0  # test xlsx has no embedded objects


class TestConvertCsv:
    def test_creates_csv(self, tmp_xlsx, tmp_path):
        out_dir = tmp_path / "output"
        convert_workbook(tmp_xlsx, out_dir, fmt="csv")

        csv_path = out_dir / "sheets" / "Sheet1" / "data.csv"
        assert csv_path.exists()

    def test_csv_content(self, tmp_xlsx, tmp_path):
        out_dir = tmp_path / "output"
        convert_workbook(tmp_xlsx, out_dir, fmt="csv")

        csv_path = out_dir / "sheets" / "Sheet1" / "data.csv"
        with open(csv_path, newline="") as f:
            reader = csv.reader(f)
            rows = list(reader)
        # Should have at least header + data rows
        assert len(rows) >= 3
