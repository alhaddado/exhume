import json

import pytest
from click.testing import CliRunner

from exhume.cli import main


@pytest.fixture
def runner():
    return CliRunner()


class TestInfoCommand:
    def test_json_output(self, runner, tmp_xlsx):
        result = runner.invoke(main, ["info", str(tmp_xlsx)])
        assert result.exit_code == 0
        data = json.loads(result.output)
        assert data["filename"] == "test.xlsx"
        assert data["sheetCount"] == 2

    def test_text_output(self, runner, tmp_xlsx):
        result = runner.invoke(main, ["info", str(tmp_xlsx), "--output", "text"])
        assert result.exit_code == 0
        assert "test.xlsx" in result.output
        assert "Sheet1" in result.output

    def test_pretty_json(self, runner, tmp_xlsx):
        result = runner.invoke(main, ["info", str(tmp_xlsx), "--pretty"])
        assert result.exit_code == 0
        assert "\n" in result.output

    def test_file_not_found(self, runner, tmp_path):
        result = runner.invoke(main, ["info", str(tmp_path / "nope.xlsx")])
        assert result.exit_code != 0


class TestCellsCommand:
    def test_json_output(self, runner, tmp_xlsx):
        result = runner.invoke(main, ["cells", str(tmp_xlsx), "--sheet", "Sheet1"])
        assert result.exit_code == 0
        data = json.loads(result.output)
        assert isinstance(data, list)
        assert len(data) > 0

    def test_all_sheets(self, runner, tmp_xlsx):
        result = runner.invoke(main, ["cells", str(tmp_xlsx)])
        assert result.exit_code == 0
        data = json.loads(result.output)
        assert isinstance(data, dict)
        assert "Sheet1" in data
        assert "Sheet2" in data

    def test_text_output(self, runner, tmp_xlsx):
        result = runner.invoke(main, ["cells", str(tmp_xlsx), "--sheet", "Sheet1", "--output", "text"])
        assert result.exit_code == 0
        assert "A1" in result.output


class TestObjectsCommand:
    def test_json_output_empty(self, runner, tmp_xlsx):
        result = runner.invoke(main, ["objects", str(tmp_xlsx)])
        assert result.exit_code == 0
        data = json.loads(result.output)
        assert data["totalObjects"] == 0
        assert data["objects"] == []


class TestExtractCommand:
    def test_extract_empty(self, runner, tmp_xlsx, tmp_path):
        out = tmp_path / "extracted"
        result = runner.invoke(main, ["extract", str(tmp_xlsx), "--out", str(out)])
        assert result.exit_code == 0


class TestConvertCommand:
    def test_convert_json(self, runner, tmp_xlsx, tmp_path):
        out = tmp_path / "converted"
        result = runner.invoke(main, ["convert", str(tmp_xlsx), "--format", "json", "--out", str(out)])
        assert result.exit_code == 0
        assert (out / "metadata.json").exists()
        assert (out / "sheets" / "Sheet1" / "data.json").exists()

    def test_convert_csv(self, runner, tmp_xlsx, tmp_path):
        out = tmp_path / "converted"
        result = runner.invoke(main, ["convert", str(tmp_xlsx), "--format", "csv", "--out", str(out)])
        assert result.exit_code == 0
        assert (out / "sheets" / "Sheet1" / "data.csv").exists()
