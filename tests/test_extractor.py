from exhume.extractor import parse_ole10native, extract_objects_metadata, extract_object_content


class TestParseOle10Native:
    def test_extracts_filename(self, ole10native_data):
        filename, content, raw = ole10native_data
        result_filename, result_content = parse_ole10native(raw)
        assert result_filename == "sample.txt"

    def test_extracts_content(self, ole10native_data):
        filename, content, raw = ole10native_data
        result_filename, result_content = parse_ole10native(raw)
        assert result_content == content

    def test_cleans_url_prefix(self, ole10native_data):
        filename, content, raw = ole10native_data
        _, result_content = parse_ole10native(raw, clean=True)
        assert result_content.startswith(b"https://api.example.com/test")


class TestExtractObjectsMetadata:
    def test_returns_empty_for_no_objects(self, tmp_xlsx):
        objects = extract_objects_metadata(tmp_xlsx)
        assert objects == []


class TestExtractObjectContent:
    def test_strips_ole_header(self, ole10native_data):
        filename, content, raw = ole10native_data
        result = extract_object_content(raw)
        assert result["filename"] == "sample.txt"
        assert result["content"] == content
