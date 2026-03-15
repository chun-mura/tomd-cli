"""Tests for tomd.converter."""

import pytest
from tomd.converter import (
    strip_base64_images,
    _clean_cell,
    _detect_table_lines,
    _table_to_markdown,
    _tables_have_same_header,
    _merge_continuation_row,
    _collect_table_cell_texts,
    _find_table_region,
    _strip_page_headers,
    _apply_headings,
    _apply_bullets,
)


class TestStripBase64Images:
    def test_replaces_base64_png(self):
        text = '![alt text](data:image/png;base64,iVBORw0KGgoAAAANSUhEUg==)'
        result = strip_base64_images(text)
        assert result == '![alt text]()'

    def test_replaces_base64_jpeg(self):
        text = '![photo](data:image/jpeg;base64,/9j/4AAQSkZJRg==)'
        result = strip_base64_images(text)
        assert result == '![photo]()'

    def test_preserves_empty_alt(self):
        text = '![](data:image/png;base64,abc123)'
        result = strip_base64_images(text)
        assert result == '![]()'

    def test_preserves_normal_images(self):
        text = '![alt](https://example.com/image.png)'
        result = strip_base64_images(text)
        assert result == '![alt](https://example.com/image.png)'

    def test_preserves_plain_text(self):
        text = 'Hello world'
        result = strip_base64_images(text)
        assert result == 'Hello world'

    def test_multiple_images(self):
        text = (
            '![a](data:image/png;base64,AAA) middle '
            '![b](data:image/gif;base64,BBB)'
        )
        result = strip_base64_images(text)
        assert result == '![a]() middle ![b]()'


class TestCleanCell:
    def test_none(self):
        assert _clean_cell(None) == ""

    def test_newlines_replaced(self):
        assert _clean_cell("hello\nworld") == "hello world"

    def test_null_byte_replaced(self):
        assert _clean_cell("GPT\x004o") == "GPT-4o"

    def test_pipe_escaped(self):
        assert _clean_cell("a|b") == "a\\|b"

    def test_strips_whitespace(self):
        assert _clean_cell("  text  ") == "text"


class TestDetectTableLines:
    def test_vertical_thin_rects(self):
        rects = [
            {"x0": 10.0, "x1": 10.5, "top": 100.0, "bottom": 300.0},
            {"x0": 200.0, "x1": 200.5, "top": 100.0, "bottom": 300.0},
        ]
        v, h = _detect_table_lines(rects)
        assert len(v) == 2
        assert 100.0 in h and 300.0 in h

    def test_horizontal_thin_rects(self):
        rects = [
            {"x0": 10.0, "x1": 200.0, "top": 100.0, "bottom": 100.5},
        ]
        v, h = _detect_table_lines(rects)
        assert len(v) == 0
        assert len(h) == 1

    def test_cell_rects(self):
        rects = [
            {"x0": 10.0, "x1": 100.0, "top": 50.0, "bottom": 80.0},
        ]
        v, h = _detect_table_lines(rects)
        assert 50.0 in h and 80.0 in h


class TestTableToMarkdown:
    def test_basic_table(self):
        table = [
            ["A", "B"],
            ["1", "2"],
            ["3", "4"],
        ]
        result = _table_to_markdown(table)
        assert "| A | B |" in result
        assert "| 1 | 2 |" in result
        assert "| 3 | 4 |" in result

    def test_skips_empty_rows(self):
        table = [
            ["A", "B"],
            [None, None],
            ["1", "2"],
        ]
        result = _table_to_markdown(table)
        lines = result.strip().split("\n")
        assert len(lines) == 3  # header + separator + 1 data row

    def test_pads_short_rows(self):
        table = [
            ["A", "B", "C"],
            ["1"],
        ]
        result = _table_to_markdown(table)
        assert "| 1 |  |  |" in result

    def test_empty_table(self):
        assert _table_to_markdown([]) == ""
        assert _table_to_markdown([["A"]]) == ""


class TestTablesHaveSameHeader:
    def test_same_header(self):
        a = [["A", "B"], ["1", "2"]]
        b = [["A", "B"], ["3", "4"]]
        assert _tables_have_same_header(a, b)

    def test_different_header(self):
        a = [["A", "B"], ["1", "2"]]
        b = [["X", "Y"], ["3", "4"]]
        assert not _tables_have_same_header(a, b)

    def test_single_column_not_merged(self):
        a = [["A"], ["1"]]
        b = [["A"], ["2"]]
        assert not _tables_have_same_header(a, b)

    def test_empty_tables(self):
        assert not _tables_have_same_header([], [["A"]])


class TestMergeContinuationRow:
    def test_merges_non_empty_cells(self):
        last = ["hello", "world"]
        cont = [None, " again"]
        result = _merge_continuation_row(last, cont)
        assert result == ["hello", "world  again"]

    def test_fills_empty_cells(self):
        last = [None, "existing"]
        cont = ["new", None]
        result = _merge_continuation_row(last, cont)
        assert result == ["new", "existing"]


class TestCollectTableCellTexts:
    def test_collects_cell_lines(self):
        tables = [[["Header A", "Header B"], ["cell\nwith newline", "short"]]]
        result = _collect_table_cell_texts(tables)
        assert "Header A" in result
        assert "cell" in result
        assert "with newline" in result
        assert "short" in result

    def test_skips_empty_cells(self):
        tables = [[[None, ""], ["text", None]]]
        result = _collect_table_cell_texts(tables)
        assert "text" in result
        assert "" not in result

    def test_replaces_null_bytes(self):
        tables = [[["GPT\x004o", "other"]]]
        result = _collect_table_cell_texts(tables)
        assert "GPT-4o" in result


class TestFindTableRegion:
    def test_finds_region(self):
        lines = [
            "Title",
            "",
            "Header A",
            "cell data",
            "more cell",
            "",
            "After table text",
        ]
        cell_texts = {"Header A", "cell data", "more cell"}
        start, end = _find_table_region(lines, cell_texts)
        assert start == 2
        assert end == 4

    def test_no_match(self):
        lines = ["Title", "No table here"]
        cell_texts = {"something else entirely"}
        start, end = _find_table_region(lines, cell_texts)
        assert start is None
        assert end is None


class TestStripPageHeaders:
    def test_removes_repeated_lines(self):
        text = "Header\nContent A\nHeader\nContent B\nHeader"
        result = _strip_page_headers(text)
        assert "Header" not in result
        assert "Content A" in result

    def test_removes_page_numbers(self):
        text = "Line 1\nLine 2\nLine 3\nLine 4\nAIモデル精度チューニング観点2"
        result = _strip_page_headers(text)
        assert "AIモデル精度チューニング観点2" not in result

    def test_cleans_header_prefix(self):
        text = "Line 1\nLine 2\nLine 3\nLine 4\nAIモデル精度チューニング観点2max_tokens（最大出力量）"
        result = _strip_page_headers(text)
        assert "max_tokens（最大出力量）" in result

    def test_preserves_short_text(self):
        text = "Short"
        result = _strip_page_headers(text)
        assert result == "Short"


class TestApplyHeadings:
    def test_applies_h1_and_h2(self):
        heading_map = {"Title": 1, "Section": 2}
        text = "Title\n\nSection\n\nBody text"
        result = _apply_headings(text, heading_map)
        assert "# Title" in result
        assert "## Section" in result
        assert "Body text" in result

    def test_no_headings(self):
        text = "Just body text"
        result = _apply_headings(text, {})
        assert result == text

    def test_does_not_affect_non_matching_lines(self):
        heading_map = {"Title": 1}
        text = "Title\nNot a heading\nMore text"
        result = _apply_headings(text, heading_map)
        assert result == "# Title\nNot a heading\nMore text"


class TestApplyBullets:
    def test_exact_match(self):
        items = {"temperature", "top-k"}
        text = "temperature\n\nbody text\n\ntop-k"
        result = _apply_bullets(text, items)
        assert "- temperature" in result
        assert "- top-k" in result
        assert "body text" in result
        assert "- body" not in result

    def test_startswith_match(self):
        items = {"temperature"}
        text = "temperature（温度）\n\nbody text"
        result = _apply_bullets(text, items)
        assert "- temperature（温度）" in result

    def test_skips_parameter_values(self):
        items = {"temperature"}
        text = "temperature=0.2, top_p=0.9"
        result = _apply_bullets(text, items)
        assert not result.startswith("- ")

    def test_skips_headings_and_tables(self):
        items = {"Section"}
        text = "## Section\n| Section | col |"
        result = _apply_bullets(text, items)
        assert "## Section" in result
        assert "- ## Section" not in result
        assert "- | Section" not in result

    def test_no_items(self):
        text = "Just text"
        result = _apply_bullets(text, set())
        assert result == text
