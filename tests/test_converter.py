"""Tests for tomd.converter."""

import pytest
from tomd.converter import (
    strip_base64_images,
    _clean_cell,
    _detect_table_lines,
    _table_to_markdown,
    _tables_have_same_header,
    _merge_continuation_row,
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
