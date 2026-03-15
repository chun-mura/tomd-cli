"""Core conversion logic."""

from __future__ import annotations

import re
from typing import Any
from pathlib import Path

import pdfplumber
from markitdown import MarkItDown

SUPPORTED_EXTENSIONS = {".pptx", ".docx", ".xlsx", ".pdf"}


def strip_base64_images(text: str) -> str:
    """Replace embedded base64 image data with a placeholder."""
    return re.sub(
        r'!\[([^\]]*)\]\(data:image/[^)]+\)',
        lambda m: f'![{m.group(1)}]()',
        text,
    )


def _clean_cell(value: str | None) -> str:
    """Clean a table cell value for Markdown output."""
    if value is None:
        return ""
    text = value.replace("\n", " ").strip()
    text = text.replace("\x00", "-")
    text = text.replace("|", "\\|")
    return text


def _detect_table_lines(
    rects: list[dict[str, Any]],
) -> tuple[list[float], list[float]]:
    """Detect vertical and horizontal table lines from PDF rectangles."""
    v_lines: set[float] = set()
    h_lines: set[float] = set()

    for r in rects:
        w = r["x1"] - r["x0"]
        h = r["bottom"] - r["top"]
        # Thin vertical rects are column borders
        if w < 2 and h > 10:
            v_lines.add(round((r["x0"] + r["x1"]) / 2, 1))
            # Top/bottom of vertical borders mark row boundaries
            h_lines.add(round(r["top"], 1))
            h_lines.add(round(r["bottom"], 1))
        # Thin horizontal rects are row borders
        if h < 2 and w > 10:
            h_lines.add(round((r["top"] + r["bottom"]) / 2, 1))
        # Cell rects — use top/bottom as horizontal lines
        if w > 50 and h > 20:
            h_lines.add(round(r["top"], 1))
            h_lines.add(round(r["bottom"], 1))

    return sorted(v_lines), sorted(h_lines)


def _extract_tables_from_page(
    page: Any,
) -> list[list[list[str | None]]]:
    """Extract tables from a single PDF page using rect-based line detection."""
    rects = page.rects
    if not rects:
        return page.extract_tables() or []

    v_lines, h_lines = _detect_table_lines(rects)
    if len(v_lines) < 2 or len(h_lines) < 2:
        return page.extract_tables() or []

    settings = {
        "explicit_vertical_lines": v_lines,
        "explicit_horizontal_lines": h_lines,
    }
    return page.extract_tables(table_settings=settings) or []


def _tables_have_same_header(
    a: list[list[str | None]], b: list[list[str | None]],
) -> bool:
    """Check if two tables share the same header row (cross-page merge)."""
    if not a or not b:
        return False
    header_a = [_clean_cell(c) for c in a[0]]
    header_b = [_clean_cell(c) for c in b[0]]
    return header_a == header_b and len(header_a) > 1


def _merge_continuation_row(
    last_row: list[str | None], cont_row: list[str | None],
) -> list[str | None]:
    """Merge a continuation row (with empty leading cells) into the previous row."""
    merged = list(last_row)
    for i, cell in enumerate(cont_row):
        if cell and cell.strip():
            if merged[i] and merged[i].strip():
                merged[i] = (merged[i] or "") + " " + cell
            else:
                merged[i] = cell
    return merged


def _extract_pdf_tables(pdf_path: str | Path) -> list[list[list[str | None]]]:
    """Extract and merge tables across pages from a PDF."""
    all_tables: list[list[list[str | None]]] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            page_tables = _extract_tables_from_page(page)
            for table in page_tables:
                if not table:
                    continue

                if all_tables and _tables_have_same_header(all_tables[-1], table):
                    prev = all_tables[-1]
                    for row in table[1:]:
                        non_empty = sum(1 for c in row if c and c.strip())
                        if non_empty < len(row) and non_empty > 0 and prev:
                            prev[-1] = _merge_continuation_row(prev[-1], row)
                        else:
                            prev.append(row)
                else:
                    all_tables.append(table)

    return all_tables


def _table_to_markdown(table: list[list[str | None]]) -> str:
    """Convert an extracted table to a Markdown table string."""
    if not table or len(table) < 2:
        return ""

    header = [_clean_cell(c) for c in table[0]]
    separator = ["-" * max(3, len(h)) for h in header]

    rows = []
    for row in table[1:]:
        cells = [_clean_cell(c) for c in row]
        while len(cells) < len(header):
            cells.append("")
        cells = cells[: len(header)]
        if not any(cells):
            continue
        rows.append(cells)

    if not rows:
        return ""

    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(separator) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)


def _get_table_bboxes(page: Any) -> list[tuple[float, ...]]:
    """Get bounding boxes covering all table rects on a page."""
    rects = page.rects
    if not rects:
        tables = page.find_tables()
        return [t.bbox for t in tables]

    v_lines, h_lines = _detect_table_lines(rects)
    if len(v_lines) < 2 or len(h_lines) < 2:
        tables = page.find_tables()
        return [t.bbox for t in tables]

    # Build bbox from the detected lines
    x0 = min(v_lines)
    x1 = max(v_lines)
    y0 = min(h_lines)
    y1 = max(h_lines)
    return [(x0, y0, x1, y1)]


def _extract_non_table_text(
    page: Any,
    table_bboxes: list[tuple[float, ...]],
) -> str:
    """Extract text from a page excluding table regions."""
    if not table_bboxes:
        text = page.extract_text()
        return text.strip() if text else ""

    chars_outside = []
    for char in page.chars:
        in_table = False
        for bbox in table_bboxes:
            x0, top, x1, bottom = bbox
            if (char["x0"] >= x0 - 2 and char["x1"] <= x1 + 2
                    and char["top"] >= top - 2 and char["bottom"] <= bottom + 2):
                in_table = True
                break
        if not in_table:
            chars_outside.append(char)

    if not chars_outside:
        return ""

    lines_map: dict[float, list[dict[str, Any]]] = {}
    for char in chars_outside:
        key = round(char["top"], 1)
        lines_map.setdefault(key, []).append(char)

    text_lines = []
    for key in sorted(lines_map.keys()):
        line_chars = sorted(lines_map[key], key=lambda c: c["x0"])
        text_lines.append("".join(c["text"] for c in line_chars).strip())

    return "\n".join(line for line in text_lines if line)


def _convert_pdf_with_tables(src: Path) -> str:
    """Convert a PDF, extracting tables as proper Markdown tables."""
    tables = _extract_pdf_tables(src)

    parts: list[str] = []

    with pdfplumber.open(str(src)) as pdf:
        # Track which merged table starts on which page
        merged_table_idx = 0
        table_start_pages: list[int] = []
        pages_with_table_start: set[int] = set()

        temp_idx = 0
        for page_idx, page in enumerate(pdf.pages):
            page_tables = _extract_tables_from_page(page)
            for pt in page_tables:
                if not pt:
                    continue
                if temp_idx < len(tables):
                    header_pt = [_clean_cell(c) for c in pt[0]]
                    header_merged = [_clean_cell(c) for c in tables[temp_idx][0]]
                    if header_pt == header_merged:
                        if page_idx not in pages_with_table_start:
                            table_start_pages.append(page_idx)
                            pages_with_table_start.add(page_idx)
                            temp_idx += 1

        merged_table_mds = [_table_to_markdown(t) for t in tables]

        merged_table_idx = 0
        for page_idx, page in enumerate(pdf.pages):
            table_bboxes = _get_table_bboxes(page)
            non_table_text = _extract_non_table_text(page, table_bboxes)

            if non_table_text:
                parts.append(non_table_text)

            if page_idx in pages_with_table_start:
                while (merged_table_idx < len(table_start_pages)
                       and table_start_pages[merged_table_idx] == page_idx):
                    md_table = merged_table_mds[merged_table_idx]
                    if md_table:
                        parts.append(md_table)
                    merged_table_idx += 1

            if not table_bboxes and not non_table_text:
                page_text = page.extract_text()
                if page_text and page_text.strip():
                    parts.append(page_text.strip())

    return "\n\n".join(parts)


def convert_file(
    input_path: str | Path,
    output_path: str | Path | None = None,
) -> Path:
    """Convert a single file to Markdown.

    Returns the path to the generated Markdown file.
    Raises FileNotFoundError if input does not exist.
    """
    src = Path(input_path).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"File not found: {input_path}")

    dest = Path(output_path).resolve() if output_path else src.with_suffix(".md")

    if src.suffix.lower() == ".pdf":
        content = _convert_pdf_with_tables(src)
    else:
        md = MarkItDown()
        result = md.convert(str(src))
        content = result.text_content

    content = strip_base64_images(content)

    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(content, encoding="utf-8")
    return dest


def convert_dir(
    input_dir: str | Path,
    output_dir: str | Path | None = None,
) -> list[Path]:
    """Convert all supported files in a directory to Markdown.

    Returns a list of generated Markdown file paths.
    Raises NotADirectoryError if input is not a directory.
    """
    src_dir = Path(input_dir).resolve()
    if not src_dir.is_dir():
        raise NotADirectoryError(f"Directory not found: {input_dir}")

    dest_dir = Path(output_dir).resolve() if output_dir else src_dir / "output"
    dest_dir.mkdir(parents=True, exist_ok=True)

    files = sorted(
        f for f in src_dir.iterdir() if f.suffix.lower() in SUPPORTED_EXTENSIONS
    )
    if not files:
        return []

    results: list[Path] = []
    for f in files:
        dest = dest_dir / f.with_suffix(".md").name
        if f.suffix.lower() == ".pdf":
            text = _convert_pdf_with_tables(f)
        else:
            md = MarkItDown()
            result = md.convert(str(f))
            text = result.text_content
        dest.write_text(strip_base64_images(text), encoding="utf-8")
        results.append(dest)

    return results
