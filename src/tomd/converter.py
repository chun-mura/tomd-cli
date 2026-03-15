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
        return []

    v_lines, h_lines = _detect_table_lines(rects)
    if len(v_lines) < 2 or len(h_lines) < 2:
        return []

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
                if not table or len(table) < 2:
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


def _collect_table_cell_texts(
    tables: list[list[list[str | None]]],
) -> set[str]:
    """Collect unique text fragments from table cells for matching."""
    cell_texts: set[str] = set()
    for table in tables:
        for row in table:
            for cell in row:
                if not cell or not cell.strip():
                    continue
                for line in cell.strip().split("\n"):
                    cleaned = line.strip().replace("\x00", "-")
                    if cleaned:
                        cell_texts.add(cleaned)
    return cell_texts


def _find_table_region(
    markitdown_lines: list[str],
    cell_texts: set[str],
) -> tuple[int | None, int | None]:
    """Find the start and end line indices of table content in MarkItDown output."""
    start: int | None = None
    end: int | None = None

    # Build a set of longer phrases for more reliable matching
    long_texts = {t for t in cell_texts if len(t) > 8}

    for i, line in enumerate(markitdown_lines):
        stripped = line.strip()
        if not stripped:
            continue
        # Check if this line matches a table cell
        is_table = stripped in cell_texts
        if not is_table:
            is_table = any(ct in stripped for ct in long_texts)
        if is_table:
            if start is None:
                start = i
            end = i

    return start, end


def _strip_page_headers(text: str) -> str:
    """Remove repeated page headers/footers and page numbers from text."""
    lines = text.split("\n")
    if len(lines) < 5:
        return text

    # Detect repeated lines (page headers/footers)
    line_counts: dict[str, int] = {}
    for line in lines:
        stripped = line.strip()
        if stripped:
            line_counts[stripped] = line_counts.get(stripped, 0) + 1

    repeated = {line for line, count in line_counts.items() if count >= 2}

    # Also detect page-number-like patterns (standalone digits, or "header+digit")
    page_num_pattern = re.compile(
        r'^(AIモデル精度チューニング観点)?\s*\d+\s*$'
    )

    filtered = []
    for line in lines:
        stripped = line.strip()
        if stripped in repeated:
            continue
        if page_num_pattern.match(stripped):
            continue
        # Clean lines like "AIモデル精度チューニング観点2max_tokens" -> "max_tokens"
        cleaned = re.sub(
            r'^AIモデル精度チューニング観点\d*',
            '',
            stripped,
        )
        if cleaned != stripped:
            if cleaned:
                filtered.append(cleaned)
            continue
        filtered.append(line)

    return "\n".join(filtered)


def _extract_heading_map(pdf_path: str | Path) -> dict[str, int]:
    """Build a mapping from text to heading level based on font size.

    Analyzes font sizes across all pages and assigns heading levels:
    - Largest font size → # (h1)
    - Second largest → ## (h2)
    - Third largest (if clearly distinct from body) → ### (h3)
    - Everything else → body text (no heading)
    """
    # Collect all font sizes and their associated text lines
    size_to_lines: dict[float, list[str]] = {}

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            if not page.chars:
                continue

            # Group chars into lines by top position
            lines_by_top: dict[float, list[dict[str, Any]]] = {}
            for char in page.chars:
                key = round(char["top"], 1)
                lines_by_top.setdefault(key, []).append(char)

            for top in sorted(lines_by_top.keys()):
                chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
                # Use the dominant (most common) font size for this line
                line_sizes = [round(c["size"], 1) for c in chars]
                if not line_sizes:
                    continue
                dominant_size = max(set(line_sizes), key=line_sizes.count)
                text = "".join(c["text"] for c in chars).strip()
                if text:
                    size_to_lines.setdefault(dominant_size, []).append(text)

    if not size_to_lines:
        return {}

    # Determine body text size (the size with the most total characters)
    size_char_counts: dict[float, int] = {}
    for size, texts in size_to_lines.items():
        size_char_counts[size] = sum(len(t) for t in texts)
    body_size = max(size_char_counts, key=size_char_counts.get)  # type: ignore[arg-type]

    # Only sizes significantly larger than body text are headings
    # (at least 1.3x body size to avoid false positives)
    heading_sizes = sorted(
        [s for s in size_to_lines if s > body_size * 1.3],
        reverse=True,
    )

    if not heading_sizes:
        return {}

    # Assign heading levels (max 2 levels to avoid over-tagging)
    heading_map: dict[str, int] = {}
    for level, size in enumerate(heading_sizes[:2], start=1):
        for text in size_to_lines[size]:
            heading_map[text] = level

    return heading_map


def _extract_bullet_items(pdf_path: str | Path) -> set[str]:
    """Detect lines that should be bullet list items based on bold font + indent.

    Lines starting with a bold font and indented from the left margin
    (outside table regions) are treated as bullet items.
    """
    bullet_texts: set[str] = set()

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            if not page.chars:
                continue

            # Determine table bbox to exclude table chars
            table_bbox: tuple[float, ...] | None = None
            if page.rects:
                v_lines, h_lines = _detect_table_lines(page.rects)
                if len(v_lines) >= 2 and len(h_lines) >= 2:
                    table_bbox = (
                        min(v_lines) - 2, min(h_lines) - 2,
                        max(v_lines) + 2, max(h_lines) + 2,
                    )

            # Group chars into lines, excluding table region
            lines_by_top: dict[float, list[dict[str, Any]]] = {}
            for char in page.chars:
                if table_bbox:
                    x0t, y0t, x1t, y1t = table_bbox
                    if (char["x0"] >= x0t and char["x1"] <= x1t
                            and char["top"] >= y0t and char["bottom"] <= y1t):
                        continue
                key = round(char["top"], 1)
                lines_by_top.setdefault(key, []).append(char)

            for top in sorted(lines_by_top.keys()):
                chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
                if not chars:
                    continue
                first = chars[0]
                font = first.get("fontname", "")
                is_bold = "Bold" in font or "700" in font
                x0 = first["x0"]

                # Bold text that's indented (not at left margin ~72)
                # and not heading-sized — these are bullet items
                if is_bold and x0 > 80 and round(first["size"], 1) < 15:
                    text = "".join(c["text"] for c in chars).strip()
                    if text and len(text) > 2:
                        bullet_texts.add(text)

    return bullet_texts


def _apply_bullets(text: str, bullet_items: set[str]) -> str:
    """Convert lines matching bullet items to markdown list items."""
    if not bullet_items:
        return text

    # Build a set of longer items for startswith matching
    long_items = {item for item in bullet_items if len(item) > 2}

    lines = text.split("\n")
    result = []
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or stripped.startswith("|"):
            result.append(line)
            continue
        # Exact match
        if stripped in bullet_items:
            result.append(f"- {stripped}")
            continue
        # Line starts with a bold/bullet item text
        # Exclude parameter value lines (contain '=')
        if "=" in stripped and any(c.isdigit() for c in stripped):
            result.append(line)
            continue
        matched = any(stripped.startswith(item) for item in long_items)
        if matched:
            result.append(f"- {stripped}")
        else:
            result.append(line)
    return "\n".join(result)


def _apply_headings(text: str, heading_map: dict[str, int]) -> str:
    """Apply heading markers to lines that match the heading map."""
    if not heading_map:
        return text

    lines = text.split("\n")
    result = []
    for line in lines:
        stripped = line.strip()
        if stripped in heading_map:
            level = heading_map[stripped]
            result.append(f"{'#' * level} {stripped}")
        else:
            result.append(line)
    return "\n".join(result)


def _convert_pdf_with_tables(src: Path) -> str:
    """Convert a PDF using MarkItDown for text and pdfplumber for tables.

    Hybrid approach:
    - MarkItDown provides good structure (headings, lists, formatting)
    - pdfplumber provides accurate table extraction
    - Table regions in MarkItDown output are replaced with pdfplumber tables
    """
    # 1. Get MarkItDown output (good for non-table text)
    md = MarkItDown()
    result = md.convert(str(src))
    markitdown_text = result.text_content

    # 2. Extract tables with pdfplumber
    tables = _extract_pdf_tables(src)
    if not tables:
        heading_map = _extract_heading_map(src)
        bullet_items = _extract_bullet_items(src)
        output = _apply_headings(_strip_page_headers(markitdown_text), heading_map)
        return _apply_bullets(output, bullet_items)

    # 3. Find table region in MarkItDown output
    cell_texts = _collect_table_cell_texts(tables)
    markitdown_lines = markitdown_text.split("\n")
    table_start, table_end = _find_table_region(markitdown_lines, cell_texts)

    if table_start is None or table_end is None:
        heading_map = _extract_heading_map(src)
        bullet_items = _extract_bullet_items(src)
        output = _apply_headings(_strip_page_headers(markitdown_text), heading_map)
        return _apply_bullets(output, bullet_items)

    # 4. Build output: before-table + tables + after-table
    before_table = "\n".join(markitdown_lines[:table_start]).strip()
    after_table = "\n".join(markitdown_lines[table_end + 1 :]).strip()

    table_mds = [_table_to_markdown(t) for t in tables if _table_to_markdown(t)]

    parts: list[str] = []
    if before_table:
        parts.append(before_table)
    parts.extend(table_mds)
    if after_table:
        parts.append(after_table)

    output = _strip_page_headers("\n\n".join(parts))

    # 5. Apply heading markers and bullet lists based on font analysis
    heading_map = _extract_heading_map(src)
    bullet_items = _extract_bullet_items(src)
    output = _apply_headings(output, heading_map)
    return _apply_bullets(output, bullet_items)


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
