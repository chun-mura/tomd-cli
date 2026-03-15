"""Core conversion logic."""

from __future__ import annotations

import logging
import re
from typing import Any
from pathlib import Path

import pdfplumber
from markitdown import MarkItDown

# Suppress noisy pdfminer warnings (e.g. "Could not get FontBBox")
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdffont").setLevel(logging.ERROR)

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


def _extract_page_header_texts(pdf_path: str | Path) -> set[str]:
    """Extract repeated page header/footer text from a PDF.

    Detects text that appears at the top or bottom margin of multiple pages
    at a smaller font size than body text (typical of headers/footers).
    Adjacent small-font lines on the same page are merged (handles headers
    split across ASCII/CJK fonts). Table regions are excluded.
    Returns the set of merged header/footer strings.
    """
    _, _, body_size = _analyze_pdf_layout(pdf_path)
    # Header/footer text is typically smaller than body text
    header_size_threshold = body_size * 0.85

    # Collect merged header/footer strings per page
    page_headers: list[list[str]] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            if not page.chars:
                page_headers.append([])
                continue

            # Use table-excluded line grouping
            lines_by_top = _group_lines_excluding_tables(page)
            page_height = float(page.height)

            # Collect lines in top/bottom margin at small font size
            margin_lines: list[tuple[float, str]] = []
            for top in sorted(lines_by_top.keys()):
                # Only consider top ~10% or bottom ~10% of the page
                in_margin = top < page_height * 0.1 or top > page_height * 0.9
                if not in_margin:
                    continue
                chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
                if not chars:
                    continue
                size = round(chars[0]["size"], 1)
                if size >= header_size_threshold:
                    continue
                text = "".join(c["text"] for c in chars).strip()
                if text:
                    margin_lines.append((top, text))

            # Merge adjacent margin lines into single strings
            # Skip standalone page numbers (pure digits)
            margin_lines = [
                (t, txt) for t, txt in margin_lines
                if not txt.strip().isdigit()
            ]
            merged: list[str] = []
            if margin_lines:
                current = margin_lines[0][1]
                prev_top = margin_lines[0][0]
                for top, text in margin_lines[1:]:
                    if abs(top - prev_top) < 5:
                        current += text
                    else:
                        if current:
                            merged.append(current)
                        current = text
                    prev_top = top
                if current:
                    merged.append(current)

            page_headers.append(merged)

    # Find strings that appear on 2+ pages
    string_counts: dict[str, int] = {}
    for page_strings in page_headers:
        for s in set(page_strings):  # deduplicate within page
            string_counts[s] = string_counts.get(s, 0) + 1

    num_pages = len(page_headers)
    min_occurrences = min(2, num_pages)
    return {
        text for text, count in string_counts.items()
        if count >= min_occurrences and len(text) > 1
    }


def _strip_page_headers(
    text: str,
    header_texts: set[str] | None = None,
) -> str:
    """Remove repeated page headers/footers and page numbers from text.

    Args:
        text: The text to clean.
        header_texts: Optional set of known header/footer text fragments
            extracted from PDF font analysis. If not provided, falls back
            to heuristic detection of repeated lines.
    """
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

    # Combine with PDF-detected header texts if provided
    all_header_texts: set[str] = set(header_texts) if header_texts else set()

    # Also detect from text: lines appearing 3+ times
    for line, count in line_counts.items():
        if count >= 3 and len(line) > 2:
            all_header_texts.add(line)

    # Build combined header string for prefix matching
    # (used to detect concatenated "Header2Content" patterns)
    # Sort by length descending to match longest first
    sorted_headers = sorted(all_header_texts, key=len, reverse=True)

    # Build page-number pattern
    if sorted_headers:
        header_alts = "|".join(re.escape(h) for h in sorted_headers)
        page_num_pattern = re.compile(
            rf'^({header_alts})?\s*\d+\s*$'
        )
    else:
        page_num_pattern = re.compile(r'^\s*\d+\s*$')

    filtered: list[str] = []
    for line in lines:
        stripped = line.strip()
        if stripped in repeated:
            continue
        if page_num_pattern.match(stripped):
            continue
        # Clean lines like "HeaderText2remaining" -> "remaining"
        # where a page header is concatenated with page number and content
        cleaned = False
        for header in sorted_headers:
            if stripped.startswith(header) and stripped != header:
                after = stripped[len(header):]
                m = re.match(r'^\d+(.*)', after)
                if m:
                    rest = m.group(1).strip()
                    if rest:
                        filtered.append(rest)
                    cleaned = True
                    break
        if not cleaned:
            filtered.append(line)

    return "\n".join(filtered)


def _extract_heading_map(pdf_path: str | Path) -> dict[str, int]:
    """Build a mapping from text to heading level based on font size.

    Analyzes font sizes across all pages and assigns heading levels:
    - Largest font size → # (h1)
    - Second largest → ## (h2)
    - Third largest (if clearly distinct from body) → ### (h3)
    - Everything else → body text (no heading)

    Adjacent lines at the same heading size are merged into one heading
    (handles cases where a heading spans multiple PDF lines, e.g. different
    fonts for ASCII and CJK portions).
    """
    # Collect all font sizes and their associated text lines
    size_to_lines: dict[float, list[str]] = {}
    # Also collect (page, top, size, text) for adjacency merging
    all_heading_lines: list[tuple[int, float, float, str]] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf.pages):
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
                    all_heading_lines.append(
                        (page_idx, top, dominant_size, text)
                    )

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

    heading_size_set = set(heading_sizes[:2])

    # Merge adjacent lines at the same heading size on the same page
    # (e.g. "AI" at top=70.7 + "モデル精度チューニング観点" at top=78.0)
    merged_headings: list[tuple[float, str]] = []
    prev_page = -1
    prev_top = -999.0
    prev_size = -1.0
    prev_text = ""
    for page_idx, top, size, text in all_heading_lines:
        if size not in heading_size_set:
            if prev_text:
                merged_headings.append((prev_size, prev_text))
                prev_text = ""
                prev_page = -1
            continue
        # Adjacent if same page and within 15 points vertically
        if (page_idx == prev_page and size == prev_size
                and abs(top - prev_top) < 15):
            prev_text += text
            prev_top = top
        else:
            if prev_text:
                merged_headings.append((prev_size, prev_text))
            prev_page = page_idx
            prev_top = top
            prev_size = size
            prev_text = text
    if prev_text:
        merged_headings.append((prev_size, prev_text))

    # Assign heading levels
    heading_map: dict[str, int] = {}
    for size, text in merged_headings:
        level = heading_sizes.index(size) + 1 if size in heading_sizes else 1
        heading_map[text] = level

    return heading_map


def _normalize_for_match(text: str) -> str:
    """Normalize text for fuzzy matching between PDF extraction and MarkItDown."""
    import unicodedata
    result = unicodedata.normalize("NFKC", text)
    # Remove all whitespace
    result = re.sub(r'\s+', '', result)
    # Remove parens and slashes (both half/fullwidth) for matching
    result = result.replace('（', '').replace('）', '')
    result = result.replace('(', '').replace(')', '')
    result = result.replace('/', '').replace('∕', '')
    return result


def _group_lines_excluding_tables(
    page: Any,
) -> dict[float, list[dict[str, Any]]]:
    """Group page chars into lines by top position, excluding table regions."""
    table_bbox: tuple[float, ...] | None = None
    if page.rects:
        v_lines, h_lines = _detect_table_lines(page.rects)
        if len(v_lines) >= 2 and len(h_lines) >= 2:
            table_bbox = (
                min(v_lines) - 2, min(h_lines) - 2,
                max(v_lines) + 2, max(h_lines) + 2,
            )

    lines_by_top: dict[float, list[dict[str, Any]]] = {}
    for char in page.chars:
        if table_bbox:
            x0t, y0t, x1t, y1t = table_bbox
            if (char["x0"] >= x0t and char["x1"] <= x1t
                    and char["top"] >= y0t and char["bottom"] <= y1t):
                continue
        key = round(char["top"], 1)
        lines_by_top.setdefault(key, []).append(char)

    return lines_by_top


def _is_monospace_font(fontname: str) -> bool:
    """Check if a font is monospace based on common naming conventions."""
    mono_indicators = (
        "Mono", "Courier", "Consolas", "Menlo", "DejaVuSans",
        "LiberationMono", "SourceCodePro", "FiraCode", "Inconsolata",
    )
    return any(ind.lower() in fontname.lower() for ind in mono_indicators)


def _is_bold_font(fontname: str) -> bool:
    """Check if a font is bold based on common naming conventions."""
    bold_indicators = ("Bold", "700", "Bld", "Heavy", "Black")
    return any(ind in fontname for ind in bold_indicators)


def _analyze_pdf_layout(
    pdf_path: str | Path,
) -> tuple[float, float, float]:
    """Analyze PDF layout to determine left margin, body indent, and body size.

    Returns (left_margin, body_indent, body_size) where:
    - left_margin: the leftmost x0 of non-table body text
    - body_indent: the most common x0 position of body text lines
    - body_size: the most common font size (body text size)
    """
    x0_counts: dict[float, int] = {}
    size_char_counts: dict[float, int] = {}

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            if not page.chars:
                continue
            lines_by_top = _group_lines_excluding_tables(page)
            for top in sorted(lines_by_top.keys()):
                chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
                if not chars:
                    continue
                first = chars[0]
                font = first.get("fontname", "")
                # Skip monospace and page header/footer sized text
                if _is_monospace_font(font):
                    continue
                x0 = round(first["x0"], 0)
                size = round(first["size"], 1)
                text = "".join(c["text"] for c in chars).strip()
                if text and len(text) > 1:
                    x0_counts[x0] = x0_counts.get(x0, 0) + len(text)
                    size_char_counts[size] = (
                        size_char_counts.get(size, 0) + len(text)
                    )

    if not x0_counts or not size_char_counts:
        return 72.0, 90.0, 10.0

    body_indent = float(max(x0_counts, key=x0_counts.get))  # type: ignore[arg-type]
    left_margin = float(min(x0_counts.keys()))
    body_size = float(max(size_char_counts, key=size_char_counts.get))  # type: ignore[arg-type]
    return left_margin, body_indent, body_size


def _extract_bullet_items(pdf_path: str | Path) -> set[str]:
    """Detect lines that should be bullet list items based on bold font + indent.

    Lines starting with a bold font and indented from the left margin
    (outside table regions) are treated as bullet items.
    The indent threshold and heading size are determined dynamically
    from the PDF layout rather than hardcoded.
    """
    left_margin, body_indent, body_size = _analyze_pdf_layout(pdf_path)
    # Bullet items are bold text indented beyond the left margin
    # but not heading-sized (headings are > 1.3x body size)
    heading_threshold = body_size * 1.3
    # Indent threshold: midpoint between left margin and body indent
    indent_threshold = (left_margin + body_indent) / 2

    bullet_texts: set[str] = set()

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            if not page.chars:
                continue

            lines_by_top = _group_lines_excluding_tables(page)

            for top in sorted(lines_by_top.keys()):
                chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
                if not chars:
                    continue
                first = chars[0]
                font = first.get("fontname", "")
                is_bold = _is_bold_font(font)
                x0 = first["x0"]
                size = round(first["size"], 1)

                # Bold text that's indented (not at left margin)
                # and not heading-sized — these are bullet items
                if is_bold and x0 > indent_threshold and size < heading_threshold:
                    text = "".join(c["text"] for c in chars).strip()
                    if text and len(text) > 2:
                        bullet_texts.add(text)

    return bullet_texts


def _extract_sub_items(pdf_path: str | Path) -> set[str]:
    """Detect lines that should be sub-list items based on indent depth.

    Lines that are indented deeper than the primary bullet indent level,
    use a non-bold / non-monospace / non-CJK-regular font (typically italic
    or medium weight Latin font), and are within the body text size range
    are treated as sub-list items.
    """
    _, body_indent, body_size = _analyze_pdf_layout(pdf_path)
    heading_threshold = body_size * 1.3
    # Sub-items are indented deeper than bullet items
    # Typically ~1.2x the body indent
    sub_indent_min = body_indent * 1.1
    sub_indent_max = body_indent * 1.4

    sub_texts: set[str] = set()

    # Collect all font names used at the sub-indent level
    # to distinguish label fonts from regular body text fonts
    fonts_at_sub_indent: dict[str, int] = {}
    fonts_at_body_indent: dict[str, int] = {}

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            if not page.chars:
                continue
            lines_by_top = _group_lines_excluding_tables(page)
            for top in sorted(lines_by_top.keys()):
                chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
                if not chars:
                    continue
                first = chars[0]
                x0 = first["x0"]
                font = first.get("fontname", "")
                text = "".join(c["text"] for c in chars).strip()
                if not text or len(text) <= 2:
                    continue
                if sub_indent_min <= x0 <= sub_indent_max:
                    fonts_at_sub_indent[font] = (
                        fonts_at_sub_indent.get(font, 0) + 1
                    )
                elif abs(x0 - body_indent) < 5:
                    fonts_at_body_indent[font] = (
                        fonts_at_body_indent.get(font, 0) + 1
                    )

    if not fonts_at_sub_indent:
        return sub_texts

    # The most common font at body indent is the regular body font.
    # Sub-item label fonts are fonts at sub-indent that are NOT:
    # - the regular body font
    # - bold fonts
    # - monospace fonts
    body_font: str = (
        str(max(fonts_at_body_indent, key=fonts_at_body_indent.get))  # type: ignore[arg-type]
        if fonts_at_body_indent else ""
    )
    label_fonts: set[str] = set()
    for font in fonts_at_sub_indent:
        if (font != body_font
                and not _is_bold_font(font)
                and not _is_monospace_font(font)):
            label_fonts.add(font)

    if not label_fonts:
        return sub_texts

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            if not page.chars:
                continue
            lines_by_top = _group_lines_excluding_tables(page)
            for top in sorted(lines_by_top.keys()):
                chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
                if not chars:
                    continue
                first = chars[0]
                x0 = first["x0"]
                size = round(first["size"], 1)
                font = first.get("fontname", "")

                if (font in label_fonts
                        and sub_indent_min <= x0 <= sub_indent_max
                        and size < heading_threshold):
                    text = "".join(c["text"] for c in chars).strip()
                    if text and len(text) > 2:
                        sub_texts.add(text)

    return sub_texts


def _extract_inline_code_map(pdf_path: str | Path) -> dict[str, list[str]]:
    """Extract inline code spans (monospace font) from the PDF.

    Returns a dict with key "__code_only__" containing a list of code spans.
    Monospace text on the same PDF line is split into separate spans when
    there are significant x-position gaps (indicating separate code fragments
    embedded within regular text).
    """
    code_spans: set[str] = set()

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            if not page.chars:
                continue

            lines_by_top = _group_lines_excluding_tables(page)

            for top in sorted(lines_by_top.keys()):
                chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
                if not chars:
                    continue

                fonts = {c.get("fontname", "") for c in chars}
                all_mono = all(_is_monospace_font(f) for f in fonts)

                if not all_mono:
                    continue

                # Split into separate spans based on x-position gaps
                # (gaps > char_width indicate separate code fragments)
                spans: list[str] = []
                current = chars[0]["text"]
                for i in range(1, len(chars)):
                    gap = chars[i]["x0"] - chars[i - 1]["x1"]
                    char_width = chars[i - 1]["x1"] - chars[i - 1]["x0"]
                    # Gap larger than ~2x char width means separate span
                    if gap > char_width * 2:
                        span = current.strip()
                        if span:
                            spans.append(span)
                        current = chars[i]["text"]
                    else:
                        current += chars[i]["text"]
                span = current.strip()
                if span:
                    spans.append(span)

                for s in spans:
                    # Filter out: single chars, continuation fragments
                    # (starting with comma), and incomplete assignments
                    # (ending with '=')
                    if (s and len(s) > 1
                            and not s.startswith(",")
                            and not s.endswith("=")):
                        code_spans.add(s)

    return {"__code_only__": sorted(code_spans)}


def _apply_bullets(text: str, bullet_items: set[str]) -> str:
    """Convert lines matching bullet items to markdown list items."""
    if not bullet_items:
        return text

    # Build normalized versions for fuzzy matching
    norm_to_original = {}
    for item in bullet_items:
        norm = _normalize_for_match(item)
        if len(norm) > 2:
            norm_to_original[norm] = item

    lines = text.split("\n")
    result = []
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or stripped.startswith("|"):
            result.append(line)
            continue
        # Already a bullet
        if stripped.startswith("- "):
            result.append(line)
            continue
        # Exclude parameter value lines (contain '=')
        if "=" in stripped and any(c.isdigit() for c in stripped):
            result.append(line)
            continue

        norm_stripped = _normalize_for_match(stripped)

        # Exact match (original or normalized)
        if stripped in bullet_items:
            result.append(f"- {stripped}")
            continue
        if norm_stripped in norm_to_original:
            result.append(f"- {stripped}")
            continue

        # Normalized startswith match
        matched = any(
            norm_stripped.startswith(n) for n in norm_to_original if len(n) > 3
        )
        if matched:
            result.append(f"- {stripped}")
        else:
            result.append(line)
    return "\n".join(result)


def _apply_sub_items(text: str, sub_items: set[str]) -> str:
    """Convert lines matching sub-items to indented markdown list items."""
    if not sub_items:
        return text

    norm_subs = {}
    for item in sub_items:
        norm = _normalize_for_match(item)
        if len(norm) > 3:
            norm_subs[norm] = item

    lines = text.split("\n")
    result = []
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or stripped.startswith("|"):
            result.append(line)
            continue
        if stripped.startswith("- ") or stripped.startswith("  - "):
            result.append(line)
            continue
        # Exclude parameter value lines
        if "=" in stripped and any(c.isdigit() for c in stripped):
            result.append(line)
            continue

        norm_stripped = _normalize_for_match(stripped)

        if stripped in sub_items or norm_stripped in norm_subs:
            result.append(f"  - {stripped}")
            continue

        # Normalized startswith match for sub-items
        matched = any(
            norm_stripped.startswith(n) for n in norm_subs if len(n) > 4
        )
        if matched:
            result.append(f"  - {stripped}")
        else:
            result.append(line)
    return "\n".join(result)


def _apply_inline_code(text: str, code_map: dict[str, list[str]]) -> str:
    """Wrap inline code spans in backticks based on PDF monospace font detection."""
    code_spans = code_map.get("__code_only__", [])
    if not code_spans:
        return text

    all_spans = {s for s in code_spans if len(s) > 1}

    lines = text.split("\n")
    result = []
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or stripped.startswith("|"):
            result.append(line)
            continue
        # Skip bullet/sub-item lines
        if stripped.startswith("- ") or stripped.startswith("  - "):
            result.append(line)
            continue

        # Check if the entire line matches a code span → wrap whole line
        if stripped in all_spans:
            leading = line[: len(line) - len(line.lstrip())]
            result.append(leading + f"`{stripped}`")
            continue

        # Apply code spans within the line (longest first to avoid
        # partial matches), skipping spans already inside backticks
        modified = stripped
        for code_span in sorted(all_spans, key=len, reverse=True):
            if code_span not in modified:
                continue
            if f"`{code_span}`" in modified:
                continue
            # Skip if the match position is already inside backticks
            idx = modified.find(code_span)
            # Check for surrounding backticks
            if idx > 0 and modified[idx - 1] == '`':
                continue
            end = idx + len(code_span)
            if end < len(modified) and modified[end] == '`':
                continue
            modified = modified.replace(code_span, f"`{code_span}`")

        if modified != stripped:
            leading = line[: len(line) - len(line.lstrip())]
            result.append(leading + modified)
        else:
            result.append(line)
    return "\n".join(result)


def _apply_headings(text: str, heading_map: dict[str, int]) -> str:
    """Apply heading markers to lines that match the heading map."""
    if not heading_map:
        return text

    # Build normalized lookup
    norm_headings: dict[str, tuple[str, int]] = {}
    for h_text, level in heading_map.items():
        norm = _normalize_for_match(h_text)
        norm_headings[norm] = (h_text, level)

    lines = text.split("\n")
    result = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            result.append(line)
            continue
        # Skip lines already marked as headings
        if stripped.startswith("#"):
            result.append(line)
            continue
        if stripped in heading_map:
            level = heading_map[stripped]
            result.append(f"{'#' * level} {stripped}")
        else:
            norm = _normalize_for_match(stripped)
            if norm in norm_headings:
                _, level = norm_headings[norm]
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

    # 2. Extract page header/footer texts for stripping
    header_texts = _extract_page_header_texts(src)

    # 3. Extract tables with pdfplumber
    tables = _extract_pdf_tables(src)
    if not tables:
        output = _strip_page_headers(markitdown_text, header_texts)
        return _apply_pdf_formatting(output, src)

    # 4. Find table region in MarkItDown output
    cell_texts = _collect_table_cell_texts(tables)
    markitdown_lines = markitdown_text.split("\n")
    table_start, table_end = _find_table_region(markitdown_lines, cell_texts)

    if table_start is None or table_end is None:
        output = _strip_page_headers(markitdown_text, header_texts)
        return _apply_pdf_formatting(output, src)

    # 5. Build output: before-table + tables + after-table
    before_table = "\n".join(markitdown_lines[:table_start]).strip()
    after_table = "\n".join(markitdown_lines[table_end + 1 :]).strip()

    table_mds = [_table_to_markdown(t) for t in tables if _table_to_markdown(t)]

    parts: list[str] = []
    if before_table:
        parts.append(before_table)
    parts.extend(table_mds)
    if after_table:
        parts.append(after_table)

    output = _strip_page_headers("\n\n".join(parts), header_texts)

    # 6. Apply all formatting based on font analysis
    return _apply_pdf_formatting(output, src)


def _apply_pdf_formatting(text: str, pdf_path: str | Path) -> str:
    """Apply all PDF-derived formatting: headings, bullets, sub-items, inline code."""
    heading_map = _extract_heading_map(pdf_path)
    bullet_items = _extract_bullet_items(pdf_path)
    sub_items = _extract_sub_items(pdf_path)
    code_map = _extract_inline_code_map(pdf_path)

    output = _apply_headings(text, heading_map)
    output = _apply_bullets(output, bullet_items)
    output = _apply_sub_items(output, sub_items)
    output = _apply_inline_code(output, code_map)
    return output


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
