"""PDF conversion logic with pdfplumber + MarkItDown hybrid approach."""

from __future__ import annotations

import re
import unicodedata
from typing import Any
from pathlib import Path

import pdfplumber
from markitdown import MarkItDown


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
        elif h < 2 and w > 10:
            h_lines.add(round((r["top"] + r["bottom"]) / 2, 1))
        # Cell-sized rects: use top/bottom edges
        elif w > 10 and h > 5:
            h_lines.add(round(r["top"], 1))
            h_lines.add(round(r["bottom"], 1))

    return sorted(v_lines), sorted(h_lines)


def _table_to_markdown(table: list[list[str | None]]) -> str:
    """Convert a 2D table to a Markdown table string."""
    if not table or len(table) < 2:
        return ""

    # Clean cells
    cleaned = []
    for row in table:
        cells = [_clean_cell(c) for c in row]
        if any(c for c in cells):
            cleaned.append(cells)

    if len(cleaned) < 2:
        return ""

    num_cols = max(len(row) for row in cleaned)
    # Pad short rows
    for row in cleaned:
        while len(row) < num_cols:
            row.append("")

    header = cleaned[0]
    separator = ["---"] * num_cols
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(separator) + " |",
    ]
    for row in cleaned[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def _tables_have_same_header(
    a: list[list[str | None]],
    b: list[list[str | None]],
) -> bool:
    """Check if two tables have the same header row (for cross-page merging)."""
    if not a or not b:
        return False
    ha = [_clean_cell(c) for c in a[0]]
    hb = [_clean_cell(c) for c in b[0]]
    if ha != hb:
        return False
    # Single-column tables are not merged (too ambiguous)
    return len(ha) > 1


def _merge_continuation_row(
    last_row: list[str | None],
    cont_row: list[str | None],
) -> list[str | None]:
    """Merge a continuation row (with empty leading cells) into the previous row."""
    merged: list[str | None] = list(last_row)
    for i, cell in enumerate(cont_row):
        if cell and cell.strip():
            existing = merged[i] if i < len(merged) else None
            if existing and existing.strip():
                merged[i] = existing + " " + cell
            elif i < len(merged):
                merged[i] = cell
            else:
                merged.append(cell)
    return merged


def _normalize_for_match(text: str) -> str:
    """Normalize text for fuzzy matching between PDF extraction and MarkItDown."""
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


# ---------------------------------------------------------------------------
# PDF analysis helpers that work on pre-opened page data
# ---------------------------------------------------------------------------

def _analyze_layout_from_pages(
    pages_data: list[dict[str, Any]],
) -> tuple[float, float, float]:
    """Analyze PDF layout from pre-extracted page data.

    Returns (left_margin, body_indent, body_size).
    """
    x0_counts: dict[float, int] = {}
    size_char_counts: dict[float, int] = {}

    for pdata in pages_data:
        for top in sorted(pdata["lines_excl_tables"].keys()):
            chars = sorted(pdata["lines_excl_tables"][top], key=lambda c: c["x0"])
            if not chars:
                continue
            first = chars[0]
            font = first.get("fontname", "")
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


def _extract_page_header_texts_from_pages(
    pages_data: list[dict[str, Any]],
    body_size: float,
) -> set[str]:
    """Extract repeated page header/footer text from pre-extracted page data."""
    header_size_threshold = body_size * 0.85

    page_headers: list[list[str]] = []

    for pdata in pages_data:
        lines_by_top = pdata["lines_excl_tables"]
        page_height = pdata["height"]

        if not lines_by_top:
            page_headers.append([])
            continue

        margin_lines: list[tuple[float, str]] = []
        for top in sorted(lines_by_top.keys()):
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

        # Merge adjacent margin lines, skip standalone page numbers
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
        for s in set(page_strings):
            string_counts[s] = string_counts.get(s, 0) + 1

    num_pages = len(page_headers)
    min_occurrences = min(2, num_pages)
    return {
        text for text, count in string_counts.items()
        if count >= min_occurrences and len(text) > 1
    }


def _extract_heading_map_from_pages(
    pages_data: list[dict[str, Any]],
    body_size: float,
) -> dict[str, int]:
    """Build heading map from pre-extracted page data."""
    size_to_lines: dict[float, list[str]] = {}
    all_heading_lines: list[tuple[int, float, float, str]] = []

    for page_idx, pdata in enumerate(pages_data):
        lines_by_top = pdata["all_lines"]
        for top in sorted(lines_by_top.keys()):
            chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
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

    # Only sizes significantly larger than body text are headings
    heading_sizes = sorted(
        [s for s in size_to_lines if s > body_size * 1.3],
        reverse=True,
    )

    if not heading_sizes:
        return {}

    heading_size_set = set(heading_sizes[:2])

    # Merge adjacent lines at the same heading size on the same page
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

    heading_map: dict[str, int] = {}
    for size, text in merged_headings:
        level = heading_sizes.index(size) + 1 if size in heading_sizes else 1
        heading_map[text] = level

    return heading_map


def _extract_bullet_items_from_pages(
    pages_data: list[dict[str, Any]],
    left_margin: float,
    body_indent: float,
    body_size: float,
) -> set[str]:
    """Detect bullet list items from pre-extracted page data."""
    heading_threshold = body_size * 1.3
    indent_threshold = (left_margin + body_indent) / 2

    bullet_texts: set[str] = set()

    for pdata in pages_data:
        for top in sorted(pdata["lines_excl_tables"].keys()):
            chars = sorted(pdata["lines_excl_tables"][top], key=lambda c: c["x0"])
            if not chars:
                continue
            first = chars[0]
            font = first.get("fontname", "")
            is_bold = _is_bold_font(font)
            x0 = first["x0"]
            size = round(first["size"], 1)

            if is_bold and x0 > indent_threshold and size < heading_threshold:
                text = "".join(c["text"] for c in chars).strip()
                if text and len(text) > 2:
                    bullet_texts.add(text)

    return bullet_texts


def _extract_sub_items_from_pages(
    pages_data: list[dict[str, Any]],
    body_indent: float,
    body_size: float,
) -> set[str]:
    """Detect sub-list items from pre-extracted page data."""
    heading_threshold = body_size * 1.3
    sub_indent_min = body_indent * 1.1
    sub_indent_max = body_indent * 1.4

    # Collect font info at sub-indent and body-indent levels
    fonts_at_sub_indent: dict[str, int] = {}
    fonts_at_body_indent: dict[str, int] = {}

    for pdata in pages_data:
        for top in sorted(pdata["lines_excl_tables"].keys()):
            chars = sorted(pdata["lines_excl_tables"][top], key=lambda c: c["x0"])
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
        return set()

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
        return set()

    sub_texts: set[str] = set()
    for pdata in pages_data:
        for top in sorted(pdata["lines_excl_tables"].keys()):
            chars = sorted(pdata["lines_excl_tables"][top], key=lambda c: c["x0"])
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


def _extract_inline_code_map_from_pages(
    pages_data: list[dict[str, Any]],
) -> dict[str, list[str]]:
    """Extract inline code spans from pre-extracted page data."""
    code_spans: set[str] = set()

    for pdata in pages_data:
        for top in sorted(pdata["lines_excl_tables"].keys()):
            chars = sorted(pdata["lines_excl_tables"][top], key=lambda c: c["x0"])
            if not chars:
                continue

            fonts = {c.get("fontname", "") for c in chars}
            all_mono = all(_is_monospace_font(f) for f in fonts)

            if not all_mono:
                continue

            # Split into separate spans based on x-position gaps
            spans: list[str] = []
            current = chars[0]["text"]
            for i in range(1, len(chars)):
                gap = chars[i]["x0"] - chars[i - 1]["x1"]
                char_width = chars[i - 1]["x1"] - chars[i - 1]["x0"]
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
                if (s and len(s) > 1
                        and not s.startswith(",")
                        and not s.endswith("=")):
                    code_spans.add(s)

    return {"__code_only__": sorted(code_spans)}


# ---------------------------------------------------------------------------
# Text transformation helpers (apply extracted info to markdown text)
# ---------------------------------------------------------------------------

def _apply_bullets(text: str, bullet_items: set[str]) -> str:
    """Convert lines matching bullet items to markdown list items."""
    if not bullet_items:
        return text

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
        if stripped.startswith("- "):
            result.append(line)
            continue
        if "=" in stripped and any(c.isdigit() for c in stripped):
            result.append(line)
            continue

        norm_stripped = _normalize_for_match(stripped)

        if stripped in bullet_items:
            result.append(f"- {stripped}")
            continue
        if norm_stripped in norm_to_original:
            result.append(f"- {stripped}")
            continue

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
        if "=" in stripped and any(c.isdigit() for c in stripped):
            result.append(line)
            continue

        norm_stripped = _normalize_for_match(stripped)

        if stripped in sub_items or norm_stripped in norm_subs:
            result.append(f"  - {stripped}")
            continue

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
        if stripped.startswith("- ") or stripped.startswith("  - "):
            result.append(line)
            continue

        if stripped in all_spans:
            leading = line[: len(line) - len(line.lstrip())]
            result.append(leading + f"`{stripped}`")
            continue

        modified = stripped
        for code_span in sorted(all_spans, key=len, reverse=True):
            if code_span not in modified:
                continue
            if f"`{code_span}`" in modified:
                continue
            idx = modified.find(code_span)
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


def _strip_page_headers(
    text: str,
    header_texts: set[str] | None = None,
) -> str:
    """Remove repeated page headers/footers and page numbers from text."""
    lines = text.split("\n")
    if len(lines) < 5:
        return text

    line_counts: dict[str, int] = {}
    for line in lines:
        stripped = line.strip()
        if stripped:
            line_counts[stripped] = line_counts.get(stripped, 0) + 1

    repeated = {line for line, count in line_counts.items() if count >= 2}

    all_header_texts: set[str] = set(header_texts) if header_texts else set()

    for line, count in line_counts.items():
        if count >= 3 and len(line) > 2:
            all_header_texts.add(line)

    sorted_headers = sorted(all_header_texts, key=len, reverse=True)

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


# ---------------------------------------------------------------------------
# Table extraction from pdfplumber
# ---------------------------------------------------------------------------

def _extract_pdf_tables(pdf_path: str | Path) -> list[list[list[str | None]]]:
    """Extract tables from PDF using pdfplumber with rectangle-based line detection."""
    all_tables: list[list[list[str | None]]] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            if not page.rects:
                continue

            v_lines, h_lines = _detect_table_lines(page.rects)
            if len(v_lines) < 2 or len(h_lines) < 2:
                continue

            table = page.extract_table(
                table_settings={
                    "vertical_strategy": "explicit",
                    "horizontal_strategy": "explicit",
                    "explicit_vertical_lines": v_lines,
                    "explicit_horizontal_lines": h_lines,
                }
            )
            if not table:
                continue

            # Merge continuation tables (same header across pages)
            if (all_tables
                    and _tables_have_same_header(all_tables[-1], table)):
                for row in table[1:]:
                    if any(c for c in row if c):
                        last_row = all_tables[-1][-1]
                        first_cell = row[0] if row else None
                        if first_cell is None or first_cell.strip() == "":
                            all_tables[-1][-1] = _merge_continuation_row(
                                last_row, row,
                            )
                        else:
                            all_tables[-1].append(row)
            else:
                all_tables.append(table)

    return all_tables


def _collect_table_cell_texts(
    tables: list[list[list[str | None]]],
) -> set[str]:
    """Collect all cell text fragments for matching against MarkItDown output."""
    texts: set[str] = set()
    for table in tables:
        for row in table:
            for cell in row:
                if cell is None:
                    continue
                cleaned = cell.replace("\x00", "-")
                for line in cleaned.split("\n"):
                    line = line.strip()
                    if line:
                        texts.add(line)
    return texts


def _find_table_region(
    lines: list[str],
    cell_texts: set[str],
) -> tuple[int | None, int | None]:
    """Find the start and end line indices of the table region in MarkItDown output."""
    long_texts = {t for t in cell_texts if len(t) > 5}

    start: int | None = None
    end: int | None = None

    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        is_table = stripped in cell_texts
        if not is_table:
            is_table = any(ct in stripped for ct in long_texts)
        if is_table:
            if start is None:
                start = i
            end = i

    return start, end


# ---------------------------------------------------------------------------
# Pre-extract page data (open PDF once)
# ---------------------------------------------------------------------------

def _extract_pages_data(pdf_path: str | Path) -> list[dict[str, Any]]:
    """Open the PDF once and extract all per-page data needed for analysis.

    Each page dict contains:
    - "height": page height
    - "chars": raw char list
    - "all_lines": dict[float, list[char]] grouped by top (all chars)
    - "lines_excl_tables": dict[float, list[char]] grouped by top (excluding table regions)
    """
    pages_data: list[dict[str, Any]] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            all_lines: dict[float, list[dict[str, Any]]] = {}
            if page.chars:
                for char in page.chars:
                    key = round(char["top"], 1)
                    all_lines.setdefault(key, []).append(char)

            lines_excl = _group_lines_excluding_tables(page)

            pages_data.append({
                "height": float(page.height),
                "chars": page.chars if page.chars else [],
                "all_lines": all_lines,
                "lines_excl_tables": lines_excl,
            })

    return pages_data


# ---------------------------------------------------------------------------
# High-level PDF conversion
# ---------------------------------------------------------------------------

def _apply_pdf_formatting(text: str, pdf_path: str | Path) -> str:
    """Apply all PDF-derived formatting: headings, bullets, sub-items, inline code.

    Opens the PDF once and reuses extracted page data for all analysis.
    """
    pages_data = _extract_pages_data(pdf_path)
    left_margin, body_indent, body_size = _analyze_layout_from_pages(pages_data)

    heading_map = _extract_heading_map_from_pages(pages_data, body_size)
    bullet_items = _extract_bullet_items_from_pages(
        pages_data, left_margin, body_indent, body_size,
    )
    sub_items = _extract_sub_items_from_pages(
        pages_data, body_indent, body_size,
    )
    code_map = _extract_inline_code_map_from_pages(pages_data)

    output = _apply_headings(text, heading_map)
    output = _apply_bullets(output, bullet_items)
    output = _apply_sub_items(output, sub_items)
    output = _apply_inline_code(output, code_map)
    return output


def convert_pdf(src: Path) -> str:
    """Convert a PDF using MarkItDown for text and pdfplumber for tables.

    Hybrid approach:
    - MarkItDown provides good structure (headings, lists, formatting)
    - pdfplumber provides accurate table extraction
    - Table regions in MarkItDown output are replaced with pdfplumber tables
    """
    # 1. Get MarkItDown output
    md = MarkItDown()
    result = md.convert(str(src))
    markitdown_text = result.text_content

    # 2. Extract page header/footer texts for stripping
    # (opens PDF once for header detection)
    pages_data = _extract_pages_data(src)
    _, _, body_size = _analyze_layout_from_pages(pages_data)
    header_texts = _extract_page_header_texts_from_pages(pages_data, body_size)

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
    after_table = "\n".join(markitdown_lines[table_end + 1:]).strip()

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


# ---------------------------------------------------------------------------
# Legacy compatibility aliases (used by the old single-file converter.py)
# These map old function names to the new ones for backward compatibility
# in tests that import from tomd.converter.
# ---------------------------------------------------------------------------

# The old functions that opened the PDF themselves:
def _analyze_pdf_layout(pdf_path: str | Path) -> tuple[float, float, float]:
    """Legacy wrapper: analyze PDF layout."""
    pages_data = _extract_pages_data(pdf_path)
    return _analyze_layout_from_pages(pages_data)


def _extract_page_header_texts(pdf_path: str | Path) -> set[str]:
    """Legacy wrapper: extract page header texts."""
    pages_data = _extract_pages_data(pdf_path)
    _, _, body_size = _analyze_layout_from_pages(pages_data)
    return _extract_page_header_texts_from_pages(pages_data, body_size)


def _extract_heading_map(pdf_path: str | Path) -> dict[str, int]:
    """Legacy wrapper: extract heading map."""
    pages_data = _extract_pages_data(pdf_path)
    _, _, body_size = _analyze_layout_from_pages(pages_data)
    return _extract_heading_map_from_pages(pages_data, body_size)


def _extract_bullet_items(pdf_path: str | Path) -> set[str]:
    """Legacy wrapper: extract bullet items."""
    pages_data = _extract_pages_data(pdf_path)
    left_margin, body_indent, body_size = _analyze_layout_from_pages(pages_data)
    return _extract_bullet_items_from_pages(
        pages_data, left_margin, body_indent, body_size,
    )


def _extract_sub_items(pdf_path: str | Path) -> set[str]:
    """Legacy wrapper: extract sub items."""
    pages_data = _extract_pages_data(pdf_path)
    _, body_indent, body_size = _analyze_layout_from_pages(pages_data)
    return _extract_sub_items_from_pages(pages_data, body_indent, body_size)


def _extract_inline_code_map(pdf_path: str | Path) -> dict[str, list[str]]:
    """Legacy wrapper: extract inline code map."""
    pages_data = _extract_pages_data(pdf_path)
    return _extract_inline_code_map_from_pages(pages_data)


# Alias for backward compat
_convert_pdf_with_tables = convert_pdf
