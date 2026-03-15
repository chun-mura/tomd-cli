"""Core conversion logic.

Orchestrates PDF and Office document conversion to Markdown.
Individual format logic lives in pdf.py, office.py, and images.py.
"""

from __future__ import annotations

import logging
from pathlib import Path

from markitdown import MarkItDown

from tomd.pdf import convert_pdf
from tomd.office import (
    _correct_docx_headings,
    _restore_pptx_hyperlinks,
    _add_pptx_slide_separators,
    _correct_xlsx_merged_cells,
)
from tomd.images import (
    strip_base64_images,
    _extract_images,
    _replace_image_placeholders,
)

# Suppress noisy pdfminer warnings (e.g. "Could not get FontBBox")
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdffont").setLevel(logging.ERROR)

SUPPORTED_EXTENSIONS = {".pptx", ".docx", ".xlsx", ".pdf"}


def _convert_single(src: Path, dest: Path) -> str:
    """Convert a single file and return the Markdown content.

    This is the shared logic used by both convert_file and convert_dir.
    """
    if src.suffix.lower() == ".pdf":
        content = convert_pdf(src)
        content = strip_base64_images(content)
        return content

    # Office formats: use MarkItDown + format-specific post-processing
    md = MarkItDown()
    result = md.convert(str(src))
    content = result.text_content

    ext = src.suffix.lower()

    if ext == ".docx":
        content = _correct_docx_headings(src, content)
    elif ext == ".pptx":
        content = _restore_pptx_hyperlinks(src, content)
        content = _add_pptx_slide_separators(content)
    elif ext == ".xlsx":
        content = _correct_xlsx_merged_cells(src, content)

    # Strip base64 images, then replace placeholders with extracted files
    content = strip_base64_images(content)
    image_map = _extract_images(src, dest)
    content = _replace_image_placeholders(content, image_map)

    return content


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
    dest.parent.mkdir(parents=True, exist_ok=True)

    content = _convert_single(src, dest)
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
        content = _convert_single(f, dest)
        dest.write_text(content, encoding="utf-8")
        results.append(dest)

    return results
