"""Core conversion logic."""

from __future__ import annotations

import re
from pathlib import Path

from markitdown import MarkItDown

SUPPORTED_EXTENSIONS = {".pptx", ".docx", ".xlsx", ".pdf"}


def strip_base64_images(text: str) -> str:
    """Replace embedded base64 image data with a placeholder."""
    return re.sub(
        r'!\[([^\]]*)\]\(data:image/[^)]+\)',
        lambda m: f'![{m.group(1)}]()',
        text,
    )


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

    md = MarkItDown()
    result = md.convert(str(src))
    content = strip_base64_images(result.text_content)

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

    md = MarkItDown()
    results: list[Path] = []
    for f in files:
        dest = dest_dir / f.with_suffix(".md").name
        result = md.convert(str(f))
        dest.write_text(strip_base64_images(result.text_content), encoding="utf-8")
        results.append(dest)

    return results
