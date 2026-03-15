"""Image extraction and base64 stripping for document-to-Markdown conversion."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path


def strip_base64_images(text: str) -> str:
    """Replace embedded base64 image data with a placeholder."""
    return re.sub(
        r'!\[([^\]]*)\]\(data:image/[^)]+\)',
        lambda m: f'![{m.group(1)}]()',
        text,
    )


def _extract_images(src: Path, dest: Path) -> dict[str, str]:
    """Extract embedded images from docx/pptx files.

    Saves images to an ``images/`` subdirectory next to *dest* and returns
    a mapping from the internal relationship path (e.g. ``image1.png``)
    to the relative Markdown reference path (e.g. ``images/image1.png``).
    """
    suffix = src.suffix.lower()
    if suffix not in (".docx", ".pptx"):
        return {}

    images_dir = dest.parent / "images"
    image_map: dict[str, str] = {}

    media_prefixes = {
        ".docx": "word/media/",
        ".pptx": "ppt/media/",
    }
    prefix = media_prefixes[suffix]

    try:
        with zipfile.ZipFile(str(src), "r") as zf:
            for entry in zf.namelist():
                if not entry.startswith(prefix):
                    continue
                filename = Path(entry).name
                if not filename or ".." in filename or "/" in filename:
                    continue
                if not filename.lower().endswith(
                    (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".svg", ".emf", ".wmf"),
                ):
                    continue
                images_dir.mkdir(parents=True, exist_ok=True)
                out_name = f"{src.stem}_{filename}"
                out_path = (images_dir / out_name).resolve()
                if not out_path.parent == images_dir.resolve():
                    continue
                out_path.write_bytes(zf.read(entry))
                rel_path = f"images/{out_name}"
                image_map[filename] = rel_path
    except (zipfile.BadZipFile, OSError):
        pass

    return image_map


def _replace_image_placeholders(
    text: str, image_map: dict[str, str],
) -> str:
    """Replace image references with extracted image paths.

    Handles two patterns produced by MarkItDown:
    1. Empty parens: ``![alt]()`` -- from base64-stripped images (docx)
    2. Non-path refs: ``![alt](SomeRef.jpg)`` -- from pptx slide images

    Extracted images are assigned in order of appearance.
    """
    if not image_map:
        return text

    image_paths = [
        path for _, path in sorted(image_map.items())
        if not path.lower().endswith((".emf", ".wmf"))
    ]
    if not image_paths:
        return text

    idx = 0

    def _replacer(m: re.Match[str]) -> str:
        nonlocal idx
        if idx >= len(image_paths):
            return m.group(0)
        alt = m.group(1)
        path = image_paths[idx]
        idx += 1
        return f"![{alt}]({path})"

    text = re.sub(r'!\[([^\]]*)\]\(\)', _replacer, text)

    if idx < len(image_paths):
        def _pptx_replacer(m: re.Match[str]) -> str:
            nonlocal idx
            ref = m.group(2)
            if ref.startswith(("http://", "https://", "images/")):
                return m.group(0)
            if idx >= len(image_paths):
                return m.group(0)
            alt = m.group(1)
            path = image_paths[idx]
            idx += 1
            return f"![{alt}]({path})"

        text = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', _pptx_replacer, text)

    return text
