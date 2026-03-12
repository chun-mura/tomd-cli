"""tomd-cli: Convert Office files to Markdown using MarkItDown."""

from tomd.converter import convert_file, convert_dir, strip_base64_images

__all__ = ["convert_file", "convert_dir", "strip_base64_images"]
