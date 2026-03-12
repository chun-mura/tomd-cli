"""Command-line interface for tomd."""

from __future__ import annotations

import argparse
import sys

from tomd.converter import SUPPORTED_EXTENSIONS, convert_dir, convert_file


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="tomd",
        description="Convert Office files to Markdown using MarkItDown.",
    )
    parser.add_argument(
        "input",
        help="file or directory to convert",
    )
    parser.add_argument(
        "-o", "--output",
        help="output file or directory path",
    )
    parser.add_argument(
        "--dir",
        action="store_true",
        dest="is_dir",
        help="treat input as a directory and convert all supported files",
    )
    fmt = ", ".join(sorted(SUPPORTED_EXTENSIONS))
    parser.epilog = f"Supported formats: {fmt}"
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = _build_parser()
    args = parser.parse_args(argv)

    try:
        if args.is_dir:
            results = convert_dir(args.input, args.output)
            if not results:
                print(f"No convertible files found in {args.input}")
            else:
                for dest in results:
                    print(f"Converted: {dest}")
        else:
            dest = convert_file(args.input, args.output)
            print(f"Converted: {dest}")
    except (FileNotFoundError, NotADirectoryError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Conversion failed: {e}", file=sys.stderr)
        sys.exit(1)
