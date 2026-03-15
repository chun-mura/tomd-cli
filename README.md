# tomd-cli

[![PyPI](https://img.shields.io/pypi/v/tomd-cli)](https://pypi.org/project/tomd-cli/)

Convert Office/PDF files to Markdown.

## Supported formats

- `.docx` (Word)
- `.pptx` (PowerPoint)
- `.xlsx` (Excel)
- `.pdf`

## Installation

```bash
# pipx (recommended)
pipx install tomd-cli

# or pip
pip install tomd-cli
```

### Update

```bash
# pipx
pipx upgrade tomd-cli

# or pip
pip install --upgrade tomd-cli
```

## Usage

### CLI

```bash
# Convert a single file
tomd document.docx

# Specify output path
tomd document.docx -o output.md

# Convert all files in a directory
tomd --dir ./documents/

# Convert directory with custom output
tomd --dir ./documents/ -o ./markdown/
```

### Python API

```python
from tomd import convert_file, convert_dir

# Single file
output_path = convert_file("document.docx")
output_path = convert_file("document.docx", "output.md")

# Directory
output_paths = convert_dir("./documents/")
output_paths = convert_dir("./documents/", "./markdown/")
```

## License

MIT
