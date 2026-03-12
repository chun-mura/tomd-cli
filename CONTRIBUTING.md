# Contributing

## Development Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
```

## Running Tests

```bash
pytest
```

## Release

1. Update the version in `pyproject.toml`
2. Build the distribution

   ```bash
   python -m build
   ```

3. Upload to PyPI

   ```bash
   twine upload dist/*
   ```

4. Tag the release

   ```bash
   git tag v<version>
   git push origin v<version>
   ```
