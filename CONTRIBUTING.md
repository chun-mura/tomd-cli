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

3. Local verification (install into a temporary venv and test)

   ```bash
   python3 -m venv /tmp/tomd-test-env
   /tmp/tomd-test-env/bin/pip install dist/tomd_cli-<version>-py3-none-any.whl

   # Verify CLI works
   /tmp/tomd-test-env/bin/tomd --help
   /tmp/tomd-test-env/bin/tomd /path/to/file.pdf -o /tmp/output.md

   # Clean up
   rm -rf /tmp/tomd-test-env
   ```

4. Upload to PyPI

   ```bash
   twine upload dist/*
   ```

5. Tag the release

   ```bash
   git tag v<version>
   git push origin v<version>
   ```
