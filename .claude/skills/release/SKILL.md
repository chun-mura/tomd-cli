---
name: release
description: Release tomd-cli to PyPI and GitHub. Run tests, build, verify locally, publish, and tag.
disable-model-invocation: true
argument-hint: "[version]"
---

Release tomd-cli to PyPI and GitHub.

The argument $ARGUMENTS is the new version number (e.g. "0.1.4").
If no version is provided, ask the user what version to release.

PyPI publish is automated via GitHub Actions (`.github/workflows/publish.yml`).
Pushing a `v*` tag triggers the workflow, which builds and publishes to PyPI
using Trusted Publisher (no API token needed).

## Steps

1. Update the version in `pyproject.toml` to the specified version.
2. Run tests: `PYTHONPATH=src python3 -m pytest tests/ -v`
   - If tests fail, stop and report the failure. Do NOT proceed.
3. Build the package locally to verify: `rm -rf dist/ && python3 -m build`
4. Local verification:
   ```
   python3 -m venv /tmp/tomd-test-env
   /tmp/tomd-test-env/bin/pip install dist/tomd_cli-<version>-py3-none-any.whl
   /tmp/tomd-test-env/bin/tomd --help
   rm -rf /tmp/tomd-test-env
   ```
   - If install or `--help` fails, stop and report the failure.
5. Commit the version bump: `git add pyproject.toml && git commit -m "chore: bump version to <version>"`
6. Create annotated tag: `git tag -a v<version> -m "v<version>"`
7. Push commit and tag: `git push origin main --tags`
   - This triggers GitHub Actions to automatically publish to PyPI.
8. Report the result with links:
   - PyPI: https://pypi.org/project/tomd-cli/<version>/
   - GitHub Actions: check the Actions tab for publish status
   - Tag: v<version>
