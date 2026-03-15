---
name: release
description: Release tomd-cli to PyPI and GitHub. Run tests, build, verify locally, publish, and tag.
disable-model-invocation: true
argument-hint: "[version]"
---

Release tomd-cli to PyPI and GitHub.

The argument $ARGUMENTS is the new version number (e.g. "0.1.3").
If no version is provided, ask the user what version to release.

## Steps

1. Update the version in `pyproject.toml` to the specified version.
2. Run tests: `PYTHONPATH=src python3 -m pytest tests/ -v`
   - If tests fail, stop and report the failure. Do NOT proceed.
3. Build the package: `rm -rf dist/ && python3 -m build`
4. Local verification:
   ```
   python3 -m venv /tmp/tomd-test-env
   /tmp/tomd-test-env/bin/pip install dist/tomd_cli-<version>-py3-none-any.whl
   /tmp/tomd-test-env/bin/tomd --help
   rm -rf /tmp/tomd-test-env
   ```
   - If install or `--help` fails, stop and report the failure.
5. Commit the version bump: `git add pyproject.toml && git commit -m "chore: bump version to <version>"`
6. Upload to PyPI: `python3 -m twine upload dist/*`
7. Tag and push: `git tag v<version> && git push origin v<version>`
8. Push the commit: `git push`
9. Report the result with links:
   - PyPI: https://pypi.org/project/tomd-cli/<version>/
   - Tag: v<version>
