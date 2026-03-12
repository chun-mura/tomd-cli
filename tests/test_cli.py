"""Tests for tomd.cli."""

import pytest
from tomd.cli import main


def test_help_flag(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["--help"])
    assert exc_info.value.code == 0
    captured = capsys.readouterr()
    assert "Convert Office files" in captured.out


def test_missing_file(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["nonexistent.docx"])
    assert exc_info.value.code == 1
    captured = capsys.readouterr()
    assert "Error" in captured.err


def test_missing_dir(capsys):
    with pytest.raises(SystemExit) as exc_info:
        main(["--dir", "nonexistent_dir"])
    assert exc_info.value.code == 1
    captured = capsys.readouterr()
    assert "Error" in captured.err
