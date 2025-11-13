"""Tests for the CLI module."""

import sys
from unittest.mock import patch

import pytest

from spreadsheet_safety_check.cli import main


class TestCLI:
    """Tests for the command-line interface."""

    @pytest.mark.asyncio
    async def test_main_with_nonexistent_file(self, capsys):
        """Test CLI with a file that doesn't exist."""
        with patch.object(sys, "argv", ["prog", "nonexistent.xlsx"]):
            with pytest.raises(SystemExit) as exc_info:
                await main()

            assert exc_info.value.code == 1

        captured = capsys.readouterr()
        assert "not found" in captured.out.lower()

    @pytest.mark.asyncio
    async def test_main_with_unsupported_format(self, tmp_path, capsys):
        """Test CLI with unsupported file format."""
        unsupported_file = tmp_path / "test.txt"
        unsupported_file.write_text("not a spreadsheet")

        with patch.object(sys, "argv", ["prog", str(unsupported_file)]):
            with pytest.raises(SystemExit) as exc_info:
                await main()

            assert exc_info.value.code == 1

        captured = capsys.readouterr()
        assert "unsupported" in captured.out.lower()

    @pytest.mark.asyncio
    async def test_main_with_valid_file(self, temp_xlsx_file, monkeypatch):
        """Test CLI with a valid Excel file."""

        # Mock the scan_file method to avoid actual Claude API calls
        async def mock_scan_file(self):
            return True

        monkeypatch.setattr(
            "spreadsheet_safety_check.checker.MacroChecker.scan_file", mock_scan_file
        )

        with patch.object(sys, "argv", ["prog", str(temp_xlsx_file)]):
            await main()

        # If we get here without exception, the test passed
        assert True

    @pytest.mark.asyncio
    async def test_main_with_custom_threshold(self, temp_xlsx_file, monkeypatch):
        """Test CLI with custom removal threshold."""

        # Mock the scan_file method
        async def mock_scan_file(self):
            return True

        monkeypatch.setattr(
            "spreadsheet_safety_check.checker.MacroChecker.scan_file", mock_scan_file
        )

        # Mock MacroChecker.__init__ to capture the threshold
        original_init = None
        captured_threshold = None

        def capture_init(self, input_file, remove_threshold=5):
            nonlocal captured_threshold
            captured_threshold = remove_threshold
            original_init(self, input_file, remove_threshold)

        from spreadsheet_safety_check.checker import MacroChecker

        original_init = MacroChecker.__init__
        monkeypatch.setattr(MacroChecker, "__init__", capture_init)

        with patch.object(
            sys, "argv", ["prog", str(temp_xlsx_file), "--remove-threshold", "7"]
        ):
            await main()

        assert captured_threshold == 7

    @pytest.mark.asyncio
    async def test_main_with_output_dir(self, temp_xlsx_file, tmp_path, monkeypatch):
        """Test CLI with custom output directory."""
        output_dir = tmp_path / "output"

        # Mock the scan_file method
        async def mock_scan_file(self):
            return True

        monkeypatch.setattr(
            "spreadsheet_safety_check.checker.MacroChecker.scan_file", mock_scan_file
        )

        with patch.object(
            sys, "argv", ["prog", str(temp_xlsx_file), "--output-dir", str(output_dir)]
        ):
            await main()

        # Check that the output directory was created
        assert output_dir.exists()
        assert output_dir.is_dir()

    @pytest.mark.asyncio
    async def test_scan_failure(self, temp_xlsx_file, monkeypatch, capsys):
        """Test CLI behavior when scanning fails."""

        # Mock scan_file to return False
        async def mock_scan_file(self):
            return False

        monkeypatch.setattr(
            "spreadsheet_safety_check.checker.MacroChecker.scan_file", mock_scan_file
        )

        with patch.object(sys, "argv", ["prog", str(temp_xlsx_file)]):
            with pytest.raises(SystemExit) as exc_info:
                await main()

            assert exc_info.value.code == 1

        captured = capsys.readouterr()
        assert "failed" in captured.out.lower()
