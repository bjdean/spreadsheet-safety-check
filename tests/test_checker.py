"""Tests for the MacroChecker class."""

from pathlib import Path

import pytest

from spreadsheet_safety_check.checker import MacroChecker, MacroFinding


class TestMacroFinding:
    """Tests for MacroFinding dataclass."""

    def test_macro_finding_creation(self):
        """Test creating a MacroFinding instance."""
        finding = MacroFinding(
            item_number=1,
            location="Sheet1!A1",
            code="=SUM(1,2)",
            score=10,
            analysis="Safe formula",
        )

        assert finding.item_number == 1
        assert finding.location == "Sheet1!A1"
        assert finding.code == "=SUM(1,2)"
        assert finding.score == 10
        assert finding.analysis == "Safe formula"
        assert finding.cell_reference is None

    def test_macro_finding_with_cell_reference(self):
        """Test MacroFinding with cell reference."""
        finding = MacroFinding(
            item_number=2,
            location="Sheet1!B2",
            code="=HYPERLINK('http://example.com')",
            score=5,
            analysis="Suspicious",
            cell_reference=("Sheet1", "B", 2),
        )

        assert finding.cell_reference == ("Sheet1", "B", 2)


class TestMacroChecker:
    """Tests for MacroChecker class."""

    def test_initialization(self, temp_xlsx_file):
        """Test MacroChecker initialization."""
        checker = MacroChecker(temp_xlsx_file, remove_threshold=5)

        assert checker.input_file == Path(temp_xlsx_file)
        assert checker.remove_threshold == 5
        assert checker.findings == []
        assert checker.workbook is None
        assert checker.item_counter == 0

    def test_load_excel_spreadsheet(self, temp_xlsx_file):
        """Test loading an Excel spreadsheet."""
        checker = MacroChecker(temp_xlsx_file)
        result = checker.load_spreadsheet()

        assert result is True
        assert checker.file_type == "excel"
        assert checker.workbook is not None

    def test_load_unsupported_file(self, tmp_path):
        """Test loading an unsupported file format."""
        unsupported_file = tmp_path / "test.txt"
        unsupported_file.write_text("not a spreadsheet")

        checker = MacroChecker(unsupported_file)
        result = checker.load_spreadsheet()

        assert result is False

    def test_detect_excel_formulas(self, temp_xlsx_file):
        """Test detecting formulas in Excel files."""
        checker = MacroChecker(temp_xlsx_file)
        checker.load_spreadsheet()

        formulas = checker.detect_formula_cells()

        # Should find the two formulas we added (=SUM(1,2) and =HYPERLINK(...))
        assert len(formulas) == 2
        assert any("SUM" in formula[1] for formula in formulas)
        assert any("HYPERLINK" in formula[1] for formula in formulas)

    def test_generate_markdown_report_empty(self, temp_xlsx_file):
        """Test generating a report with no findings."""
        checker = MacroChecker(temp_xlsx_file)
        checker.load_spreadsheet()

        report = checker.generate_markdown_report()

        assert "# Macro Security Analysis Report" in report
        assert checker.input_file.name in report
        assert "Total Items Found:** 0" in report

    def test_generate_markdown_report_with_findings(self, temp_xlsx_file):
        """Test generating a report with findings."""
        checker = MacroChecker(temp_xlsx_file)
        checker.load_spreadsheet()

        # Add a mock finding
        finding = MacroFinding(
            item_number=1,
            location="Sheet1!A2",
            code="=SUM(1,2)",
            score=10,
            analysis="Safe formula",
            cell_reference=("Sheet1", "A", 2),
        )
        checker.findings.append(finding)

        report = checker.generate_markdown_report()

        assert "Total Items Found:** 1" in report
        assert "Sheet1!A2" in report
        assert "=SUM(1,2)" in report
        assert "Safe formula" in report

    @pytest.mark.asyncio
    async def test_analyze_code_with_claude_format(self, monkeypatch):
        """Test Claude analysis response parsing."""
        # This test mocks the Claude API to test parsing logic
        from spreadsheet_safety_check.checker import MacroChecker

        checker = MacroChecker("dummy.xlsx")

        # Mock the query function
        async def mock_query(*args, **kwargs):
            from claude_agent_sdk import AssistantMessage, TextBlock

            # Create a proper AssistantMessage with required model parameter
            yield AssistantMessage(
                content=[TextBlock(text="SCORE: 8\nANALYSIS: This is a test analysis")],
                model="claude-test",
            )

        monkeypatch.setattr("spreadsheet_safety_check.checker.query", mock_query)

        score, analysis = await checker.analyze_code_with_claude(
            "=SUM(1,2)", "Sheet1!A1"
        )

        assert score == 8
        assert "test analysis" in analysis.lower()

    def test_findings_sorting_by_score(self, temp_xlsx_file):
        """Test that findings are sorted by score in the report."""
        checker = MacroChecker(temp_xlsx_file)

        # Add findings with different scores
        checker.findings = [
            MacroFinding(1, "A1", "code1", 8, "analysis1"),
            MacroFinding(2, "A2", "code2", 3, "analysis2"),
            MacroFinding(3, "A3", "code3", 10, "analysis3"),
        ]

        report = checker.generate_markdown_report()

        # Report should list items sorted by score (lowest first)
        a2_pos = report.find("A2")  # score 3
        a1_pos = report.find("A1")  # score 8
        a3_pos = report.find("A3")  # score 10

        assert a2_pos < a1_pos < a3_pos
