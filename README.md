# Spreadsheet Safety Check

[![CI](https://github.com/bjdean/spreadsheet-safety-check/workflows/CI/badge.svg)](https://github.com/bjdean/spreadsheet-safety-check/actions)
[![PyPI version](https://badge.fury.io/py/spreadsheet-safety-check.svg)](https://badge.fury.io/py/spreadsheet-safety-check)
[![Python Versions](https://img.shields.io/pypi/pyversions/spreadsheet-safety-check.svg)](https://pypi.org/project/spreadsheet-safety-check/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A Python tool that detects and analyzes macros and suspicious code in spreadsheets (Excel and OpenOffice) using AI-powered analysis via Claude.

## Features

- **VBA Macro Detection**: Extracts and analyzes VBA macros from `.xlsm` files
- **OpenOffice Support**: Analyzes `.ods` files from OpenOffice/LibreOffice
- **Formula Analysis**: Detects potentially dangerous formulas (HYPERLINK, WEBSERVICE, etc.)
- **AI-Powered Scoring**: Uses Claude SDK to score each macro/code from 1-10
  - 1-3: Definitely malicious
  - 4-6: Suspicious
  - 7-9: Potentially risky but may be legitimate
  - 10: Safe
- **Markdown Reports**: Generates detailed analysis reports
- **Sanitized Copies**: Creates a cleaned version of the spreadsheet with suspicious code removed and highlighted in yellow

## Installation

### From PyPI (Recommended)

```bash
pip install spreadsheet-safety-check
```

### From Source

```bash
git clone https://github.com/bjdean/spreadsheet-safety-check.git
cd spreadsheet-safety-check
pip install -e .
```

### Prerequisites

Ensure you have Claude CLI configured with your API key and you are logged in with Claude Code. The `claude-agent-sdk` uses your Claude Code credentials.

## Usage

### Command Line

Basic usage:

```bash
spreadsheet-safety-check input_spreadsheet.xlsx
```

With custom removal threshold:

```bash
spreadsheet-safety-check input_spreadsheet.xlsm --remove-threshold 7
```

With custom output directory:

```bash
spreadsheet-safety-check input_spreadsheet.xlsx --output-dir ./results
```

### Python API

```python
import asyncio
from spreadsheet_safety_check import MacroChecker

async def analyze_spreadsheet():
    checker = MacroChecker("suspicious_file.xlsx", remove_threshold=5)
    await checker.scan_file()

    # Generate report
    report = checker.generate_markdown_report()
    print(report)

    # Create sanitized copy
    checker.create_sanitized_copy("sanitized_output.xlsx")

asyncio.run(analyze_spreadsheet())
```

### Command-Line Arguments

- `input_file` (required): Path to the spreadsheet to analyze (`.xlsx`, `.xlsm`, or `.ods`)
- `--remove-threshold N`: Score threshold for removing code (default: 5). Items with score < N will be removed
- `--output-dir DIR`: Directory for output files (default: same directory as input file)

## Output Files

The tool generates two files:

1. **Report**: `{filename}_report_{timestamp}.md`
   - Summary statistics
   - Detailed analysis of each finding
   - Code snippets with scores

2. **Sanitized Copy**: `{filename}_sanitized_{timestamp}.{xlsx|xlsm|ods}`
   - Copy of original spreadsheet
   - Cells with score < threshold replaced with "CODE REMOVED: Item #N"
   - Highlighted in bright yellow

## Example

```bash
$ spreadsheet-safety-check suspicious_file.xlsm --remove-threshold 6

Loaded spreadsheet: suspicious_file.xlsm

=== Scanning for macros and suspicious code ===

Checking for VBA macros...
VBA macros detected!
Analyzing VBA macro 1: VBA Module: Module1...
  Score: 2/10
Analyzing VBA macro 2: VBA Module: ThisWorkbook...
  Score: 8/10

Checking for suspicious formulas...
Analyzing formula 3: Sheet1!A5...
  Score: 4/10

=== Scan complete: 3 items found ===

Report saved to: suspicious_file_report_20231113_143022.md

Sanitized copy created: suspicious_file_sanitized_20231113_143022.xlsm
Items removed/highlighted: 2

=== Analysis Complete ===
```

## Detected Patterns

The tool looks for:

### VBA Macros
- All VBA code in macro-enabled files (.xlsm)
- Analyzes each module separately

### Suspicious Formulas
- HYPERLINK - Can link to external resources
- WEBSERVICE - Makes HTTP requests
- FILTERXML - Can parse external XML
- IMPORTDATA - Imports external data
- CALL - Calls external functions
- REGISTER - Registers external functions
- EXEC - Executes commands
- SHELL - Executes shell commands
- INDIRECT - Dynamic cell references

## Security Considerations

- The tool requires Claude SDK access which uses your API credits
- VBA macro detection requires the `oletools` library
- The tool does NOT execute any macros - it only reads and analyzes them
- Sanitized copies preserve the original file structure

## Development

### Setup Development Environment

```bash
# Clone the repository
git clone https://github.com/bjdean/spreadsheet-safety-check.git
cd spreadsheet-safety-check

# Create a virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install in development mode with dev dependencies
pip install -e ".[dev]"
```

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=spreadsheet_safety_check --cov-report=html

# Run specific test file
pytest tests/test_checker.py
```

### Code Quality

```bash
# Format code with black
black src tests

# Lint with ruff
ruff check src tests

# Type check with mypy
mypy src/spreadsheet_safety_check
```

### Building the Package

```bash
# Install build tools
pip install build

# Build the package
python -m build
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Run tests and linting
5. Commit your changes (`git commit -m 'Add some amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## Future Enhancements

- [x] Support for OpenOffice/LibreOffice formats (.ods)
- [ ] Support for older Excel formats (.xls)
- [ ] Batch processing of multiple files
- [ ] Custom pattern definitions
- [ ] HTML report output
- [ ] Quarantine mode for high-risk files

## Dependencies

### Required
- `claude-agent-sdk`: Claude AI integration
- `openpyxl`: Excel file reading/writing
- `oletools`: VBA macro extraction
- `anyio`: Async runtime
- `odfpy`: OpenOffice format support

### Development
- `pytest`: Testing framework
- `pytest-cov`: Coverage reporting
- `pytest-asyncio`: Async test support
- `ruff`: Fast Python linter
- `black`: Code formatter
- `mypy`: Static type checker

## License

MIT

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history.

## Support

If you encounter any issues or have questions:
- Open an issue on [GitHub](https://github.com/bjdean/spreadsheet-safety-check/issues)
- Check existing issues for solutions

## Citation

If you use this tool in your research or project, please consider citing it:

```
@software{spreadsheet_safety_check,
  title = {Spreadsheet Safety Check},
  author = {Contributors},
  year = {2025},
  url = {https://github.com/bjdean/spreadsheet-safety-check}
}
```
