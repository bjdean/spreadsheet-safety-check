# Macro Checker

A Python tool that detects and analyzes macros and suspicious code in Excel spreadsheets using AI-powered analysis via Claude.

## Features

- **VBA Macro Detection**: Extracts and analyzes VBA macros from `.xlsm` files
- **Formula Analysis**: Detects potentially dangerous formulas (HYPERLINK, WEBSERVICE, etc.)
- **AI-Powered Scoring**: Uses Claude SDK to score each macro/code from 1-10
  - 1-3: Definitely malicious
  - 4-6: Suspicious
  - 7-9: Potentially risky but may be legitimate
  - 10: Safe
- **Markdown Reports**: Generates detailed analysis reports
- **Sanitized Copies**: Creates a cleaned version of the spreadsheet with suspicious code removed and highlighted in yellow

## Installation

1. Install Python dependencies:

```bash
pip install -r requirements.txt
```

2. Ensure you have Claude CLI configured with your API key and you are logged in with Claude Code. claude-agent-sdk uses your Claude Code credentials.

## Usage

Basic usage:

```bash
python macro-checker.py input_spreadsheet.xlsx
```

With custom removal threshold:

```bash
python macro-checker.py input_spreadsheet.xlsm --remove-threshold 7
```

With custom output directory:

```bash
python macro-checker.py input_spreadsheet.xlsx --output-dir ./results
```

### Command-Line Arguments

- `input_file` (required): Path to the Excel spreadsheet to analyze (.xlsx or .xlsm)
- `--remove-threshold N`: Score threshold for removing code (default: 5). Items with score < N will be removed
- `--output-dir DIR`: Directory for output files (default: same directory as input file)

## Output Files

The tool generates two files:

1. **Report**: `{filename}_report_{timestamp}.md`
   - Summary statistics
   - Detailed analysis of each finding
   - Code snippets with scores

2. **Sanitized Copy**: `{filename}_sanitized_{timestamp}.{xlsx|xlsm}`
   - Copy of original spreadsheet
   - Cells with score < threshold replaced with "CODE REMOVED: Item #N"
   - Highlighted in bright yellow

## Example

```bash
$ python macro-checker.py suspicious_file.xlsm --remove-threshold 6

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

## Future Enhancements

- [ ] Support for OpenOffice/LibreOffice formats (.ods)
- [ ] Support for older Excel formats (.xls)
- [ ] Batch processing of multiple files
- [ ] Custom pattern definitions
- [ ] HTML report output
- [ ] Quarantine mode for high-risk files

## Dependencies

- `claude-agent-sdk`: Claude AI integration
- `openpyxl`: Excel file reading/writing
- `oletools`: VBA macro extraction
- `anyio`: Async runtime

## License

MIT

## Author

Created for detecting and analyzing potentially malicious macros in spreadsheets.
