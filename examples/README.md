# Examples

This directory contains example scripts and test files for Spreadsheet Safety Check.

## Creating Test Files

Run the test file generator to create sample spreadsheets with both safe and suspicious formulas:

```bash
python create_test_files.py
```

This will create:
- `test_spreadsheet.xlsx` - Excel file with test formulas
- `test_spreadsheet.ods` - OpenOffice file with test formulas

## Testing the Tool

Once you've created the test files, analyze them with:

```bash
# Analyze Excel file
spreadsheet-safety-check test_spreadsheet.xlsx

# Analyze OpenOffice file
spreadsheet-safety-check test_spreadsheet.ods

# Use custom threshold
spreadsheet-safety-check test_spreadsheet.xlsx --remove-threshold 7

# Save to custom directory
spreadsheet-safety-check test_spreadsheet.xlsx --output-dir ./results
```

## What's in the Test Files

The test files contain:

### Safe Formulas
- `=SUM(B1:B10)` - Simple sum
- `=AVERAGE(C1:C5)` - Average calculation
- `=IF(D1>10, "Yes", "No")` - Conditional logic

### Suspicious Formulas
- `=HYPERLINK("http://example.com")` - External links
- `=WEBSERVICE("http://api.example.com")` - HTTP requests
- `=INDIRECT("A1")` - Dynamic cell references

### Malicious Patterns (for testing only!)
- DDE commands
- EXEC functions
- External function calls
- File system operations

**Note:** These test files are safe - they contain formula text but will not execute malicious code. They are designed to test the detection capabilities of the tool.
