# Quick Setup Guide

## Prerequisites

- Python 3.8 or higher
- Claude Code CLI configured with your API key

## Installation Steps

1. **Install dependencies:**

```bash
pip install -r requirements.txt
```

2. **Verify Claude SDK is working:**

```bash
python -c "from claude_agent_sdk import query; print('Claude SDK installed successfully')"
```

## Quick Test

1. **Create a test spreadsheet:**

```bash
python create_test_file.py
```

This will create `test_spreadsheet.xlsx` with some safe and suspicious formulas.

2. **Run the macro checker:**

```bash
python macro-checker.py test_spreadsheet.xlsx
```

3. **Check the output:**
   - A markdown report will be created with the analysis
   - A sanitized copy of the spreadsheet will be created with suspicious cells highlighted

## Common Issues

### oletools not installing

If you get errors installing `oletools`, you can still use the tool - it will just skip VBA macro detection:

```bash
pip install claude-agent-sdk openpyxl anyio
```

### Claude SDK authentication errors

Make sure you're logged into Claude Code:

```bash
# The claude-agent-sdk uses your Claude Code credentials
# Ensure you have an active Claude subscription
```

### Import errors

If you see import errors, make sure you're in a virtual environment or have installed all dependencies:

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Usage Examples

### Basic scan with default threshold (5)

```bash
python macro-checker.py my_file.xlsx
```

### More aggressive cleaning (removes scores < 8)

```bash
python macro-checker.py my_file.xlsx --remove-threshold 8
```

### Conservative cleaning (only removes very malicious, scores < 3)

```bash
python macro-checker.py my_file.xlsx --remove-threshold 3
```

### Save results to specific directory

```bash
python macro-checker.py my_file.xlsx --output-dir ./analysis_results
```

## Understanding the Scores

- **1-3 (Malicious)**: Definitely dangerous - file access, network calls, process execution, obfuscation
- **4-6 (Suspicious)**: Potentially dangerous - external references, dynamic execution
- **7-9 (Risky)**: May be legitimate but could be misused - common functions
- **10 (Safe)**: Harmless calculations and formulas

## Next Steps

Once you're comfortable with the tool:

1. Test with real spreadsheets from your organization
2. Adjust the `--remove-threshold` based on your risk tolerance
3. Review the markdown reports to understand what's being detected
4. Share sanitized copies with colleagues safely

## Need Help?

- Check the main README.md for detailed documentation
- Review the code in macro-checker.py for implementation details
- File issues or contribute improvements
