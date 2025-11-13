# Quick Start Guide

## Installation

```bash
pip install spreadsheet-safety-check
```

## Basic Usage

```bash
# Analyze a spreadsheet
spreadsheet-safety-check your_file.xlsx

# Analyze with custom threshold
spreadsheet-safety-check your_file.xlsx --remove-threshold 7

# Save results to specific directory
spreadsheet-safety-check your_file.xlsx --output-dir ./results
```

## What You'll Get

1. **Markdown Report** - Detailed analysis with scores and explanations
2. **Sanitized Copy** - Cleaned spreadsheet with suspicious cells highlighted in yellow

## Understanding Scores

- **1-3**: Malicious (file access, network calls, process execution)
- **4-6**: Suspicious (external references, dynamic execution)
- **7-9**: Potentially risky (common functions that could be misused)
- **10**: Safe (simple calculations)

## Development

```bash
# Clone and setup
git clone https://github.com/bjdean/spreadsheet-safety-check.git
cd spreadsheet-safety-check
pip install -e ".[dev]"

# Run tests
pytest

# Check code quality
make check
```

## Creating Test Files

```bash
cd examples
python create_test_files.py
spreadsheet-safety-check test_spreadsheet.xlsx
```

For more details, see [README.md](README.md)
