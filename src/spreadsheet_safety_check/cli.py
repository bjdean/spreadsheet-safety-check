"""Command-line interface for Spreadsheet Safety Check."""

import argparse
import sys
from datetime import datetime
from pathlib import Path

import anyio

from spreadsheet_safety_check.checker import MacroChecker


async def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Detect and analyze macros/code in spreadsheets (Excel and OpenOffice)"
    )
    parser.add_argument(
        "input_file", help="Path to the spreadsheet to analyze (.xlsx, .xlsm, .ods)"
    )
    parser.add_argument(
        "--remove-threshold",
        type=int,
        default=5,
        help="Score threshold for removing code (default: 5, removes items with score < 5)",
    )
    parser.add_argument(
        "--output-dir",
        help="Directory for output files (default: same as input file)",
    )

    args = parser.parse_args()

    # Validate input file
    input_path = Path(args.input_file)
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        sys.exit(1)

    if input_path.suffix.lower() not in [".xlsx", ".xlsm", ".ods"]:
        print("Error: Unsupported file format. Use .xlsx, .xlsm, or .ods files")
        sys.exit(1)

    # Determine output directory
    if args.output_dir:
        output_dir = Path(args.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    else:
        output_dir = input_path.parent

    # Create checker instance
    checker = MacroChecker(input_path, args.remove_threshold)

    # Scan the file
    success = await checker.scan_file()

    if not success:
        print("Scanning failed")
        sys.exit(1)

    # Generate report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_file = output_dir / f"{input_path.stem}_report_{timestamp}.md"

    report_content = checker.generate_markdown_report()
    report_file.write_text(report_content)

    print(f"\nReport saved to: {report_file}")

    # Create sanitized copy
    sanitized_file = (
        output_dir / f"{input_path.stem}_sanitized_{timestamp}{input_path.suffix}"
    )
    checker.create_sanitized_copy(sanitized_file)

    print("\n=== Analysis Complete ===")


def cli_entry_point():
    """Entry point for console scripts."""
    anyio.run(main)


if __name__ == "__main__":
    cli_entry_point()
