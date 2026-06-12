#!/usr/bin/env python3
"""Unified CLI entrypoint for the add-one-row workflow."""

import argparse
import sys

from _common import _setup_path, localize_error

# Make sibling imports resolvable when launched as a top-level script
# or via `python3 -m scripts.add_one_row` (where sys.path[0] is cwd,
# not the scripts/ directory).
_setup_path()

try:
    import openpyxl  # noqa: F401  (imported for error messaging in __main__)
except ImportError:
    print("Error: openpyxl is required. Install with: pip install openpyxl")
    sys.exit(1)


def _add_extract_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("excel", help="Path to the Excel file")
    parser.add_argument("--output", "-o", help="Output Markdown file path (optional)")
    parser.add_argument(
        "--sheet", "-s", metavar="NAME",
        help="Worksheet name to extract from (optional; defaults to active sheet)",
    )
    parser.add_argument(
        "--force", action="store_true",
        help="Overwrite an existing template file (default: refuse to overwrite)",
    )


def _add_write_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("excel", help="Path to the Excel file")
    parser.add_argument("md", help="Path to the filled Markdown template")
    parser.add_argument("--dry-run", action="store_true", help="Preview only, do not write file")
    parser.add_argument(
        "--sheet", "-s", metavar="NAME",
        help="Target worksheet name (overrides sheet in Markdown template; optional)",
    )
    parser.add_argument(
        "--output", "-o", metavar="PATH",
        help="Output xlsx path (optional; default: <stem>_增加一行.xlsx, auto-incremented if exists)",
    )
    parser.add_argument(
        "--force", action="store_true",
        help="Overwrite an existing --output file (default: refuse to overwrite explicit output)",
    )


# Mapping of exception type → exit code. Centralized here so the three
# entrypoints (add_one_row.py, extract_template.py main(), write_back.py
# main()) share the same exit-code contract.
_EXIT_CODE_BY_EXC_TYPE = {
    FileExistsError: 2,
    FileNotFoundError: 1,
    PermissionError: 1,
}


def _handle_error(exc: Exception) -> None:
    print(localize_error(exc))
    sys.exit(_EXIT_CODE_BY_EXC_TYPE.get(type(exc), 1))


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract a row template or append filled data to Excel")
    subparsers = parser.add_subparsers(dest="command", required=True)

    extract_parser = subparsers.add_parser("extract", help="Extract Excel headers into a Markdown template")
    _add_extract_args(extract_parser)

    write_parser = subparsers.add_parser("write", help="Append Markdown data back to Excel")
    _add_write_args(write_parser)

    args = parser.parse_args()

    try:
        if args.command == "extract":
            from extract_template import extract_template

            result = extract_template(args.excel, args.output, args.sheet, force=args.force)
            print(f"Template generated: {result}")
            return

        from write_back import write_back

        write_back(
            args.excel,
            args.md,
            dry_run=args.dry_run,
            sheet_name=args.sheet,
            output_path=args.output,
            force=args.force,
        )
    except Exception as exc:  # noqa: BLE001
        _handle_error(exc)


if __name__ == "__main__":
    main()
