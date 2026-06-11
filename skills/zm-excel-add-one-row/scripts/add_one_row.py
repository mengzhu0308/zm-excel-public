#!/usr/bin/env python3
"""Unified CLI entrypoint for the add-one-row workflow."""

import argparse
import sys


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


def _handle_error(exc: Exception) -> None:
    if isinstance(exc, FileExistsError):
        print(f"Error: {exc}")
        sys.exit(2)
    if isinstance(exc, FileNotFoundError):
        print(f"Error: {exc}")
        sys.exit(1)
    if isinstance(exc, PermissionError):
        fname = exc.filename or "<unknown>"
        print(f"Error: Excel 文件被其他程序占用或无写权限: {fname}。请关闭 Excel 后重试。")
        sys.exit(1)

    print(f"Error: {exc}")
    sys.exit(1)


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
