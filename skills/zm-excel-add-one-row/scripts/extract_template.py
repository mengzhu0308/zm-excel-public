#!/usr/bin/env python3
"""Extract Excel headers and sample row into a Markdown template."""

import argparse
import sys
from pathlib import Path
from typing import Optional

from _common import (
    assert_headers_present,
    detect_header_row,
    find_sample_row,
    load_workbook,
    read_headers,
    select_worksheet,
    validate_excel_extension,
)

try:
    import openpyxl  # noqa: F401  (imported for error messaging in __main__)
except ImportError:
    print("Error: openpyxl is required. Install with: pip install openpyxl")
    sys.exit(1)


def extract_template(
    excel_path: str,
    output_md: Optional[str] = None,
    sheet_name: Optional[str] = None,
    *,
    force: bool = False,
) -> str:
    """Extract headers and last row sample from Excel into Markdown template.

    Args:
        excel_path: Path to the Excel file.
        output_md: Optional explicit output path for the Markdown file.
                   Defaults to <excel_stem>_row_template.md in the same directory.
        sheet_name: Optional worksheet name to extract from.
                    Defaults to the active worksheet.
        force: If False (default), refuse to overwrite an existing template.
               If True, overwrite the existing file.

    Returns:
        Path to the generated Markdown file.

    Raises:
        FileNotFoundError: If the Excel file does not exist.
        FileExistsError: If the output Markdown file already exists and force is False.
        PermissionError: If the Excel file is locked or unwritable.
        ValueError: If the worksheet cannot be selected or no headers are found.
    """
    excel_file = Path(excel_path)
    if not excel_file.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    # Reject unsupported extensions early: .xls will produce confusing
    # zipfile errors if we let openpyxl try to open it.
    validate_excel_extension(excel_file)

    # openpyxl Workbook does NOT implement the context manager protocol,
    # so we release the file handle explicitly in a try/finally to avoid
    # file-lock leaks (notably on Windows).
    wb = load_workbook(str(excel_file), data_only=True)
    try:
        ws = select_worksheet(wb, sheet_name)

        header_row_idx = detect_header_row(ws)
        headers = read_headers(ws, header_row_idx)

        assert_headers_present(headers)

        sample_row_idx = find_sample_row(ws, header_row_idx)
        sample_row: dict[int, str] = {}
        if sample_row_idx is not None:
            for idx, cell in enumerate(ws[sample_row_idx], start=1):
                val = cell.value
                if val is not None:
                    sample_row[idx] = str(val)

        # Determine output path
        if output_md:
            md_path = Path(output_md)
        else:
            md_path = excel_file.with_name(f"{excel_file.stem}_row_template.md")

        # Refuse to overwrite an existing template unless force=True.
        # This protects user-edited content during cross-session resume.
        if md_path.exists() and not force:
            raise FileExistsError(
                f"Template already exists: {md_path}. "
                f"Refusing to overwrite. Use --force to overwrite explicitly."
            )

        lines: list[str] = []
        lines.append("## Excel 数据录入模板")
        lines.append("")
        lines.append(f"> 来源文件: `{excel_file.name}`")
        lines.append(f"> 工作表: `{ws.title}`")
        lines.append("")
        lines.append("请在下方的每个字段下填写实际内容，将每段末尾的占位符替换为你的数据。")
        lines.append("示例值来自表格中的已有数据，供你参考格式。")
        lines.append("")
        lines.append("---")
        lines.append("")

        for idx, header in enumerate(headers, start=1):
            if header == "":
                continue
            lines.append(f"# {header}")
            if idx in sample_row:
                lines.append(f"<!-- 示例: {sample_row[idx]} -->")
            lines.append("[在此填写]")
            lines.append("")

        md_path.write_text("\n".join(lines), encoding="utf-8")
    finally:
        wb.close()
    return str(md_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract Excel headers into a Markdown template")
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
    args = parser.parse_args()

    try:
        result = extract_template(args.excel, args.output, args.sheet, force=args.force)
        print(f"Template generated: {result}")
    except FileExistsError as e:
        print(f"Error: {e}")
        sys.exit(2)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except PermissionError as e:
        # PermissionError raised by openpyxl itself may have e.filename
        # set to None; fall back to a placeholder to avoid rendering
        # the bare string "None" in the user-facing message.
        fname = e.filename or "<unknown>"
        print(f"Error: Excel 文件被其他程序占用或无写权限: {fname}。请关闭 Excel 后重试。")
        sys.exit(1)
    except Exception as e:  # noqa: BLE001
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
