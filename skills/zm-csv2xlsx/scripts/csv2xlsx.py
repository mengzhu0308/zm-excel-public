#!/usr/bin/env python3
"""
CSV 转 Excel 脚本：支持单文件、目录批量、文件列表三种输入模式。
多 CSV 可合并到一个 Excel 的不同 sheet 中。
"""

import argparse
import re
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError as e:
    print("错误: 缺少 pandas。请安装: pip install pandas openpyxl", file=sys.stderr)
    sys.exit(1)

try:
    import openpyxl
except ImportError as e:
    print("错误: 缺少 openpyxl。请安装: pip install openpyxl", file=sys.stderr)
    sys.exit(1)

def _get_version():
    """从 VERSION.yaml 读取版本号。"""
    try:
        version_file = Path(__file__).parent.parent / "VERSION.yaml"
        with open(version_file, "r", encoding="utf-8") as f:
            in_skill_info = False
            for line in f:
                stripped = line.strip()
                if stripped.startswith("#"):
                    continue
                if stripped == "skill_info:" or stripped.startswith("skill_info:"):
                    in_skill_info = True
                    continue
                if in_skill_info:
                    if stripped.startswith("version:"):
                        return stripped.split("version:", 1)[1].strip().strip('"').strip("'")
                    # 如果遇到非缩进行，离开 skill_info 块
                    if stripped and not line.startswith((" ", "	")):
                        in_skill_info = False
    except Exception:
        pass
    return "unknown"


def split_file_list(text):
    """按逗号、顿号、分号、换行拆分文件列表，去空去重。保留文件名中的空格。"""
    lines = text.splitlines()
    parts = []
    for line in lines:
        parts.extend(re.split(r"[,，、;；]+", line))
    return [p.strip() for p in parts if p.strip()]


def detect_encoding(file_path, max_lines=100):
    """尝试检测 CSV 文件编码，返回编码名。

    通过读取前 max_lines 行来判断编码，避免大文件全量读取导致内存爆炸。
    """
    if max_lines <= 0:
        max_lines = 100
    # 优先级: utf-8-sig > utf-8 > gb18030
    encodings = ["utf-8-sig", "utf-8", "gb18030"]
    for enc in encodings:
        try:
            with open(file_path, "r", encoding=enc) as f:
                for _ in range(max_lines):
                    if not f.readline():
                        break
            return enc
        except UnicodeDecodeError:
            continue
    raise ValueError(f"无法自动检测编码，请使用 -e 手动指定编码。文件: {file_path}")


def collect_csv_files(input_arg):
    """
    根据输入参数收集所有 CSV 文件路径。
    返回: list of Path
    """
    arg = input_arg.strip()
    p = Path(arg)

    # 情况1: 目录
    if p.is_dir():
        files = sorted([f for f in p.iterdir() if f.is_file() and f.suffix.lower() == ".csv"])
        return files

    # 情况2: 单文件直接存在
    if p.is_file() and p.suffix.lower() == ".csv":
        return [p]

    # 情况3: 解析为文件列表
    candidates = split_file_list(arg)
    files = []
    missing = []
    for c in candidates:
        cp = Path(c)
        if cp.is_file() and cp.suffix.lower() == ".csv":
            files.append(cp)
        else:
            missing.append(c)

    if not files and missing:
        print(f"错误: 无法识别任何有效的 CSV 文件。输入: {arg}", file=sys.stderr)
        sys.exit(1)

    if missing:
        print(f"警告: 跳过无效路径: {', '.join(missing)}", file=sys.stderr)

    return files


def _sheet_name_for_filename(csv_stem):
    """将文件名处理为安全的 sheet 名。"""
    safe = re.sub(r'[\\/:*?"<>|]', '_', str(csv_stem))
    # openpyxl 限制 sheet 名长度不超过 31
    if len(safe) > 31:
        safe = safe[:28] + '...'
    return safe.strip()


def _resolve_output_path(input_path, output_arg):
    """根据输入和输出参数解析最终输出路径。"""
    if not output_arg:
        return input_path.with_suffix(".xlsx")

    out_path = Path(output_arg)
    if out_path.suffix.lower() == ".xlsx":
        out_path.parent.mkdir(parents=True, exist_ok=True)
        return out_path
    elif str(out_path).endswith(("/", "\\")) or out_path.is_dir():
        out_path.mkdir(parents=True, exist_ok=True)
        return out_path / f"{input_path.stem}.xlsx"
    else:
        # 无后缀路径，自动追加 .xlsx（与合并模式行为一致）
        out_path = out_path.with_suffix(".xlsx")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        return out_path


def convert_single_file(input_path, output_dir=None, sheet_name=None, encoding=None, header=None):
    """
    转换单个 CSV 文件为 xlsx。
    返回: (sheet_name_used, output_xlsx_path)
    """
    input_path = Path(input_path)
    enc = encoding or detect_encoding(input_path)
    df = pd.read_csv(input_path, encoding=enc, header=header)

    out_path = _resolve_output_path(input_path, output_dir)

    sname = _sheet_name_for_filename(sheet_name) if sheet_name else _sheet_name_for_filename(input_path.stem)
    df.to_excel(out_path, sheet_name=sname, index=False, engine="openpyxl")

    return (sname, out_path)


def convert_combine(files, output_path, encoding=None, header=None):
    """
    将多个 CSV 合并到一个 Excel 中，每个 CSV 作为一个 sheet。
    返回: (list of (sheet_name, output_xlsx_path), list of failed_paths)
    """
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    results = []
    failed = []

    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            for f in files:
                f = Path(f)
                try:
                    enc = encoding or detect_encoding(f)
                    df = pd.read_csv(f, encoding=enc, header=header)
                    sname = _sheet_name_for_filename(f.stem)
                    # 避免 sheet 名冲突
                    original_sname = sname
                    counter = 1
                    while sname in writer.sheets:
                        sname = f"{original_sname}_{counter}"
                        counter += 1
                    df.to_excel(writer, sheet_name=sname, index=False)
                    results.append((sname, out_path))
                except Exception as e:
                    print(f"错误: 合并时读取 {f} 失败: {e}", file=sys.stderr)
                    failed.append(str(f))
    except IndexError as e:
        # 空工作簿保存失败（所有文件均读取失败时），删除残留文件
        if "At least one sheet must be visible" in str(e) and not results and failed:
            if out_path.exists():
                out_path.unlink()
            return [], failed
        raise

    # 如果没有任何成功结果，删除空输出文件
    if not results and out_path.exists():
        out_path.unlink()

    return results, failed


def main():
    parser = argparse.ArgumentParser(
        description="将 CSV 文件转换为 Excel 文件(.xlsx)。"
    )
    parser.add_argument(
        "input",
        help=(
            "输入：单个 CSV 文件路径、包含 CSV 文件的目录路径、"
            "或以逗号、顿号、分号或换行分隔的多个 CSV 文件路径"
        ),
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="输出路径或目录（默认与源文件同目录，单文件时自动推导 xlsx 文件名）",
    )
    parser.add_argument(
        "-e", "--encoding",
        default=None,
        help="CSV 文件编码（默认自动检测: utf-8-sig / utf-8 / gb18030）",
    )
    parser.add_argument(
        "-n", "--sheet-name",
        default=None,
        help="指定输出 Excel 中的 sheet 名（单文件模式；默认使用文件名）",
    )
    parser.add_argument(
        "--combine",
        action="store_true",
        help="将多个 CSV 合并到一个 Excel 中，每个 CSV 作为一个 sheet（需配合 -o 指定输出文件）",
    )
    parser.add_argument(
        "--no-header",
        action="store_true",
        help="CSV 没有表头行",
    )
    parser.add_argument(
        "-v", "--version",
        action="version",
        version=f"%(prog)s {_get_version()}",
    )
    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="静默模式，只输出错误信息",
    )

    args = parser.parse_args()

    def log(msg):
        if not args.quiet:
            print(msg)

    files = collect_csv_files(args.input)

    if not files:
        print("未找到任何 CSV 文件。", file=sys.stderr)
        sys.exit(1)

    header = None if args.no_header else 0

    if args.combine:
        # 合并模式：所有 CSV 写入一个 xlsx
        if not args.output:
            print("错误: --combine 模式需要配合 -o 指定输出文件路径。", file=sys.stderr)
            sys.exit(1)
        if args.sheet_name:
            print("警告: --sheet-name 在合并模式下无效，将使用 CSV 文件名作为 sheet 名。", file=sys.stderr)
        out_path = Path(args.output)
        if out_path.suffix.lower() != ".xlsx":
            out_path = out_path.with_suffix(".xlsx")
        results, failed = convert_combine(files, out_path, args.encoding, header)
        if not results:
            print(f"错误: 所有文件合并失败。", file=sys.stderr)
            if failed:
                print(f"失败文件: {', '.join(failed)}", file=sys.stderr)
            sys.exit(1)
        log(f"\n完成: 共合并 {len(files)} 个 CSV，输出到 {out_path}")
        for sname, _ in results:
            log(f"  sheet: {sname}")
        if failed:
            print(f"警告: 以下文件合并失败: {', '.join(failed)}", file=sys.stderr)
            sys.exit(1)
    else:
        # 独立模式：每个 CSV 输出为独立的 xlsx
        if len(files) > 1 and args.sheet_name:
            print("警告: --sheet-name 在批量模式下仅对单文件生效，将忽略。", file=sys.stderr)
        if len(files) > 1 and args.output and Path(args.output).suffix.lower() == ".xlsx":
            print(
                "错误: 批量模式下 -o 应指定输出目录，或使用 --combine 合并为一个文件。"
                f"当前输入: {args.output}",
                file=sys.stderr,
            )
            sys.exit(1)

        total = 0
        failed = []
        for f in files:
            try:
                sname, out_path = convert_single_file(
                    f, args.output, args.sheet_name, args.encoding, header
                )
                log(f"  {f.name} [{sname}] -> {out_path}")
                total += 1
            except Exception as e:
                print(f"错误: 转换 {f} 失败: {e}", file=sys.stderr)
                failed.append(str(f))

        if failed:
            print(f"\n完成: 共处理 {total} 个文件，失败 {len(failed)} 个。", file=sys.stderr)
            print(f"失败文件: {', '.join(failed)}", file=sys.stderr)
            sys.exit(1)
        else:
            log(f"\n完成: 共处理 {total} 个文件，生成 {total} 个 xlsx。")


if __name__ == "__main__":
    main()
