#!/usr/bin/env python3
"""删除 Excel/CSV 中匹配关键词的行。"""

import argparse
import os
import re
import sys

import pandas as pd

# B-C2: 把 CSV 单表在 data_dict 中的哨兵键名抽为模块常量，避免字面量散落
CSV_KEY = "__csv__"

# 用于把 sheet 名中含文件系统非法字符的字符替换为下划线
_SHEET_NAME_INVALID_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _sanitize_sheet_name(sheet_name: str) -> str:
    """将 sheet 名中含文件系统非法字符的字符替换为下划线。"""
    return _SHEET_NAME_INVALID_RE.sub("_", sheet_name)


def parse_args():
    parser = argparse.ArgumentParser(
        description="删除 Excel/CSV 中匹配搜索关键词的行，输出到同目录。"
    )
    parser.add_argument(
        "-f", "--file", required=True, help="输入文件路径（.xlsx/.xlsm/.csv）"
    )
    parser.add_argument(
        "-k",
        "--keyword",
        action="append",
        required=True,
        help="搜索关键词，可多次指定。匹配任意列中包含该关键词的行将被删除。",
    )
    parser.add_argument(
        "--sheet",
        help="指定处理的 sheet 名称或索引（仅 Excel）。不指定则处理所有 sheet。",
    )
    parser.add_argument(
        "--match-mode",
        choices=["contains", "exact"],
        default="contains",
        help="匹配模式：contains（子串匹配，默认）或 exact（精确匹配）。",
    )
    parser.add_argument(
        "--case-sensitive",
        action="store_true",
        help="区分大小写（默认不区分）。",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="自定义输出文件路径。不指定则自动推导为 原文件名_删除多行。",
    )
    parser.add_argument(
        "--format",
        choices=["xlsx", "csv"],
        help="输出格式（默认与输入格式相同）。",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="预览模式：只打印将要删除的行数和行号，不实际生成文件。",
    )
    parser.add_argument(
        "--header-row",
        type=int,
        default=0,
        help="表头所在行（0-based，默认 0 即第 1 行）。"
        "若 xlsx 注释行/合并标题在数据上方，请用此参数指定真正的表头位置。",
    )
    return parser.parse_args()


def read_file(file_path, sheet=None, header_row=0):
    """读取 Excel 或 CSV 文件，返回 (数据字典, 输入格式) 元组。

    ``sheet`` 支持名称或整数索引；索引在多 sheet 工作簿中按 0-based 解析为真实名。
    ``header_row`` 指定表头所在行（0-based），传给 ``pd.read_excel``；CSV 暂不感知。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".xls":
        # .xls 不在支持范围；与同级 skill 口径一致，直接报错退出而不是让 openpyxl 抛 zipfile 错误
        raise ValueError(
            f"不支持的文件格式: {ext}。本 skill 仅支持 .xlsx / .xlsm / .csv；如需处理 .xls，请先用 Excel 另存为 .xlsx"
        )
    if ext in (".xlsx", ".xlsm"):
        if sheet is not None:
            xls = pd.ExcelFile(file_path, engine="openpyxl")
            # 支持整数索引 / 纯数字字符串
            if isinstance(sheet, int) or (
                isinstance(sheet, str) and sheet.isdigit()
            ):
                idx = int(sheet)
                if not (0 <= idx < len(xls.sheet_names)):
                    raise ValueError(
                        f"sheet 索引 {idx} 超出范围（0..{len(xls.sheet_names) - 1}）；"
                        f"可用 sheet: {xls.sheet_names}"
                    )
                real_sheet = xls.sheet_names[idx]
            else:
                if sheet not in xls.sheet_names:
                    raise ValueError(
                        f"未找到 sheet: {sheet}。可用 sheet: {xls.sheet_names}"
                    )
                real_sheet = sheet
            df = pd.read_excel(
                file_path, sheet_name=real_sheet, engine="openpyxl", header=header_row
            )
            return {real_sheet: df}, "xlsx"
        xls = pd.ExcelFile(file_path, engine="openpyxl")
        data = {}
        for sheet_name in xls.sheet_names:
            data[sheet_name] = pd.read_excel(
                file_path, sheet_name=sheet_name, engine="openpyxl", header=header_row
            )
        return data, "xlsx"
    elif ext == ".csv":
        # 编码顺序：utf-8-sig 能解码所有 utf-8（含/不含 BOM）故放最前；gbk/gb18030 覆盖简体中文；
        # utf-8 留作最后兜底（前面都失败时仍可尝试）。gb2312 已被 gbk/gb18030 完全覆盖，从列表中移除。
        encodings = ["utf-8-sig", "gbk", "gb18030", "utf-8"]
        for enc in encodings:
            try:
                df = pd.read_csv(file_path, encoding=enc)
                return {CSV_KEY: df}, "csv"
            except UnicodeDecodeError:
                continue
        raise ValueError(f"无法识别 {file_path} 的编码格式")
    else:
        raise ValueError(f"不支持的文件格式: {ext}")


def should_delete_row(row, keywords, match_mode, case_sensitive):
    """判断一行是否匹配任意关键词。"""
    flags = 0 if case_sensitive else re.IGNORECASE

    for value in row:
        if pd.isna(value):
            continue
        text = str(value)
        for kw in keywords:
            if match_mode == "exact":
                if not case_sensitive:
                    if text.lower() == kw.lower():
                        return True
                else:
                    if text == kw:
                        return True
            else:
                if re.search(re.escape(kw), text, flags):
                    return True
    return False


def delete_rows(df, keywords, match_mode, case_sensitive):
    """删除匹配关键词的行，返回 (过滤后的 DataFrame, 删除的行号列表, 删除行数)。"""
    mask = df.apply(
        lambda row: should_delete_row(row, keywords, match_mode, case_sensitive), axis=1
    )
    deleted_indices = df[mask].index.tolist()
    result = df[~mask].reset_index(drop=True)
    return result, deleted_indices, len(deleted_indices)


def derive_output_path(input_path, output_arg, fmt, sheet_name=None):
    """推导输出文件路径。

    当 ``output_arg`` 显式提供时，禁止与 ``input_path`` 指向同一文件，
    否则会覆盖源文件（与"源文件只读"核心原则冲突）；其他场景沿用自动命名 + 序号递增。
    ``sheet_name`` 为 None 时不追加 ``_<sheet名>`` 后缀；非空时含特殊字符的 sheet 名
    会被替换为下划线，避免生成无法创建的文件名。
    """
    if output_arg:
        if os.path.abspath(output_arg) == os.path.abspath(input_path):
            raise ValueError(
                f"--output 不得与 -f 指向同一文件（会覆盖源文件）：{input_path}。"
                "请省略 -o 让脚本自动命名（如 <原文件名>_删除多行.<ext>）。"
            )
        return output_arg

    base, ext = os.path.splitext(input_path)
    if fmt == "csv":
        ext = ".csv"
    elif fmt == "xlsx":
        ext = ".xlsx"

    suffix = "_删除多行"
    if sheet_name and sheet_name != CSV_KEY:
        # 含特殊字符的 sheet 名会被替换为下划线，避免生成无法创建的文件名
        suffix += f"_{_sanitize_sheet_name(sheet_name)}"

    out_path = f"{base}{suffix}{ext}"
    counter = 1
    while os.path.exists(out_path):
        out_path = f"{base}{suffix}_{counter}{ext}"
        counter += 1
    return out_path


def write_file(df, output_path, fmt, sheet_name=None):
    """将 DataFrame 写入文件。"""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    if fmt == "csv":
        df.to_csv(output_path, index=False, encoding="utf-8-sig")
    elif fmt == "xlsx":
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            sheet = sheet_name if sheet_name and sheet_name != CSV_KEY else "Sheet1"
            df.to_excel(writer, sheet_name=sheet, index=False)
    else:
        raise ValueError(f"不支持的输出格式: {fmt}")


def main():
    args = parse_args()
    input_path = args.file

    # 过滤空关键词：空字符串 -k "" 会被 re.search 当作"任意位置匹配"，会误删整张表
    args.keyword = [k for k in args.keyword if k]
    if not args.keyword:
        print("错误: --keyword 不能为空字符串", file=sys.stderr)
        sys.exit(1)

    if not os.path.exists(input_path):
        print(f"错误: 文件不存在: {input_path}", file=sys.stderr)
        sys.exit(1)

    # A3-P0-2: --output 显式与 -f 指向同一文件的保护提前到主循环之前，
    # 不进入任何行处理、不打印任何删除统计，UX 与多 sheet + --output 防护对称
    if args.output and os.path.abspath(args.output) == os.path.abspath(input_path):
        print(
            f"错误: --output 不得与 -f 指向同一文件（会覆盖源文件）：{input_path}。"
            "请省略 -o 让脚本自动命名（如 <原文件名>_删除多行.<ext>）。",
            file=sys.stderr,
        )
        sys.exit(1)

    try:
        data_dict, input_fmt = read_file(input_path, args.sheet, args.header_row)
    except Exception as e:
        print(f"读取文件失败: {e}", file=sys.stderr)
        sys.exit(1)

    # 多 sheet Excel + --output 显式路径会互相覆盖（多 sheet 都写到同一文件），
    # 提前报错退出比静默丢数据更安全
    multi_sheet = len(data_dict) > 1 and CSV_KEY not in data_dict
    if args.output and multi_sheet:
        print(
            "错误: 多 sheet Excel 不支持 --output 显式路径（每个 sheet 会覆盖同一文件）。"
            "请省略 -o 让脚本为每个 sheet 自动命名（如 <原文件名>_删除多行_<sheet名>.<ext>）。",
            file=sys.stderr,
        )
        sys.exit(1)

    # A3-P0-1 / A3-P1-1: 仅在"显式 --sheet"或"多 sheet"时按 sheet 名追加后缀。
    # 单 sheet xlsx + 不指定 --sheet 时不追加（与 SKILL.md / README / agents/openai.yaml 契约一致）。
    explicit_sheet = args.sheet is not None
    append_sheet = explicit_sheet or multi_sheet

    output_fmt = args.format or input_fmt
    total_deleted = 0
    # A3-P1-6: 写失败时清理已成功写入的 sheet 文件，避免半成品残留
    written_paths = []

    for sheet_name, df in data_dict.items():
        if df.empty:
            print(f"[{sheet_name}] 数据为空，跳过。")
            continue

        result, deleted_indices, deleted_count = delete_rows(
            df, args.keyword, args.match_mode, args.case_sensitive
        )

        display_name = sheet_name if sheet_name != CSV_KEY else "CSV"
        print(f"\n[{display_name}] 原始行数: {len(df)}, 删除行数: {deleted_count}, 剩余行数: {len(result)}")

        if deleted_count > 0:
            # Excel 物理行号 = DataFrame index + header_row + 2
            # （header_row=0 时表头=1，DataFrame index 0 对应物理行 2；
            #  header_row=N 表示前 N 行是注释/合并标题）
            physical_rows = [i + args.header_row + 2 for i in deleted_indices]
            print(f"  删除行号（1-based 物理行号，含表头；表头=1，数据行=2..n+1）: {physical_rows}")
        else:
            print(f"  未匹配到任何包含关键词的行。")

        if args.dry_run:
            total_deleted += deleted_count
            continue

        # A3-P0-1 / A3-P1-1: 单 sheet xlsx + 不指定 --sheet 时不追加 sheet 名后缀
        if sheet_name == CSV_KEY or not append_sheet:
            output_path = derive_output_path(input_path, args.output, output_fmt)
        else:
            output_path = derive_output_path(
                input_path, args.output, output_fmt, sheet_name=sheet_name
            )

        try:
            write_file(result, output_path, output_fmt, sheet_name)
            written_paths.append(output_path)
            print(f"  已保存: {output_path}")
        except ValueError as e:
            # derive_output_path 显式 --output 防护（纵深防御）
            _cleanup_written(written_paths)
            print(f"  错误: {e}", file=sys.stderr)
            sys.exit(1)
        except Exception as e:
            _cleanup_written(written_paths)
            print(f"  保存失败: {e}", file=sys.stderr)
            sys.exit(1)

        total_deleted += deleted_count

    print(f"\n总计删除: {total_deleted} 行")
    if args.dry_run:
        print("(预览模式，未实际生成文件)")


def _cleanup_written(paths):
    """删除已成功写入的输出文件（用于多 sheet 写失败时回滚半成品）。"""
    for p in paths:
        try:
            os.remove(p)
        except OSError:
            pass


if __name__ == "__main__":
    main()
