#!/usr/bin/env python3
"""删除 Excel/CSV 中匹配关键词的行。"""

import argparse
import os
import re
import sys
import zipfile

import numpy as np
import pandas as pd
import openpyxl.utils.exceptions as _openpyxl_exc

# B-C2: 把 CSV 单表在 data_dict 中的哨兵键名抽为模块常量，避免字面量散落
CSV_KEY = "__csv__"

# A3-P2-1: 把 -o / -f 同源的错误消息抽为模块常量；main() 与 derive_output_path() 共用
_ERR_OUTPUT_SAME_AS_INPUT = (
    "--output 不得与 -f 指向同一文件（会覆盖源文件）：{input_path}。"
    "请省略 -o 让脚本自动命名（如 <原文件名>_删除多行.<ext>）。"
)

# 用于把 sheet 名中含文件系统非法字符与 ASCII 控制字符的字符替换为下划线
_SHEET_NAME_INVALID_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# B7-P2-2: CSV 编码探测顺序抽为模块常量；utf-8-sig 放最前（覆盖所有 utf-8 含/不含 BOM），
# gbk/gb18030 覆盖简体中文，utf-8 留作最后兜底
_CSV_ENCODING_PROBE_ORDER = ("utf-8-sig", "gbk", "gb18030", "utf-8")


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
        # B5-P1-6: xlsx 损坏错误分类。把 zipfile / openpyxl 已知异常翻译为友好提示，
        # 区分"xlsx 损坏"与"通用读取失败"两类场景
        try:
            xls = pd.ExcelFile(file_path, engine="openpyxl")
        except (zipfile.BadZipFile, _openpyxl_exc.InvalidFileException) as e:
            raise ValueError(
                f"xlsx 损坏或不是有效的 xlsx 文件: {e}"
            ) from e
        if sheet is not None:
            sheet_names = xls.sheet_names
            # 1) 优先按精确名匹配（覆盖 "0" / "01" / "2025" 等纯数字 sheet 名场景，A1-P0-1）
            if sheet in sheet_names:
                real_sheet = sheet
            # 2) 未命中且字符串全为数字字符时回退为整数索引
            elif sheet.isdigit():
                idx = int(sheet)
                if not (0 <= idx < len(sheet_names)):
                    raise ValueError(
                        f"sheet 索引 {idx} 超出范围（0..{len(sheet_names) - 1}）；"
                        f"可用 sheet: {sheet_names}"
                    )
                real_sheet = sheet_names[idx]
            # B5-P1-3: 区分"非法索引格式"（如 0.5 / -1 / + / 1.0）与"未找到名字"，
            # 给用户更明确的诊断信息
            elif any(ch in sheet for ch in (".", "-", "+", "e", "E")):
                raise ValueError(
                    f"sheet 索引 {sheet!r} 不是有效整数（应为 0..{len(sheet_names) - 1} 的正整数）；"
                    f"可用 sheet: {sheet_names}"
                )
            else:
                raise ValueError(
                    f"未找到 sheet: {sheet!r}。可用 sheet: {sheet_names}"
                )
            df = pd.read_excel(
                file_path, sheet_name=real_sheet, engine="openpyxl", header=header_row
            )
            return {real_sheet: df}, "xlsx"
        data = {}
        for sheet_name in xls.sheet_names:
            data[sheet_name] = pd.read_excel(
                file_path, sheet_name=sheet_name, engine="openpyxl", header=header_row
            )
        return data, "xlsx"
    elif ext == ".csv":
        # A1-P1-2: --header-row 仅对 xlsx 生效；CSV 暂不感知（pd.read_csv 默认 header=0）。
        # 当用户显式传非零值时打一次 stderr 警告，避免静默忽略
        if header_row != 0:
            print(
                f"警告: --header-row={header_row} 仅对 xlsx/.xlsm 生效；CSV 已忽略，仍按第 1 行当表头处理。",
                file=sys.stderr,
            )
        # 编码顺序：utf-8-sig 能解码所有 utf-8（含/不含 BOM）故放最前；gbk/gb18030 覆盖简体中文；
        # utf-8 留作最后兜底（前面都失败时仍可尝试）。gb2312 已被 gbk/gb18030 完全覆盖，从列表中移除。
        for enc in _CSV_ENCODING_PROBE_ORDER:
            try:
                df = pd.read_csv(file_path, encoding=enc)
                return {CSV_KEY: df}, "csv"
            except UnicodeDecodeError:
                continue
        raise ValueError(f"无法识别 {file_path} 的编码格式")
    else:
        raise ValueError(f"不支持的文件格式: {ext}")


def delete_rows(df, keywords, match_mode, case_sensitive):
    """删除匹配关键词的行，返回 (过滤后的 DataFrame, 删除的行号列表, 删除行数)。

    A1-P1-1: 改向量化实现——按列用 str.contains / == 命中，再用 logical_or 在关键词间合并，
    最后 any(axis=1) 跨列合并。原 df.apply(lambda row: ...) 的 row-wise Python 循环在大表下
    不可用，10w+ 行耗时从分钟级降到秒级。
    """
    # 字符串化一次：astype(str) 让 int / float / Timestamp 等也能参与匹配；
    # A2-P0-1: 但 astype(str) 会把 NaN 变成字符串 "nan"，导致关键词 "nan"（或任何恰好
    # 等于 NaN astype(str) 后字符串值的子串）误删含 NaN 的行。用 .where(notna(), np.nan)
    # 把 NaN 还原，让 contains (na=False) / eq (NaN 比较) 自动屏蔽
    str_df = df.astype(str).where(df.notna(), other=np.nan)
    mask = pd.Series(False, index=df.index)
    for kw in keywords:
        if match_mode == "exact":
            if case_sensitive:
                kw_mask = str_df.eq(kw)
            else:
                # 大小写不敏感：str.lower() 后比较
                kw_mask = str_df.apply(lambda c: c.str.lower()).eq(kw.lower())
        else:
            # contains: re.escape 避免关键词里的正则元字符误匹配
            escaped = re.escape(kw)
            kw_mask = str_df.apply(
                lambda c: c.str.contains(escaped, case=case_sensitive, na=False, regex=True)
            )
        # 行级合并（任一列命中即真）
        kw_mask = kw_mask.any(axis=1)
        # 关键词间 OR
        mask = mask | kw_mask
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
            raise ValueError(_ERR_OUTPUT_SAME_AS_INPUT.format(input_path=input_path))
        # A1-P0-2: 校验 -o 路径扩展名与 --format 一致；不一致时让 main 报明确错误退出，
        # 避免被通用 except 兜成笼统的"保存失败: InvalidFileException"
        out_ext = os.path.splitext(output_arg)[1].lower()
        expected_ext = ".csv" if fmt == "csv" else ".xlsx"
        if out_ext and out_ext != expected_ext:
            raise ValueError(
                f"--output 路径扩展名 {out_ext!r} 与 --format {fmt!r} 不一致；"
                f"输出 {fmt} 时扩展名应为 {expected_ext!r}。请调整 -o 路径或省略 -o 让脚本自动命名。"
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

    # A3-P0-2: 同时过滤空字符串与纯空白关键词（半角空格 / 全角空格 / Tab / NBSP / 换行等）
    # 空字符串 -k "" 会被 re.search 当作"任意位置匹配"，会误删整张表；
    # 纯空白关键词也会被静默接受并实际删除含空格的行
    args.keyword = [k for k in args.keyword if k and not k.isspace()]
    if not args.keyword:
        print("错误: --keyword 不能为空字符串或纯空白", file=sys.stderr)
        sys.exit(1)

    # A2-P0-3: 重复 -k 去重（保留首次出现顺序）；避免 -k foo -k foo -k bar 时
    # 重复计算 mask，浪费内存与时间
    args.keyword = list(dict.fromkeys(args.keyword))

    if not os.path.exists(input_path):
        print(f"错误: 文件不存在: {input_path}", file=sys.stderr)
        sys.exit(1)

    # A3-P0-2: --output 显式与 -f 指向同一文件的保护提前到主循环之前，
    # 不进入任何行处理、不打印任何删除统计，UX 与多 sheet + --output 防护对称
    if args.output and os.path.abspath(args.output) == os.path.abspath(input_path):
        print(
            f"错误: {_ERR_OUTPUT_SAME_AS_INPUT.format(input_path=input_path)}",
            file=sys.stderr,
        )
        sys.exit(1)

    # A3-P1-4: --output 显式路径不得落在系统敏感目录（/etc、/usr、/var、/bin、/sbin、/boot）
    # os.makedirs(..., exist_ok=True) 缺护栏会越权创建；与"源文件只读"对称性补齐
    if args.output:
        _abspath = os.path.abspath(args.output)
        _forbidden_roots = ("/etc", "/usr", "/var", "/bin", "/sbin", "/boot", "/sys", "/proc", "/dev", "/root")
        for _root in _forbidden_roots:
            if _abspath == _root or _abspath.startswith(_root + "/"):
                print(
                    f"错误: --output 路径不允许落在系统目录 {_root!r}：{_abspath}。"
                    "请将输出文件放到用户工作目录（如 /tmp 或当前目录）下。",
                    file=sys.stderr,
                )
                sys.exit(1)
        # B3-P1-1: abspath 不解析 symlink。若 /tmp 是 → /etc 的 symlink，abspath 检查
        # 通过但实际写入 /etc/foo。realpath 二次校验捕获；Windows 下 realpath 行为不同
        # （不解析 NTFS junction 的目标），降级为 warning
        try:
            _realpath = os.path.realpath(_abspath)
            for _root in _forbidden_roots:
                if _realpath == _root or _realpath.startswith(_root + "/"):
                    if sys.platform.startswith("win"):
                        print(
                            f"警告: --output 路径 {_abspath} 经 realpath 解析为 {_realpath}，"
                            f"落在系统目录 {_root!r}。Windows 平台 symlink/junction 行为可能未完全解析，"
                            "请人工确认输出位置。",
                            file=sys.stderr,
                        )
                    else:
                        print(
                            f"错误: --output 路径 {_abspath} 经 realpath 解析为 {_realpath}，"
                            f"落在系统目录 {_root!r}。请将输出文件放到用户工作目录（如 /tmp 或当前目录）下。",
                            file=sys.stderr,
                        )
                        sys.exit(1)
        except OSError as e:
            print(f"警告: 无法解析 --output 路径的 realpath（{e}）；按 abspath 校验结果继续。", file=sys.stderr)
        # B4-P1-1: --output 显式路径若已存在文件则警告（非阻断）。
        # 显式 -o 仍允许覆盖（用户主动传 -o 是显式意图），但 stderr 警告让用户知情
        if os.path.isfile(args.output):
            print(
                f"警告: 目标文件 {args.output} 已存在，将被覆盖。"
                "如需保留请省略 -o 让脚本自动命名（自动追加序号 _1 / _2 ...）。",
                file=sys.stderr,
            )

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

    # B0-P1-1: 扩展名校验提前到主循环之前（与同源 / 系统目录 / 多 sheet + -o
    # 防护对称）。单 sheet 路径在 delete_rows 之前就报格式错误，UX 改为
    # "先报错误 → 不进入任何行处理 → 不打印任何删除统计"
    if args.output:
        try:
            derive_output_path(input_path, args.output, args.format or input_fmt)
        except ValueError as e:
            print(f"  错误: {e}", file=sys.stderr)
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
            # A2-P1-2: CSV 模式下 --header-row 已被忽略（CSV 始终按第 1 行当表头），
            # 物理行号按 i + 2 计算；xlsx 模式仍按 i + header_row + 2 计算。
            if sheet_name == CSV_KEY:
                physical_rows = [i + 2 for i in deleted_indices]
                physical_note = "（1-based 物理行号，含表头；表头=1，数据行=2..n+1；CSV 忽略 --header-row）"
            else:
                # Excel 物理行号 = DataFrame index + header_row + 2
                # （header_row=0 时表头=1，DataFrame index 0 对应物理行 2；
                #  header_row=N 表示前 N 行是注释/合并标题）
                physical_rows = [i + args.header_row + 2 for i in deleted_indices]
                physical_note = "（1-based 物理行号，含表头；表头=1，数据行=2..n+1）"
            print(f"  删除行号{physical_note}: {physical_rows}")
        else:
            print(f"  未匹配到任何包含关键词的行。")

        if args.dry_run:
            # A2-P1-3: dry-run 模式跳过 derive_output_path（避免 -o 扩展名校验误报，
            # 因为 dry-run 不会写文件，扩展名不一致不影响预览结果）
            total_deleted += deleted_count
            continue

        # A3-P0-1 / A3-P1-1: 单 sheet xlsx + 不指定 --sheet 时不追加 sheet 名后缀
        try:
            if sheet_name == CSV_KEY or not append_sheet:
                output_path = derive_output_path(input_path, args.output, output_fmt)
            else:
                output_path = derive_output_path(
                    input_path, args.output, output_fmt, sheet_name=sheet_name
                )
        except ValueError as e:
            # A1-P0-2: derive_output_path 的 ValueError（同源 / 扩展名不一致）走与写失败相同的
            # 友好提示路径，不暴露 traceback
            _removed, _failed = _cleanup_written(written_paths)
            _print_cleanup_failures(_failed)
            print(f"  错误: {e}", file=sys.stderr)
            sys.exit(1)

        try:
            write_file(result, output_path, output_fmt, sheet_name)
            written_paths.append(output_path)
            print(f"  已保存: {output_path}")
        except ValueError as e:
            # derive_output_path 显式 --output 防护（纵深防御）
            _removed, _failed = _cleanup_written(written_paths)
            _print_cleanup_failures(_failed)
            print(f"  错误: {e}", file=sys.stderr)
            sys.exit(1)
        except Exception as e:
            _removed, _failed = _cleanup_written(written_paths)
            _print_cleanup_failures(_failed)
            print(f"  保存失败: {e}", file=sys.stderr)
            sys.exit(1)

        total_deleted += deleted_count

    print(f"\n总计删除: {total_deleted} 行")
    if args.dry_run:
        print("(预览模式，未实际生成文件)")


def _cleanup_written(paths):
    """删除已成功写入的输出文件（用于多 sheet 写失败时回滚半成品）。

    A2-P0-4: cleanup 自身失败时不再静默吞错；返回 (removed, failed) 列表，
    调用方负责把 failed 写到 stderr，让用户知道哪些残留需要手动清理。
    """
    removed = []
    failed = []
    for p in paths:
        try:
            os.remove(p)
            removed.append(p)
        except OSError as e:
            failed.append((p, str(e)))
    return removed, failed


def _print_cleanup_failures(failed):
    """A2-P0-4: 把 _cleanup_written 残留的失败路径打印到 stderr。

    failed 是 [(path, error_str), ...] 列表；调用方在原始错误信息之后调用。
    """
    if not failed:
        return
    print("  警告: 以下文件已写入但清理失败（请手动删除）:", file=sys.stderr)
    for p, err in failed:
        print(f"    - {p}: {err}", file=sys.stderr)


if __name__ == "__main__":
    main()
