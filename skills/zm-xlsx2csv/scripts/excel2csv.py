#!/usr/bin/env python3
"""
Excel 转 CSV 脚本：支持单文件、目录批量、文件列表（空格/逗号/顿号分隔）三种输入模式。
多 sheet 时，每个 sheet 输出为独立 CSV（文件名带 sheet 名）。

输出命名规则（与 SKILL.md / README.md 表格一一对应）：
- 未指定 sheet 且原文件仅 1 个 sheet → 原文件名.csv
- 其他所有情况（多 sheet 或用户显式指定 sheet）→ 原文件名_Sheet名.csv
"""

import argparse
import logging
import os
import re
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError as e:
    print("错误: 缺少 pandas。请安装: pip install pandas openpyxl xlrd", file=sys.stderr)
    sys.exit(1)

try:
    import signal
    HAS_SIGALRM = sys.platform != "win32"
except ImportError:  # pragma: no cover
    HAS_SIGALRM = False


LOG = logging.getLogger("zm_xlsx2csv")


class _ReadTimeout(Exception):
    pass


def _positive_int(v):
    """argparse type: 仅接受正整数；用于 --timeout。"""
    try:
        iv = int(v)
    except (TypeError, ValueError):
        raise argparse.ArgumentTypeError(f"--timeout 必须是正整数，收到: {v!r}")
    if iv <= 0:
        raise argparse.ArgumentTypeError(
            f"--timeout 必须是正整数，收到: {v}；如不需要超时请省略该参数"
        )
    return iv


def _read_excel_with_timeout(input_path, sheet_name, engine, timeout):
    """pd.read_excel 的超时包装。Unix 用 SIGALRM；其他平台无超时。"""
    if not timeout or not HAS_SIGALRM:
        return pd.read_excel(
            input_path, sheet_name=sheet_name, engine=engine,
            header=0, dtype=str,
        )

    def _handler(signum, frame):
        raise _ReadTimeout(f"读取超时（{timeout}s）")

    old = signal.signal(signal.SIGALRM, _handler)
    signal.alarm(int(timeout))
    try:
        return pd.read_excel(
            input_path, sheet_name=sheet_name, engine=engine,
            header=0, dtype=str,
        )
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, old)


EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}


def engine_for(suffix):
    """按扩展名选 pandas engine。xlrd>=2.0 已不支持 .xls，强制提醒。"""
    s = suffix.lower()
    if s == ".xls":
        return "xlrd"
    if s in (".xlsx", ".xlsm"):
        return "openpyxl"
    return None


def split_file_list(text):
    """按空格、逗号、顿号、分号、换行拆分文件列表，去空去重。"""
    parts = re.split(r"[,，、;；\s]+", text.strip())
    return [p for p in parts if p]


def collect_excel_files(input_arg, recursive=False):
    """
    根据输入参数收集所有 Excel 文件路径。
    返回: list of Path（保留顺序，已去重；dedup 用 Path.resolve()）
    """
    arg = input_arg.strip()
    p = Path(arg)

    # 情况1: 目录
    if p.is_dir():
        if recursive:
            raw = [
                f for f in p.rglob("*")
                if f.is_file() and f.suffix.lower() in EXCEL_EXTENSIONS
            ]
        else:
            raw = [
                f for f in p.iterdir()
                if f.is_file() and f.suffix.lower() in EXCEL_EXTENSIONS
            ]
        seen, out = set(), []
        for f in sorted(raw):
            try:
                key = str(f.resolve())
            except OSError:
                key = str(f)
            if key in seen:
                continue
            seen.add(key)
            out.append(f)
        return out

    # 情况2: 单文件直接存在
    if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS:
        return [p]

    # 情况3: 解析为文件列表（空格/逗号/顿号/分号分隔）
    # 提醒：shell 通配符不会被脚本展开；如未展开会落到 invalid 警告中
    if any(ch in arg for ch in ("*", "?", "[")):
        LOG.warning(
            "输入包含 shell 通配符（*/?/[）；脚本不自动展开，请用目录批量或文件列表。"
        )
    candidates = split_file_list(arg)
    seen = set()
    files = []
    missing = []
    for c in candidates:
        cp = Path(c)
        try:
            key = str(cp.resolve())
        except OSError:
            key = str(cp)
        if key in seen:
            continue
        seen.add(key)
        if cp.is_file() and cp.suffix.lower() in EXCEL_EXTENSIONS:
            files.append(cp)
        else:
            missing.append(c)

    if not files and missing:
        LOG.error("无法识别任何有效的 Excel 文件。输入: %s", arg)
        sys.exit(1)

    if missing:
        LOG.warning("跳过无效路径: %s", ", ".join(missing))

    return files


def resolve_output_dir(input_path, output_dir=None):
    """确定输出目录：优先用用户指定的，否则与源文件同目录。"""
    if output_dir:
        d = Path(output_dir)
        try:
            d.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            LOG.error("输出目录创建失败: %s: %s", d, e)
            raise SystemExit(1)
        return d
    return input_path.parent


def _sheet_name_for_filename(sheet_name):
    """将 sheet 名处理为安全的文件名片段。"""
    # 替换文件系统不友好字符
    safe = re.sub(r'[\\/:*?"<>|]', '_', str(sheet_name))
    return safe.strip()


def convert_single_file(input_path, output_dir=None, sheet_name=None, overwrite=False, timeout=None, unique=False):
    """
    转换单个 Excel 文件为 CSV。
    返回: list of (input_sheet_name, output_csv_path)
    """
    input_path = Path(input_path)
    out_dir = resolve_output_dir(input_path, output_dir)
    base_stem = input_path.stem

    # 显式按扩展名选 engine，并对 .xls 给出可读提示
    engine = engine_for(input_path.suffix)
    if engine == "xlrd":
        try:
            import xlrd  # noqa: F401
            if tuple(int(x) for x in xlrd.__VERSION__.split(".")[:2]) >= (2, 0):
                LOG.warning(
                    "%s 是 .xls 格式；当前环境 xlrd>=%s，"
                    "新版 xlrd 已不再支持 .xls 读取，请安装 xlrd<2.0 后重试。",
                    input_path, xlrd.__VERSION__,
                )
        except ImportError:
            LOG.error("读取 %s 需要 xlrd<2.0，但环境未安装 xlrd。", input_path)
            return []

    # 读取所有 sheet 信息（不读数据，仅获取 sheet 名）
    try:
        xl = pd.ExcelFile(str(input_path), engine=engine) if engine else pd.ExcelFile(str(input_path))
    except Exception as e:
        LOG.error("无法打开 %s: %s", input_path, e)
        return []
    sheet_names = xl.sheet_names

    results = []

    if sheet_name is not None:
        # 用户指定了 sheet
        if sheet_name in sheet_names:
            sheets_to_convert = [sheet_name]
        else:
            # 尝试按索引解析
            try:
                idx = int(sheet_name)
                if 0 <= idx < len(sheet_names):
                    sheets_to_convert = [sheet_names[idx]]
                else:
                    LOG.error(
                        "sheet 索引 %s 超出范围 (0-%d)。可用 sheet: %s",
                        idx, len(sheet_names) - 1, sheet_names,
                    )
                    return []
            except ValueError:
                LOG.error("找不到 sheet '%s'。可用 sheet: %s", sheet_name, sheet_names)
                return []
    else:
        sheets_to_convert = sheet_names

    # 命名规则：
    # - 未指定 sheet 且原文件仅 1 个 sheet → 原文件名.csv
    # - 其他所有情况（多 sheet 或用户显式指定 sheet）→ 原文件名_Sheet名.csv
    single_sheet_no_spec = (sheet_name is None and len(sheet_names) == 1)

    # 阶段 1：先收集每个目标 sheet 的目标输出路径（不读不写）
    plan = []
    for sname in sheets_to_convert:
        if single_sheet_no_spec:
            out_name = f"{base_stem}.csv"
        else:
            safe_sname = _sheet_name_for_filename(sname)
            out_name = f"{base_stem}_{safe_sname}.csv"
        out_path = out_dir / out_name
        plan.append((sname, out_path))

    # 阶段 2：先读所有 sheet 的 DataFrame；任何读失败则整体放弃，不写任何 CSV
    read_results = []
    for sname, out_path in plan:
        try:
            df = _read_excel_with_timeout(input_path, sname, engine, timeout)
        except _ReadTimeout as e:
            LOG.error("读取 %s 的 sheet '%s' %s", input_path, sname, e)
            return []
        except Exception as e:
            LOG.error("读取 %s 的 sheet '%s' 失败: %s", input_path, sname, e)
            return []
        read_results.append((sname, df, out_path))

    # 阶段 3：写入（写到临时文件后原子重命名，避免写一半被杀留半截）
    for sname, df, out_path in read_results:
        if out_path.exists() and not overwrite:
            LOG.warning(
                "%s 已存在，跳过（用 --overwrite 强制覆盖；"
                "--unique 需与 --overwrite 一起使用才会自动加 _1/_2 后缀）。",
                out_path,
            )
            continue
        final_out_path = out_path
        if unique and out_path.exists() and overwrite:
            # --unique 优先：生成 <stem>_<n>_<sheet>.csv 直到不存在
            parent = out_path.parent
            suffix = out_path.suffix
            stem_full = out_path.stem  # 已含 _Sheet名（多 sheet 时）
            n = 1
            while True:
                candidate = parent / f"{stem_full}_{n}{suffix}"
                if not candidate.exists():
                    final_out_path = candidate
                    break
                n += 1
                if n > 9999:
                    LOG.error("--unique 后缀超过 9999，跳过 %s", out_path)
                    break
        tmp_path = final_out_path.with_suffix(final_out_path.suffix + ".tmp")
        try:
            df.to_csv(tmp_path, index=False, encoding="utf-8-sig")
            os.replace(tmp_path, final_out_path)
        except Exception as e:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass
            LOG.error("写入 %s 失败: %s", final_out_path, e)
            return []
        results.append((sname, final_out_path))

    return results


def _setup_logging(verbose=False, quiet=False):
    """配置 logging：默认 WARNING；--verbose → INFO；--quiet → ERROR。"""
    if verbose and quiet:
        LOG.warning("--verbose 与 --quiet 互斥；按 --verbose 处理")
        quiet = False
    if verbose:
        level = logging.INFO
    elif quiet:
        level = logging.ERROR
    else:
        level = logging.WARNING
    logging.basicConfig(
        stream=sys.stderr,
        level=level,
        format="%(levelname)s: %(message)s",
        force=True,
    )


def main():
    parser = argparse.ArgumentParser(
        description="将 Excel 文件(.xlsx/.xls/.xlsm)转换为 CSV 文件。"
    )
    parser.add_argument(
        "input",
        help=(
            "输入：单个 Excel 文件路径、包含 Excel 文件的目录路径、"
            "或以空格/逗号/顿号/分号分隔的多个 Excel 文件路径"
        ),
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="输出目录（默认与源文件同目录）",
    )
    parser.add_argument(
        "-s", "--sheet",
        default=None,
        help="指定要转换的 sheet 名或索引（默认转换所有 sheet）",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="已存在同名 CSV 时强制覆盖（默认跳过并警告）",
    )
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="目录批量时递归遍历子目录（默认仅扫顶层）",
    )
    parser.add_argument(
        "--timeout",
        type=_positive_int,
        default=None,
        help="单个 sheet 读取的超时秒数，必须为正整数（默认无超时）",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="严格模式：任何文件失败 → 退出码 1（默认：部分失败仍 exit 0）",
    )
    parser.add_argument(
        "--unique",
        action="store_true",
        help="输出文件名冲突时自动加 _1/_2 后缀（与 --overwrite 不冲突）",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="输出每个 sheet 的读取/写入明细到 stderr（默认仅警告/错误）",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="只输出错误到 stderr（警告也隐藏）",
    )

    args = parser.parse_args()
    _setup_logging(args.verbose, args.quiet)

    files = collect_excel_files(args.input, recursive=args.recursive)

    if not files:
        LOG.error("未找到任何 Excel 文件。")
        sys.exit(1)

    total_files = len(files)
    success_count = 0
    failed_files = []
    total_sheets = 0

    for f in files:
        try:
            results = convert_single_file(
                f, args.output, args.sheet, args.overwrite, args.timeout, args.unique
            )
        except Exception as e:
            LOG.error("转换 %s 失败: %s", f, e)
            failed_files.append(f)
            continue
        if not results:
            failed_files.append(f)
            continue
        success_count += 1
        total_sheets += len(results)
        for sname, out_path in results:
            # 每文件成功明细走 stdout（与 A-2 决策保持一致：操作类信息走 stderr，
            # 实际产出物走 stdout 方便管道 / log 聚合）
            print(f"  {f.name} [{sname}] -> {out_path}")

    LOG.warning(
        "完成: 共处理 %d 个文件，生成 %d 个 CSV；成功 %d 个，失败 %d 个。",
        total_files, total_sheets, success_count, len(failed_files),
    )
    if failed_files:
        for f in failed_files:
            LOG.error("失败: %s", f)
        # 默认容错：部分失败仍 exit 0；仅在零成功或 --strict 时 exit 1
        if args.strict or success_count == 0:
            sys.exit(1)


if __name__ == "__main__":
    main()
