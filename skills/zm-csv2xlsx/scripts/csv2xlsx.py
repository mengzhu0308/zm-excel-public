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
except ImportError:
    print(
        "错误: 缺少 pandas。请在 agent-skills 环境中安装: "
        "conda run -n agent-skills pip install pandas openpyxl",
        file=sys.stderr,
    )
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print(
        "错误: 缺少 openpyxl。请在 agent-skills 环境中安装: "
        "conda run -n agent-skills pip install openpyxl",
        file=sys.stderr,
    )
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

    优先按 BOM 字节检测（utf-8-sig / utf-16 / utf-32）；BOM 命中后用该编码
    试读前 max_lines 行验证；验证失败时退回到 utf-8 → gb18030 顺序试编码，
    避免大文件全量读取。
    """
    if max_lines is None or max_lines <= 0:
        max_lines = 100
    # BOM 字节级快速检测（不依赖行数启发式）
    bom_candidates = []
    try:
        with open(file_path, "rb") as f:
            head = f.read(4)
        if head.startswith(b"\xef\xbb\xbf"):
            bom_candidates.append("utf-8-sig")
        if head.startswith(b"\xff\xfe\x00\x00") or head.startswith(b"\x00\x00\xfe\xff"):
            bom_candidates.append("utf-32")
        if head.startswith(b"\xff\xfe") or head.startswith(b"\xfe\xff"):
            bom_candidates.append("utf-16")
    except OSError:
        # 文件无法以二进制读取时退回到文本试编码
        pass
    # BOM 命中后必须用该编码试读验证；验证失败则继续走文本试编码兜底
    for enc in bom_candidates:
        try:
            with open(file_path, "r", encoding=enc) as f:
                for _ in range(max_lines):
                    if not f.readline():
                        break
            return enc
        except UnicodeDecodeError:
            continue
    # 优先级: utf-8 > gb18030
    encodings = ["utf-8", "gb18030"]
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
    arg = (input_arg or "").strip()
    if not arg:
        print("错误: 必须提供输入路径（单个 CSV、目录或文件列表）。", file=sys.stderr)
        sys.exit(1)
    p = Path(arg)

    # 仅当输入"看起来像单文件/单目录"（不含列表分隔符）且不存在时，
    # 给出友好错误；含分隔符的输入直接走文件列表解析分支
    looks_like_list = any(sep in arg for sep in (",", "，", "、", ";", "；", "\n"))
    if not looks_like_list and not p.exists():
        print(f"错误: 输入路径不存在: {p}", file=sys.stderr)
        sys.exit(1)

    # 情况1: 目录
    if p.is_dir():
        try:
            files = sorted(
                [f for f in p.iterdir() if f.is_file() and f.suffix.lower() == ".csv"]
            )
        except FileNotFoundError:
            print(f"错误: 目录不存在: {p}", file=sys.stderr)
            sys.exit(1)
        except PermissionError:
            print(f"错误: 无权限访问目录: {p}", file=sys.stderr)
            sys.exit(1)
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
    """根据输入和输出参数解析最终输出路径。

    分支判定优先级：先按"以 / 或 \\ 结尾"或"是已存在目录"判定为目录；
    否则按".xlsx 后缀"判定为文件；无后缀且不是目录时，按文件名处理
    （自动追加 .xlsx），并通过 stderr 提示该行为由用户预期。
    """
    if not output_arg:
        return input_path.with_suffix(".xlsx")

    out_path = Path(output_arg)
    raw = str(out_path)
    is_dir_hint = raw.endswith(("/", "\\")) or out_path.is_dir()
    if is_dir_hint:
        out_path.mkdir(parents=True, exist_ok=True)
        return out_path / f"{input_path.stem}.xlsx"
    if out_path.suffix.lower() == ".xlsx":
        out_path.parent.mkdir(parents=True, exist_ok=True)
        return out_path
    # 无后缀路径（非目录）：打印提示后按文件名处理，自动追加 .xlsx
    print(
        f"提示: -o '{output_arg}' 不是目录也不是 .xlsx 路径，"
        "按文件名处理并自动追加 .xlsx 后缀。",
        file=sys.stderr,
    )
    out_path = out_path.with_suffix(".xlsx")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    return out_path


def convert_single_file(input_path, output_dir=None, sheet_name=None, encoding=None, header=None, detect_lines=100):
    """
    转换单个 CSV 文件为 xlsx。
    返回: (sheet_name_used, output_xlsx_path)
    """
    input_path = Path(input_path)
    if encoding:
        enc = encoding
    else:
        try:
            enc = detect_encoding(input_path, max_lines=detect_lines)
        except ValueError as e:
            # 编码自动检测失败：附"建议使用 -e"提示后重新抛出，由调用方处理
            raise ValueError(f"{e} 建议使用 -e 手动指定编码后重新运行。") from e
    df = pd.read_csv(input_path, encoding=enc, header=header)

    out_path = _resolve_output_path(input_path, output_dir)

    sname = _sheet_name_for_filename(sheet_name) if sheet_name else _sheet_name_for_filename(input_path.stem)
    df.to_excel(out_path, sheet_name=sname, index=False, engine="openpyxl")

    return (sname, out_path)


def convert_combine(files, output_path, encoding=None, header=None, detect_lines=100):
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
                    if encoding:
                        enc = encoding
                    else:
                        try:
                            enc = detect_encoding(f, max_lines=detect_lines)
                        except ValueError as e:
                            # 编码自动检测失败：附"建议使用 -e"提示后转为普通失败
                            print(
                                f"错误: 合并时读取 {f} 失败: {e} 建议使用 -e 手动指定编码。",
                                file=sys.stderr,
                            )
                            failed.append(str(f))
                            continue
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
    except BaseException:
        # 任何异常（含 IndexError / PermissionError / OSError 等）下，
        # 主动清理残留 xlsx，避免留下损坏/空文件
        if out_path.exists():
            try:
                out_path.unlink()
            except OSError:
                pass
        raise

    # 如果没有任何成功结果，删除空输出文件
    if not results and out_path.exists():
        out_path.unlink()

    return results, failed


def _enforce_project_root(args, project_root, files=None, parser=None):
    """校验 args.input / args.output / files 是否都在 project_root 之下。

    越界或路径无法解析则通过 parser.error 退出（退出码 2）。仅在显式传入
    --project-root 时启用。files 优先：调用 collect_csv_files 后传入；files
    缺失时回退到 args.input 字符串（仅作单文件/单目录假设校验，不再支持
    文件列表字符串）。

    project_root 必须是绝对路径（main() 入口处已对原始 args.project_root
    字符串做 is_absolute() 校验，此处不再重复）。parser 必须传入：与 main()
    入口处使用同一 parser 以保证退出码与错误格式一致。
    """
    if parser is None:
        raise ValueError("_enforce_project_root 必须显式传入 parser 引用。")
    candidates = []
    if files:
        for f in files:
            candidates.append(("input", str(f)))
    elif args.input:
        candidates.append(("input", args.input))
    if args.output:
        candidates.append(("output", args.output))

    for label, raw in candidates:
        try:
            resolved = Path(raw).expanduser().resolve()
        except (OSError, RuntimeError) as e:
            # 路径无法解析时按 fail-closed 退出，而非静默放过；
            # 越界保护在解析失败场景下应拒绝继续，避免越界写入。
            parser.error(
                f"错误: --{label} 路径 '{raw}' 无法解析为有效路径以校验"
                f" --project-root '{project_root}': {e}"
            )
        try:
            resolved.relative_to(project_root)
        except ValueError:
            parser.error(
                f"错误: --{label} 路径 '{raw}' 越出 --project-root '{project_root}'。"
            )


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
        "--project-root",
        default=None,
        help=(
            "可选：项目根目录约束（必须是绝对路径）。启用时，所有输入/输出路径"
            "必须在该目录之下，越界则报错并退出。默认关闭。"
        ),
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
    parser.add_argument(
        "--force",
        action="store_true",
        help="强制覆盖已存在的输出文件（默认遇到已存在的输出文件报错退出）",
    )
    parser.add_argument(
        "--encoding-detect-lines",
        type=int,
        default=100,
        help=(
            "编码自动检测时读取的最大行数（用于 BOM 验证与试编码 fallback）；"
            "默认 100，范围 1-10000"
        ),
    )

    args = parser.parse_args()

    def log(msg):
        if not args.quiet:
            print(msg)

    # 校验 --encoding-detect-lines 范围
    if args.encoding_detect_lines < 1 or args.encoding_detect_lines > 10000:
        print(
            f"错误: --encoding-detect-lines 必须在 1-10000 之间（当前: {args.encoding_detect_lines}）。",
            file=sys.stderr,
        )
        sys.exit(1)

    files = collect_csv_files(args.input)

    if args.project_root is not None:
        # 先校验 --project-root 必须是绝对路径（避免 . / ./data 在不同 cwd 下被 resolve 到不同目录）
        if not Path(args.project_root).is_absolute():
            parser.error(
                f"错误: --project-root 必须是绝对路径（当前: '{args.project_root}'）。"
            )
        _enforce_project_root(
            args, Path(args.project_root).resolve(), files=files, parser=parser
        )

    if not files:
        print("未找到任何 CSV 文件。", file=sys.stderr)
        sys.exit(1)

    header = None if args.no_header else 0

    # 解析最终输出路径；若已存在且未传 --force 则报错
    if args.combine:
        if not args.output:
            print("错误: --combine 模式需要配合 -o 指定输出文件路径。", file=sys.stderr)
            sys.exit(1)
        if args.sheet_name:
            print("警告: --sheet-name 在合并模式下无效，将使用 CSV 文件名作为 sheet 名。", file=sys.stderr)
        # 用 _resolve_output_path 统一处理"目录/.xlsx/无后缀"三种分支，
        # 这样合并模式下 -o combined（无后缀）也能打印 stderr 提示
        anchor_input = files[0] if files else Path("output")
        out_path = _resolve_output_path(anchor_input, args.output)
        if out_path.exists() and not args.force:
            print(
                f"错误: 输出文件已存在 '{out_path}'。使用 --force 覆盖。",
                file=sys.stderr,
            )
            sys.exit(1)
    else:
        # 独立模式：对每个目标 xlsx 做覆盖校验（单文件场景只有一个目标；多文件时先解析目录或单文件路径）
        if len(files) == 1 and args.output:
            # 单文件 + 指定 -o：解析后做覆盖校验
            _probe_path = _resolve_output_path(files[0], args.output)
            if _probe_path.exists() and not args.force:
                print(
                    f"错误: 输出文件已存在 '{_probe_path}'。使用 --force 覆盖。",
                    file=sys.stderr,
                )
                sys.exit(1)

    if args.combine:
        # 合并模式：所有 CSV 写入一个 xlsx
        results, failed = convert_combine(files, out_path, args.encoding, header, detect_lines=args.encoding_detect_lines)
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
                    f, args.output, args.sheet_name, args.encoding, header,
                    detect_lines=args.encoding_detect_lines,
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
    try:
        main()
    except ValueError as e:
        # detect_encoding 抛出的"无法自动检测编码"或类似 ValueError
        # 在 main() 顶层捕获后转为对用户更友好的提示，建议使用 -e 手动指定编码
        print(
            f"错误: {e}\n提示: 请使用 -e 参数手动指定编码后重新运行。",
            file=sys.stderr,
        )
        sys.exit(1)
