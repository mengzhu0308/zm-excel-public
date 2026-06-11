#!/usr/bin/env python3
"""
Excel/CSV 合并脚本：
- 支持目录扫描或多文件路径输入
- 字段兼容性分析，按列名相似度自动分组
- --preview 生成合并清单，--plan 按清单执行合并
- 默认模式下自动检测字段差异并智能分组输出
"""

import argparse
import fnmatch
import logging
import os
import re
import sys
import warnings
from pathlib import Path

try:
    import pandas as pd
except ImportError as e:
    print("错误: 缺少 pandas。请安装: pip install pandas openpyxl xlrd", file=sys.stderr)
    sys.exit(1)


# ---------- 模块级常量（单一真相来源，避免脚本/文档/SKILL 漂移） ----------
DEFAULT_PATTERN = "*.xlsx,*.xls,*.xlsm,*.csv"
DEFAULT_SOURCE_COL = "来源文件"
DEFAULT_SIMILARITY_THRESHOLD = 0.8
MAX_SHEET_NAME_LEN = 31
# Excel sheet 名非法字符：\ / ? * : [ ]
EXCEL_INVALID_SHEET_CHARS = re.compile(r"[\\/*?:[\]]")


# ---------- logger（默认 INFO；可由 --log-level 调整；按 A-2 P1-6 引入） ----------
_LOGGER_NAME = "zm_excels_merge"
logger = logging.getLogger(_LOGGER_NAME)
if not logger.handlers:
    _h = logging.StreamHandler(sys.stderr)
    _h.setFormatter(logging.Formatter("  %(levelname)s: %(message)s"))
    logger.addHandler(_h)


def _set_log_level(level_name: str) -> None:
    """按 argparse choices 字符串设置 logger 级别。"""
    level = getattr(logging, level_name.upper(), None)
    if not isinstance(level, int):
        raise ValueError(f"未知日志级别: {level_name}")
    logger.setLevel(level)


def parse_file_paths(paths_str):
    """解析用户输入的多文件路径字符串，支持逗号、空格、中文逗号分隔。"""
    if not paths_str:
        return []
    # 统一替换中文逗号、多个空格为英文逗号
    normalized = re.sub(r'[\s，、]+', ',', paths_str.strip())
    raw = [p.strip() for p in normalized.split(',') if p.strip()]
    result = []
    for p in raw:
        path = Path(p)
        if path.exists():
            result.append(path)
        else:
            print(f"  警告: 路径不存在，跳过: {p}", file=sys.stderr)
    return result


def discover_files(directory, patterns, recursive=False, exclude_paths=None):
    """扫描目录下匹配的文件，返回 Path 列表。

    如果传入的 directory 实际上是一个文件路径（而非目录），
    直接返回该文件，避免误将单文件输入扩展为目录扫描。
    """
    directory = Path(directory).resolve()
    if not directory.exists():
        print(f"  警告: 路径不存在: {directory}", file=sys.stderr)
        return []

    # 单文件保护：传入的是文件而非目录时，直接返回该文件
    if directory.is_file():
        return [directory]
    files = []
    pattern_list = [p.strip() for p in patterns.split(",")]

    exclude_set = set()
    if exclude_paths:
        for ep in exclude_paths:
            try:
                exclude_set.add(Path(ep).resolve())
            except Exception:
                pass

    if recursive:
        for root, _dirs, filenames in os.walk(directory):
            for filename in filenames:
                for pattern in pattern_list:
                    # 用 fnmatchcase 替代 fnmatch：大小写敏感、行为更可预测；
                    # 用户传 -p 自定义 pattern 时，fnmatch 风格的 [seq] 通配符仍生效，
                    # 与 Python 文档与社区惯例一致。
                    if fnmatch.fnmatchcase(filename, pattern):
                        files.append(Path(root) / filename)
                        break
    else:
        for pattern in pattern_list:
            files.extend(directory.glob(pattern))

    seen = set()
    unique_files = []
    for f in sorted(files, key=lambda p: str(p)):
        resolved = f.resolve()
        if resolved in exclude_set:
            continue
        if resolved not in seen:
            seen.add(resolved)
            unique_files.append(f)
    return unique_files


def read_file(filepath, header=0, sheet_name=None, first_sheet_only=False):
    """读取单个文件，返回 {sheet_name: DataFrame} 字典。

    first_sheet_only: 为 True 时，Excel 文件只读取第一个 sheet。
    """
    filepath = Path(filepath)
    suffix = filepath.suffix.lower()

    if suffix == ".csv":
        # A-2 P1-5：CSV 编码容错——先 UTF-8（含 BOM），失败按 GBK / GB18030 重试；
        # 仍失败则按"自动跳过"语义只打警告，不阻塞同组合并。
        encodings_to_try = ("utf-8-sig", "utf-8", "gbk", "gb18030")
        last_err = None
        for enc in encodings_to_try:
            try:
                df = pd.read_csv(
                    filepath, header=header, dtype=str,
                    keep_default_na=False, encoding=enc,
                )
                if df.empty:
                    return {}
                return {"Sheet1": df}
            except UnicodeDecodeError as e:
                last_err = e
                continue
            except Exception as e:
                logger.warning(f"无法读取 CSV 文件 {filepath}: {e}")
                return {}
        logger.warning(
            f"无法读取 CSV 文件 {filepath}：候选编码 {encodings_to_try} 全部失败"
            f"（最后错误: {last_err}）；已跳过"
        )
        return {}

    if suffix in (".xlsx", ".xls", ".xlsm"):
        try:
            xl = pd.ExcelFile(filepath)
            sheet_names = xl.sheet_names
            if first_sheet_only and sheet_names:
                sheet_names = [sheet_names[0]]
            result = {}
            for sn in sheet_names:
                if sheet_name is not None and sn not in sheet_name:
                    continue
                df = pd.read_excel(filepath, sheet_name=sn, header=header, dtype=str, keep_default_na=False)
                if not df.empty:
                    result[sn] = df
            return result
        except Exception as e:
            print(f"  警告: 无法读取 Excel 文件 {filepath}: {e}", file=sys.stderr)
            return {}

    return {}


def jaccard_similarity(set_a, set_b):
    """计算两个集合的 Jaccard 相似度。"""
    if not set_a and not set_b:
        return 1.0
    intersection = len(set_a & set_b)
    union = len(set_a | set_b)
    return intersection / union if union > 0 else 0.0


def group_files_by_columns(sheet_files_data, threshold=0.8):
    """
    对同一个 sheet 下的多个文件按列名兼容性分组。
    sheet_files_data: [(filename, columns_set), ...]
    threshold: Jaccard 相似度阈值，默认 0.8
    返回: [[(filename, columns_set), ...], ...] 分组列表
    """
    if not sheet_files_data:
        return []
    if len(sheet_files_data) == 1:
        return [sheet_files_data]

    # 贪心分组：从第一个未分组的文件开始，找出所有与组内任意成员兼容的文件
    # 要求新加入的文件与组内所有已有文件的相似度均 >= threshold
    remaining = list(sheet_files_data)
    groups = []
    while remaining:
        seed = remaining.pop(0)
        seed_cols = seed[1]
        group = [seed]
        group_col_sets = [seed_cols]
        new_remaining = []
        for item in remaining:
            item_cols = item[1]
            # 检查与组内所有成员的相似度
            if all(jaccard_similarity(member_cols, item_cols) >= threshold for member_cols in group_col_sets):
                group.append(item)
                group_col_sets.append(item_cols)
            else:
                new_remaining.append(item)
        remaining = new_remaining
        groups.append(group)
    return groups


def analyze_compatibility(files_data, threshold=0.5):
    """
    分析所有文件的字段兼容性。
    files_data: {filepath: {sheet_name: df}}
    返回: {sheet_name: [group_info, ...]}
    group_info: {
        'files': [filepath, ...],
        'common_columns': [col, ...],
        'all_columns': [col, ...],
        'similarity_score': float
    }
    """
    # 按 sheet 组织数据
    sheet_data = {}
    for filepath, sheets in files_data.items():
        for sheet_name, df in sheets.items():
            cols = set(df.columns.astype(str))
            if sheet_name not in sheet_data:
                sheet_data[sheet_name] = []
            sheet_data[sheet_name].append((filepath, cols, df))

    result = {}
    for sheet_name, file_col_data in sheet_data.items():
        # 提取 (filepath, columns_set) 用于分组
        simple_data = [(fc[0], fc[1]) for fc in file_col_data]
        groups = group_files_by_columns(simple_data, threshold)

        group_infos = []
        for group in groups:
            file_paths = [item[0] for item in group]
            all_cols_sets = [item[1] for item in group]
            common_cols = set.intersection(*all_cols_sets) if all_cols_sets else set()
            all_cols = set.union(*all_cols_sets) if all_cols_sets else set()

            # 计算组内平均相似度
            n = len(group)
            if n <= 1:
                avg_sim = 1.0
            else:
                sims = []
                for i in range(n):
                    for j in range(i + 1, n):
                        sims.append(jaccard_similarity(group[i][1], group[j][1]))
                avg_sim = sum(sims) / len(sims) if sims else 1.0

            group_infos.append({
                'files': file_paths,
                'common_columns': sorted(common_cols),
                'all_columns': sorted(all_cols),
                'similarity_score': round(avg_sim, 3),
                'sheet_name': sheet_name,
            })
        result[sheet_name] = group_infos
    return result


def generate_plan_markdown(compatibility_result, output_path="merged.xlsx"):
    """根据兼容性分析结果生成合并清单 Markdown。"""
    lines = [
        "# Excel 合并清单",
        "",
        "> 本文件由 `zm-excels-merge --preview` 自动生成。",
        "> 你可以手动编辑下方的分组（增删文件、修改输出 sheet 名），",
        "> 然后使用 `--plan 合并清单.md` 执行实际合并。",
        ">",
        "> 编辑说明：",
        "> - `文件列表`：该分组包含的文件路径，每行一个，以 `- ` 开头",
        "> - `输出 sheet 名`：合并后该分组在输出文件中的 sheet 名称",
        "> - 删除整个分组段落可跳过该组合并",
        "",
    ]

    group_counter = 1
    for sheet_name, groups in compatibility_result.items():
        for group in groups:
            common_str = ", ".join(group['common_columns'])
            all_str = ", ".join(group['all_columns'])
            output_sheet = group.get('output_sheet', f"合并_{sheet_name}_组{group_counter}")

            lines.append(f"## 分组 {group_counter}: {sheet_name}")
            lines.append("- **文件列表**:")
            for f in group['files']:
                lines.append(f"  - {f}")
            lines.append(f"- **Sheet 名**: {group.get('sheet_name', sheet_name)}")
            lines.append(f"- **共同列** ({len(group['common_columns'])} 个): {common_str}")
            lines.append(f"- **全部列** ({len(group['all_columns'])} 个): {all_str}")
            lines.append(f"- **组内相似度**: {group['similarity_score']}")
            lines.append(f"- **输出 sheet 名**: {output_sheet}")
            lines.append("")
            group_counter += 1

    lines.append("---")
    lines.append(f"**输出文件**: {output_path}")
    lines.append(f"**总分组数**: {group_counter - 1}")
    return "\n".join(lines)


def parse_plan_markdown(plan_path):
    """解析合并清单 Markdown，返回分组列表。"""
    content = Path(plan_path).read_text(encoding="utf-8")
    groups = []

    # 按 "## 分组" 分割
    sections = re.split(r'\n## 分组 \d+:', content)
    for section in sections[1:]:  # 跳过第一个（是标题部分）
        group = {}
        lines = section.strip().split('\n')
        in_file_list = False

        for line in lines:
            stripped = line.strip()
            if stripped.startswith('- **文件列表**:'):
                in_file_list = True
                # 兼容旧格式：文件列表在同一行，冒号后有内容
                files_str = stripped.split(':', 1)[1].strip()
                if files_str:
                    group['files'] = [f.strip() for f in re.split(r'[,，、]', files_str) if f.strip()]
                else:
                    group['files'] = []
            elif in_file_list and stripped.startswith('- ') and not stripped.startswith('- **'):
                # 子项格式：文件路径（以 "- " 开头但不是其他字段）
                path = stripped[2:].strip()
                if path:
                    group.setdefault('files', []).append(path)
            elif stripped.startswith('- **'):
                # 遇到其他字段，退出文件列表收集状态
                in_file_list = False
                if stripped.startswith('- **Sheet 名**:'):
                    group['sheet_name'] = stripped.split(':', 1)[1].strip()
                elif stripped.startswith('- **输出 sheet 名**:'):
                    group['output_sheet'] = stripped.split(':', 1)[1].strip()
            elif stripped == '':
                in_file_list = False

        if 'files' in group and group['files']:
            groups.append(group)

    return groups


def merge_group(files, sheet_name, header=0, add_source=False, source_col="来源文件", all_columns=None):
    """合并一个分组内的文件数据。返回 DataFrame。"""
    if all_columns is None:
        all_columns = []
    seen_cols = set()
    dfs_with_name = []

    for filepath in files:
        filepath = Path(filepath)
        sheets = read_file(filepath, header=header, sheet_name={sheet_name} if sheet_name else None)
        df = sheets.get(sheet_name)
        if df is None or df.empty:
            print(f"  跳过: {filepath.name} 中未找到 sheet '{sheet_name}' 或无数据")
            continue
        dfs_with_name.append((df, filepath.name))
        for col in df.columns:
            if col not in seen_cols:
                seen_cols.add(col)
                all_columns.append(col)

    if not dfs_with_name:
        return None

    aligned_dfs = []
    for df, fname in dfs_with_name:
        aligned = pd.DataFrame()
        for col in all_columns:
            if col in df.columns:
                aligned[col] = df[col]
            else:
                aligned[col] = pd.NA
        if add_source and source_col not in aligned.columns:
            aligned[source_col] = fname
        aligned_dfs.append(aligned)

    return pd.concat(aligned_dfs, ignore_index=True)


def _sanitize_sheet_name(raw_name) -> str:
    """对 sheet 名做 Excel 合法性归一化（A-2 P2-5）：
    - 转为 str（兼容 int / None）
    - 去掉末尾空格与点（Excel 打开时会自动去除；先归一化避免 used_names 漂移）
    - 替换 \\ / * ? : [ ] 为下划线
    - 截到 31 字符
    - 空名兜底为 "_"
    """
    if raw_name is None:
        return "_"
    name = str(raw_name).rstrip().rstrip(".")
    name = EXCEL_INVALID_SHEET_CHARS.sub("_", name)
    if not name:
        return "_"
    if len(name) > MAX_SHEET_NAME_LEN:
        name = name[:MAX_SHEET_NAME_LEN]
    return name


def _dedupe_sheet_name(raw_name, used_names: set[str]) -> str:
    """把 sheet 名做 Excel 归一化后截到 31 字符；used_names 冲突时追加 _2/_3...。

    Excel 单个 sheet 名最长 31 字符；多源 sheet 截断后前 31 字符相同时
    直接写入会导致后写覆盖前写。这里维护 used_names，冲突时用 "_N" 后缀
    重新构造唯一名，并保证总长仍 <= 31。
    """
    candidate = _sanitize_sheet_name(raw_name)
    if candidate not in used_names:
        return candidate
    stem = candidate[:MAX_SHEET_NAME_LEN - 3]  # 留出 "_NN" 后缀空间
    n = 2
    while True:
        suffix = f"_{n}"
        new_name = stem + suffix
        if new_name not in used_names:
            logger.warning(
                f"sheet 名 '{raw_name}' 与已存在名冲突，已重命名为 '{new_name}'"
            )
            return new_name
        n += 1
        if n > 9999:
            raise RuntimeError(f"sheet 名去重失败：'{raw_name}' 在 9999 次尝试后仍冲突")


def write_output(merged_sheets, output_path, force: bool = False):
    """写入输出文件。

    A-2 P0-2：覆盖前给出警告（除非 --force）；通过 .tmp + os.replace 实现原子写，
    进程崩溃时不会留下破损 xlsx/csv。
    """
    output_path = Path(output_path)
    suffix = output_path.suffix.lower()

    if not merged_sheets:
        logger.error("没有数据可写入")
        sys.exit(1)

    if output_path.exists() and not force:
        logger.warning(
            f"'{output_path}' 已存在，将被覆盖（用 --force 抑制此警告）"
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = output_path.with_name(output_path.name + ".tmp")

    try:
        if suffix == ".csv":
            all_dfs = list(merged_sheets.values())
            combined = pd.concat(all_dfs, ignore_index=True)
            combined.to_csv(tmp_path, index=False, encoding="utf-8-sig")
            os.replace(tmp_path, output_path)
            logger.info(f"已保存 CSV: {output_path} ({len(combined)} 行)")
        else:
            with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
                used_names: set[str] = set()
                for sheet_name, df in merged_sheets.items():
                    safe_name = _dedupe_sheet_name(sheet_name, used_names)
                    used_names.add(safe_name)
                    df.to_excel(writer, sheet_name=safe_name, index=False)
                    logger.info(f"已写入 sheet: {safe_name} ({len(df)} 行)")
            os.replace(tmp_path, output_path)
            logger.info(f"已保存 XLSX: {output_path}")
    except Exception:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass
        raise


def infer_output_path(directory, files_str, explicit_output):
    """
    推断输出路径。
    - 如果用户显式指定了输出路径，直接使用
    - 如果输入是目录且未指定输出，在父目录创建 目录名-excel-merging/merged.xlsx
    - 如果输入是单个文件（-d 或 -f）且未指定输出，在同目录生成 文件名_merged.xlsx
    """
    if explicit_output is not None:
        return Path(explicit_output)

    if directory:
        dir_path = Path(directory).resolve()
        if dir_path.is_file():
            # 单文件输入：在同目录生成 文件名_merged.xlsx
            return dir_path.parent / f"{dir_path.stem}_merged.xlsx"
        parent = dir_path.parent
        output_dir = parent / f"{dir_path.name}-excel-merging"
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir / "merged.xlsx"

    # 尝试从 -f 参数推断
    if files_str:
        files = parse_file_paths(files_str)
        if files:
            first_file = files[0]
            if first_file.is_file():
                return first_file.parent / f"{first_file.stem}_merged.xlsx"
            # 如果第一个路径是目录，按目录处理
            parent = first_file.parent
            output_dir = parent / f"{first_file.name}-excel-merging"
            output_dir.mkdir(parents=True, exist_ok=True)
            return output_dir / "merged.xlsx"

    return Path("merged.xlsx")


def collect_output_exclude_paths(directory, files_str, explicit_output, primary_output):
    """
    收集所有可能的输出文件路径，用于在扫描输入目录时排除，避免上次运行的
    输出被当作新输入再次合并造成数据重复。

    返回: Path 列表（已 resolve）。
    """
    exclude = set()

    # 1) 显式 -o
    if explicit_output:
        exclude.add(Path(explicit_output).resolve())
        # 同一个 stem 的其他常见扩展
        stem = Path(explicit_output).with_suffix("").resolve()
        for suf in (".xlsx", ".csv", ".xls", ".xlsm"):
            exclude.add(Path(str(stem) + suf).resolve())

    # 2) 主输出（auto 或显式）
    if primary_output is not None:
        exclude.add(Path(primary_output).resolve())

    # 3) 自动推断的输出（与显式 -o 不同的另一条路径）
    auto = infer_output_path(directory, files_str, None)
    if auto is not None:
        exclude.add(Path(auto).resolve())
        stem = Path(auto).with_suffix("").resolve()
        for suf in (".xlsx", ".csv", ".xls", ".xlsm"):
            exclude.add(Path(str(stem) + suf).resolve())

    # 4) -f 单文件输入：每个文件的 <stem>_merged.{xlsx,csv}
    if files_str:
        for f in parse_file_paths(files_str):
            if f.is_file():
                stem = (f.parent / f.stem).resolve()
                for suf in ("_merged.xlsx", "_merged.csv"):
                    exclude.add(Path(str(stem) + suf).resolve())

    # 5) -d 目录输入：input 目录中所有 merged.* / *_merged.* 也排除
    # 应对"上次 -d input -o input/merged.xlsx" 之类的历史输出仍在 input 目录中的情况
    if directory:
        dir_path = Path(directory).resolve()
        if dir_path.is_dir():
            for pattern in (
                "merged.xlsx", "merged.csv", "merged.xls", "merged.xlsm",
                "*_merged.xlsx", "*_merged.csv", "*_merged.xls", "*_merged.xlsm",
                "merged*.xlsx", "merged*.csv", "merged*.xls", "merged*.xlsm",
            ):
                for f in dir_path.glob(pattern):
                    if f.is_file():
                        exclude.add(f.resolve())

    return sorted(exclude)


def collect_input_files(directory, files_str, output_exclude_paths, recursive=False):
    """收集本次合并的全部输入文件（已去重、已排除输出）。"""
    files = []
    if files_str:
        files.extend(parse_file_paths(files_str))
    if directory:
        files.extend(
            discover_files(
                directory, DEFAULT_PATTERN, recursive,
                exclude_paths=output_exclude_paths,
            )
        )
    seen = set()
    unique = []
    for f in files:
        try:
            resolved = f.resolve()
        except OSError:
            continue
        if resolved in seen:
            continue
        seen.add(resolved)
        unique.append(f)
    return unique


def validate_plan_paths(plan_groups, input_set):
    """A-2 P0-1：--plan 模式沙箱——plan 中 file 路径必须落在本次 input 范围内。

    input_set 为空时（如用户只传 --plan 不带 -d/-f）不校验；否则 plan 中任何
    不在 input_set 内的路径都视为越界，stderr 报错并 sys.exit(2)。
    """
    if not input_set:
        return
    for gi, group in enumerate(plan_groups, 1):
        for fp_str in group.get("files", []):
            try:
                resolved = Path(fp_str).resolve()
            except OSError as e:
                logger.error(
                    f"plan 分组 {gi} 中文件 '{fp_str}' 路径无效: {e}；已拒绝"
                )
                sys.exit(2)
            if resolved not in input_set:
                logger.error(
                    f"plan 分组 {gi} 中文件 '{fp_str}' 不在本次合并输入范围内，"
                    f"已拒绝（避免读取/写入 input 之外的任意文件）"
                )
                sys.exit(2)


def merge_files_no_group(files_data, add_source, source_col):
    """A-2 P1-1 抽函数：按同名 sheet 直接合并（不分组）。

    files_data: {filepath: {sheet_name: df}}
    返回: {sheet_name: merged_df}。
    """
    sheet_data_map = {}
    for filepath, sheets in files_data.items():
        for sheet_name, df in sheets.items():
            sheet_data_map.setdefault(sheet_name, []).append((df, filepath.name))

    all_columns_map = {}
    for sheet_name, df_list in sheet_data_map.items():
        all_columns = []
        seen = set()
        for df, _ in df_list:
            for col in df.columns:
                if col not in seen:
                    seen.add(col)
                    all_columns.append(col)
        all_columns_map[sheet_name] = all_columns

    merged = {}
    for sheet_name, df_list in sheet_data_map.items():
        aligned_dfs = []
        for df, fname in df_list:
            aligned = pd.DataFrame()
            for col in all_columns_map[sheet_name]:
                aligned[col] = df[col] if col in df.columns else pd.NA
            if add_source and source_col not in aligned.columns:
                aligned[source_col] = fname
            aligned_dfs.append(aligned)
        merged[sheet_name] = pd.concat(aligned_dfs, ignore_index=True)
    return merged


def _merge_into_suffix_bucket(merged, merged_all, suffix, n_suffix_groups):
    """把单后缀组的合并结果搬进累计字典；多后缀时为 sheet 名加后缀避免冲突。"""
    if n_suffix_groups <= 1:
        merged_all.update(merged)
        return
    tag = suffix.lstrip(".")
    for sheet_name, df in list(merged.items()):
        merged_all[f"{sheet_name}_{tag}"] = df


def _non_negative_int(value):
    """argparse type：非负整数。"""
    try:
        ivalue = int(value)
    except (TypeError, ValueError):
        raise argparse.ArgumentTypeError(f"必须是整数，得到: {value!r}")
    if ivalue < 0:
        raise argparse.ArgumentTypeError(f"必须 >= 0，得到: {ivalue}")
    return ivalue


def _probability(value):
    """argparse type：0 <= v <= 1 的概率值。"""
    try:
        fvalue = float(value)
    except (TypeError, ValueError):
        raise argparse.ArgumentTypeError(f"必须是数字，得到: {value!r}")
    if fvalue < 0.0 or fvalue > 1.0:
        raise argparse.ArgumentTypeError(f"必须在 [0, 1] 区间内，得到: {fvalue}")
    return fvalue


def main():
    parser = argparse.ArgumentParser(
        description="合并多个 Excel/CSV 文件，支持字段兼容性分析和合并清单。"
    )
    parser.add_argument(
        "-d", "--directory", default=None,
        help="输入目录路径（与 --files 二选一）"
    )
    parser.add_argument(
        "-f", "--files", default=None,
        help="直接指定多个文件路径，用逗号/空格/中文逗号/顿号分隔"
    )
    parser.add_argument(
        "-o", "--output", default=None,
        help="输出文件路径（默认: 目录输入时自动推断为 父目录/目录名-excel-merging/merged.xlsx）"
    )
    parser.add_argument(
        "-r", "--recursive", action="store_true", help="递归搜索子目录"
    )
    parser.add_argument(
        "-p", "--pattern", default=DEFAULT_PATTERN,
        help=f"文件匹配模式，逗号分隔（默认: {DEFAULT_PATTERN}）"
    )
    parser.add_argument(
        "-s", "--sheets", default=None, help="只合并指定 sheet 名，逗号分隔"
    )
    parser.add_argument(
        "--header", type=_non_negative_int, default=0,
        help="表头行号，0-based，必须为非负整数（默认: 0）"
    )
    parser.add_argument(
        "--add-source", action="store_true", help="添加来源列"
    )
    parser.add_argument(
        "--source-col", default=DEFAULT_SOURCE_COL,
        help=f"来源列名（默认: {DEFAULT_SOURCE_COL}）"
    )
    parser.add_argument(
        "--preview", action="store_true",
        help="预览模式：分析字段兼容性并生成合并清单，不执行实际合并"
    )
    parser.add_argument(
        "--plan", default=None,
        help="执行模式：读取合并清单文件并按清单执行合并"
    )
    parser.add_argument(
        "--similarity-threshold", type=_probability, default=DEFAULT_SIMILARITY_THRESHOLD,
        help=f"字段相似度阈值（Jaccard），必须在 [0, 1]（默认: {DEFAULT_SIMILARITY_THRESHOLD}）"
    )
    parser.add_argument(
        "--plan-output", default="合并清单.md",
        help="预览模式下合并清单的输出路径（默认: 合并清单.md）"
    )
    parser.add_argument(
        "--no-auto-group", action="store_true",
        help="禁用默认模式下的自动字段分组，强制按同名 sheet 合并（恢复旧行为）"
    )
    parser.add_argument(
        "--merge-sheets", action="store_true",
        help="合并模式：将每个文件内的所有 sheet 垂直合并为单个 sheet，而非按同名 sheet 跨文件合并"
    )
    parser.add_argument(
        "--force", action="store_true",
        help="强制覆盖已存在的输出文件而不打印警告"
    )
    parser.add_argument(
        "--log-level", default="INFO",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        help="日志级别（默认: INFO）"
    )
    args = parser.parse_args()
    _set_log_level(args.log_level)

    # 解析 sheet 筛选
    target_sheets = None
    if args.sheets:
        target_sheets = {s.strip() for s in args.sheets.split(",")}

    # 推断输出路径
    output_path = infer_output_path(args.directory, args.files, args.output)

    # 模式分支
    # 收集所有可能的输出路径作为排除集，避免上次运行的输出被当作新输入
    output_exclude_paths = collect_output_exclude_paths(
        args.directory, args.files, args.output, output_path
    )

    if args.preview:
        # 预览模式：收集文件并生成合并清单
        files = collect_input_files(args.directory, args.files, output_exclude_paths, args.recursive)

        if not files:
            logger.error("未找到有效的输入文件")
            sys.exit(1)

        logger.info(f"发现 {len(files)} 个输入文件:")
        for f in files:
            logger.info(f"  - {f}")

        # 按后缀分组，同后缀文件才合并
        files_by_suffix = {}
        for f in files:
            suf = f.suffix.lower()
            files_by_suffix.setdefault(suf, []).append(f)

        all_compatibility = {}
        group_counter = 1
        for suffix, suffix_files in sorted(files_by_suffix.items()):
            logger.info(f"\n--- 后缀组 {suffix}: {len(suffix_files)} 个文件 ---")

            files_data = {}
            for filepath in suffix_files:
                sheets = read_file(filepath, header=args.header, sheet_name=target_sheets, first_sheet_only=True)
                if sheets:
                    files_data[filepath] = sheets
                    for sn, df in sheets.items():
                        logger.info(f"  [{filepath.name}] sheet '{sn}': {len(df)} 行 x {len(df.columns)} 列")

            if not files_data:
                logger.warning(f"后缀 {suffix} 的文件都无法读取或无数据，跳过")
                continue

            compatibility = analyze_compatibility(files_data, threshold=args.similarity_threshold)

            for sheet_name, groups in compatibility.items():
                for group in groups:
                    group['output_sheet'] = f"合并_{sheet_name}_{suffix.lstrip('.')}_组{group_counter}"
                    group['sheet_name'] = sheet_name
                all_compatibility[f"{sheet_name}_{suffix.lstrip('.')}"] = groups
                group_counter += len(groups)

        if not all_compatibility:
            logger.error("所有文件都无法读取或无数据")
            sys.exit(1)

        total_groups = sum(len(g) for g in all_compatibility.values())
        logger.info(f"\n建议分为 {total_groups} 个组合并")

        for sheet_name, groups in all_compatibility.items():
            logger.info(f"\n  Sheet '{sheet_name}':")
            for i, group in enumerate(groups, 1):
                logger.info(
                    f"    组{i}: {len(group['files'])} 个文件, "
                    f"共同列 {len(group['common_columns'])} 个, "
                    f"相似度 {group['similarity_score']}"
                )
                for f in group['files']:
                    logger.info(f"      - {f.name}")

        plan_md = generate_plan_markdown(all_compatibility, output_path=output_path)
        # 若未显式指定 --plan-output，将清单保存到与输出文件同目录
        if args.plan_output == "合并清单.md":
            plan_path = output_path.parent / "合并清单.md"
        else:
            plan_path = Path(args.plan_output)
        plan_path.parent.mkdir(parents=True, exist_ok=True)
        plan_path.write_text(plan_md, encoding="utf-8")
        logger.info(f"\n合并清单已保存: {plan_path.resolve()}")
        logger.info("请编辑该文件后，使用 --plan 参数重新调用以执行合并。")
        return

    elif args.plan:
        # 执行模式：按清单合并
        plan_path = Path(args.plan)
        if not plan_path.exists():
            logger.error(f"合并清单不存在: {plan_path}")
            sys.exit(1)

        logger.info(f"\n读取合并清单: {plan_path}")
        groups = parse_plan_markdown(plan_path)
        logger.info(f"解析到 {len(groups)} 个合并分组")

        # A-2 P0-1：--plan 沙箱——只允许读取本次 input 范围内的文件
        input_set = {
            f.resolve() for f in collect_input_files(
                args.directory, args.files, output_exclude_paths, args.recursive
            )
        }
        validate_plan_paths(groups, input_set)

        merged_sheets = {}
        for group in groups:
            files_in_group = group['files']
            sheet_name = group.get('sheet_name', 'Sheet1')
            output_sheet = group.get('output_sheet', sheet_name)

            logger.info(f"\n合并分组: {output_sheet}")
            logger.info(f"  文件: {', '.join(files_in_group)}")
            logger.info(f"  Sheet: {sheet_name}")

            merged_df = merge_group(
                files_in_group, sheet_name,
                header=args.header,
                add_source=args.add_source,
                source_col=args.source_col
            )
            if merged_df is not None and not merged_df.empty:
                merged_sheets[output_sheet] = merged_df
                logger.info(f"  结果: {len(merged_df)} 行 x {len(merged_df.columns)} 列")
            else:
                logger.info("  跳过: 无数据")

        if merged_sheets:
            logger.info("\n写入输出...")
            write_output(merged_sheets, output_path, force=args.force)
        else:
            logger.error("没有数据可写入")
            sys.exit(1)
        return

    else:
        # 默认直接合并模式（智能分组版）
        files = collect_input_files(args.directory, args.files, output_exclude_paths, args.recursive)

        if not files:
            logger.error("未找到有效的输入文件")
            sys.exit(1)

        logger.info(f"发现 {len(files)} 个输入文件:")
        for f in files:
            logger.info(f"  - {f}")

        # 多 sheet 合并模式：将每个文件内的所有 sheet 按字段兼容性分组合并
        if args.merge_sheets:
            logger.info(
                f"\n开始合并各文件内的所有 sheet（字段兼容性阈值: {args.similarity_threshold}）..."
            )
            merged = {}
            for filepath in files:
                sheets = read_file(filepath, header=args.header, sheet_name=target_sheets)
                if not sheets:
                    logger.info(f"  跳过: {filepath.name} 无可用 sheet")
                    continue

                # 按字段兼容性对 sheet 分组
                sheet_col_data = [(sn, set(df.columns.astype(str))) for sn, df in sheets.items()]
                groups = group_files_by_columns(sheet_col_data, threshold=args.similarity_threshold)

                if len(groups) == 1:
                    # 所有 sheet 字段兼容，合并为一个 sheet
                    group_sheets = groups[0]
                    all_columns = []
                    seen_cols = set()
                    for sheet_name, _cols in group_sheets:
                        df = sheets[sheet_name]
                        for col in df.columns:
                            if col not in seen_cols:
                                seen_cols.add(col)
                                all_columns.append(col)

                    aligned_dfs = []
                    for sheet_name, _cols in group_sheets:
                        df = sheets[sheet_name]
                        aligned = pd.DataFrame()
                        for col in all_columns:
                            if col in df.columns:
                                aligned[col] = df[col]
                            else:
                                aligned[col] = pd.NA
                        if args.add_source:
                            if args.source_col not in aligned.columns:
                                aligned[args.source_col] = filepath.name
                            if "来源Sheet" not in aligned.columns:
                                aligned["来源Sheet"] = sheet_name
                        aligned_dfs.append(aligned)

                    merged_df = pd.concat(aligned_dfs, ignore_index=True)
                    if len(files) == 1:
                        output_sheet = "Merged"
                    else:
                        output_sheet = filepath.stem[:31]
                    merged[output_sheet] = merged_df
                    logger.info(
                        f"  [{filepath.name}] 合并 {len(sheets)} 个 sheet → "
                        f"{len(merged_df)} 行 x {len(merged_df.columns)} 列"
                    )
                else:
                    # 多个不兼容组，每组输出一个 sheet
                    for gi, group in enumerate(groups, 1):
                        all_columns = []
                        seen_cols = set()
                        for sheet_name, _cols in group:
                            df = sheets[sheet_name]
                            for col in df.columns:
                                if col not in seen_cols:
                                    seen_cols.add(col)
                                    all_columns.append(col)

                        aligned_dfs = []
                        for sheet_name, _cols in group:
                            df = sheets[sheet_name]
                            aligned = pd.DataFrame()
                            for col in all_columns:
                                if col in df.columns:
                                    aligned[col] = df[col]
                                else:
                                    aligned[col] = pd.NA
                            if args.add_source:
                                if args.source_col not in aligned.columns:
                                    aligned[args.source_col] = filepath.name
                                if "来源Sheet" not in aligned.columns:
                                    aligned["来源Sheet"] = sheet_name
                            aligned_dfs.append(aligned)

                        merged_df = pd.concat(aligned_dfs, ignore_index=True)
                        if len(files) == 1:
                            output_sheet = f"Merged_组{gi}"
                        else:
                            output_sheet = f"{filepath.stem[:25]}_组{gi}"
                        merged[output_sheet] = merged_df
                        common_cols = set.intersection(*[_cols for _, _cols in group])
                        all_group_cols = set.union(*[_cols for _, _cols in group])
                        logger.info(
                            f"  [{filepath.name}] 组{gi}: 合并 {len(group)} 个 sheet, "
                            f"共同列 {len(common_cols)} 个, 全部列 {len(all_group_cols)} 个, "
                            f"→ {len(merged_df)} 行 x {len(merged_df.columns)} 列"
                        )

            if merged:
                logger.info("\n写入输出...")
                write_output(merged, output_path, force=args.force)
                logger.info("完成")
            else:
                logger.error("没有数据可写入")
                sys.exit(1)
            return

        # 按后缀分组，同后缀文件才合并；Excel 默认只取第一个 sheet
        files_by_suffix = {}
        for f in files:
            suf = f.suffix.lower()
            files_by_suffix.setdefault(suf, []).append(f)

        merged_all = {}

        for suffix, suffix_files in sorted(files_by_suffix.items()):
            logger.info(f"\n--- 处理后缀组 {suffix}: {len(suffix_files)} 个文件 ---")

            files_data = {}
            for filepath in suffix_files:
                sheets = read_file(filepath, header=args.header, sheet_name=target_sheets, first_sheet_only=True)
                if sheets:
                    files_data[filepath] = sheets
                    for sn, df in sheets.items():
                        logger.info(f"  [{filepath.name}] sheet '{sn}': {len(df)} 行 x {len(df.columns)} 列")

            if not files_data:
                logger.warning(f"后缀 {suffix} 的文件都无法读取或无数据，跳过")
                continue

            # 判断是否需要自动分组
            if args.no_auto_group:
                # 旧行为：强制按同名 sheet 合并
                logger.info("\n开始合并（强制同名 sheet 模式，--no-auto-group）...")
                merged = merge_files_no_group(files_data, args.add_source, args.source_col)
                _merge_into_suffix_bucket(merged, merged_all, suffix, len(files_by_suffix))
                continue

            # 智能分组模式：先分析兼容性，再按组合并
            logger.info(f"\n分析字段兼容性（阈值: {args.similarity_threshold}）...")
            compatibility = analyze_compatibility(files_data, threshold=args.similarity_threshold)

            total_groups = sum(len(g) for g in compatibility.values())
            total_sheets = len(compatibility)

            if total_groups == total_sheets:
                # 每个 sheet 只有一组，说明列名完全一致，按原有行为合并
                logger.info("检测到所有文件字段完全一致，按同名 sheet 直接合并...")
                merged = merge_files_no_group(files_data, args.add_source, args.source_col)
                _merge_into_suffix_bucket(merged, merged_all, suffix, len(files_by_suffix))
                continue

            # 字段不完全一致，需要按组分别合并
            logger.info(f"检测到字段差异，自动分为 {total_groups} 个组合并:")
            merged = {}
            group_counter = 1
            for sheet_name, groups in compatibility.items():
                for group in groups:
                    file_paths = group['files']
                    n_files = len(file_paths)
                    output_sheet = f"{sheet_name}_组{group_counter}"

                    logger.info(f"\n  分组 {group_counter}: {output_sheet}")
                    logger.info(f"    文件 ({n_files} 个):")
                    for f in file_paths:
                        logger.info(f"      - {f.name}")
                    logger.info(f"    共同列 ({len(group['common_columns'])} 个): {', '.join(group['common_columns'])}")
                    if len(group['all_columns']) > len(group['common_columns']):
                        extra = set(group['all_columns']) - set(group['common_columns'])
                        logger.info(f"    差异列 ({len(extra)} 个): {', '.join(sorted(extra))}")

                    merged_df = merge_group(
                        file_paths, sheet_name,
                        header=args.header,
                        add_source=args.add_source,
                        source_col=args.source_col
                    )
                    if merged_df is not None and not merged_df.empty:
                        merged[output_sheet] = merged_df
                        logger.info(f"    结果: {len(merged_df)} 行 x {len(merged_df.columns)} 列")
                    else:
                        logger.info("    跳过: 无数据")
                    group_counter += 1

            _merge_into_suffix_bucket(merged, merged_all, suffix, len(files_by_suffix))

        # 所有后缀组处理完成后写入输出
        if merged_all:
            logger.info("\n写入输出...")
            write_output(merged_all, output_path, force=args.force)
            logger.info("完成")
        else:
            logger.error("没有数据可写入")
            sys.exit(1)


if __name__ == "__main__":
    main()
