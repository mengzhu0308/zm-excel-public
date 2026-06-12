#!/usr/bin/env python3
"""
Excel 查询脚本：读取 Excel，预览数据，交互查询，输出同目录同名 CSV。
"""

import argparse
import ast
import json
import operator
import os
import re
import sys
from functools import reduce
from pathlib import Path

try:
    import pandas as pd
except ImportError as e:
    print("错误: 缺少 pandas。请安装: pip install pandas openpyxl xlrd", file=sys.stderr)
    sys.exit(1)


# 从 VERSION.yaml 读取 skill 版本号；失败时回退到硬编码值。
def _load_skill_version() -> str:
    """优先从同目录的 VERSION.yaml 读取 skill_info.version；缺失或无法解析时抛 RuntimeError。

    A-1 P1-3: 之前回退到硬编码 "0.3.0" 会与 VERSION.yaml 漂移；改为显式抛错。
    包裹层 __version__ 在 import 时 try/except，配置错误时 --version 显示 "unknown"
    提示用户检查 VERSION.yaml，主流程不中断。
    """
    try:
        import yaml  # PyYAML；缺则降级到正则
    except ImportError:
        yaml = None
    version_yaml = Path(__file__).resolve().parent.parent / "VERSION.yaml"
    if not version_yaml.exists():
        raise RuntimeError(f"VERSION.yaml 缺失: {version_yaml}")
    try:
        if yaml is not None:
            with open(version_yaml, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            v = (data.get("skill_info") or {}).get("version")
            if v:
                return str(v)
        else:
            # 降级：正则匹配 `version: 0.2.0`
            text = version_yaml.read_text(encoding="utf-8")
            m = re.search(r"^\s*version:\s*['\"]?([0-9A-Za-z.\-_]+)", text, re.MULTILINE)
            if m:
                return m.group(1)
    except Exception as e:
        raise RuntimeError(f"无法解析 VERSION.yaml: {e}") from e
    raise RuntimeError(f"VERSION.yaml 缺少 skill_info.version: {version_yaml}")


try:
    __version__ = _load_skill_version()
except RuntimeError:
    __version__ = "unknown"  # 配置错误时 --version 提示 unknown；主流程不中断


def _write_output(df, path, fmt='csv'):
    """根据格式写入输出文件。父目录不存在时自动创建。"""
    path = Path(path)
    if path.parent and not path.parent.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
    if fmt == 'xlsx':
        df.to_excel(path, index=False, engine='openpyxl')
    else:
        df.to_csv(path, index=False, encoding='utf-8-sig')


def _read_input(input_path, sheet=None, header=0):
    """统一的输入读取入口。

    根据文件后缀路由到 `pd.read_excel` 或 `pd.read_csv`：
    - `.csv`：走 `read_csv`，`sheet` 参数被忽略；
    - `.xlsx` / `.xlsm`：走 `read_excel`，引擎 `openpyxl`；
    - `.xls`：走 `read_excel`，引擎 `xlrd`（按需 fallback）。

    关键：pandas 默认 `sheet_name=0`（仅取第一张表）。本函数必须显式传
    `sheet_name=sheet`，让 `sheet=None` 时走 `sheet_name=None`（dict 形式
    拿所有 sheet），与 SKILL.md "未传 --sheet 时多 sheet 分片" 的行为契约一致。
    """
    suffix = Path(input_path).suffix.lower()
    if suffix == '.csv':
        return pd.read_csv(input_path, header=header)
    if suffix in ('.xlsx', '.xlsm'):
        return pd.read_excel(input_path, engine='openpyxl', header=header, sheet_name=sheet)
    if suffix == '.xls':
        return pd.read_excel(input_path, engine='xlrd', header=header, sheet_name=sheet)
    # A-1 P0-3: 未知后缀不再静默 fallback（之前 except Exception 双 try
    # 会把 .txt/.json/无后缀 的真实错误掩盖成 CSV 解析错误）。显式拒绝并提示。
    raise ValueError(
        f"不支持的文件格式: {suffix!r}（支持: .csv / .xlsx / .xlsm / .xls）"
    )


def _ensure_unique(path):
    """如果目标文件已存在，自动追加序号避免覆盖。"""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 1
    while True:
        new_path = parent / f"{stem}_{counter}{suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


# 显式优先级表：列表顺序即"高 → 低"绑定强度（先切 = 绑定更松）。
# 先尝试 `and` 切分意味着 `or` 绑定更松（`A and B or C and D` → `(A and B) or (C and D)`）；
# 与 Python / pandas.DataFrame.query() 默认语义一致。
# 用列表而非 if-elif-else 是为了未来扩展更直观。
# 改顺序前请同步更新 references/where-expression.md 与 _artifacts/precedence_test.py。
# A-2 P1-3: 列表顺序由"低→高"改为"高→低"，与文档 where-expression.md 中
# "从高到低" 阅读顺序对齐，新读者不再被反序索引误导。
_PRECEDENCE: list[tuple[str, str, "operator"]] = [
    (' and ', '&', operator.and_),
    (' or ', '|', operator.or_),
]


# tag / sheet 名等"会拼到路径上的字符串"白名单。
# 禁止 `.` `/` `\\` `:` `*` `?` `"` `<` `>` `|` 等 shell/path 危险字符。
_SAFE_NAME = re.compile(r'^[A-Za-z0-9_\-]+$')


def _validate_tag(tag: str) -> str:
    """校验 tag 仅含安全字符；非法字符抛 ValueError。

    背景：tag 会拼接到输出路径，未白名单化时 `--tag ../../etc` 可越界写文件。
    """
    if not tag:
        raise ValueError("tag 不能为空")
    if not _SAFE_NAME.match(tag):
        raise ValueError(
            f"tag 仅允许字母、数字、下划线、连字符；非法 tag: {tag!r}"
        )
    return tag


def _sanitize_sheet_name(sheet_name: str) -> str:
    """把 sheet 名中的路径分隔符替换为 '_'，避免 `_query_single` 拼路径时
    越出源文件目录。"""
    if not sheet_name:
        return sheet_name
    # 替换常见路径分隔符与 Windows 保留字符
    bad = re.compile(r'[\\/:*?"<>|\.\s]+')
    sanitized = bad.sub('_', sheet_name).strip('_')
    if not sanitized:
        sanitized = 'sheet'
    return sanitized


def _fmt_dtype(dtype):
    """将 pandas dtype 格式化为友好字符串。"""
    name = str(dtype)
    if 'int' in name:
        return '整数'
    if 'float' in name:
        return '浮点数'
    if 'bool' in name:
        return '布尔'
    if 'datetime' in name:
        return '日期时间'
    if 'object' in name:
        return '文本'
    return name


def preview_data(df, n=10, title=None):
    """打印 DataFrame 的预览信息：前 N 行 + 字段元信息。

    A-2 P1-1: n 必须为非负整数；负数会被 pandas 解释为"除最后 |n| 行外全部"，
    与用户直觉相反（与 A-1 P1-7 同源修复）。
    """
    if n < 0:
        raise ValueError(f"preview 行数不能为负数: {n}")
    lines = []
    if title:
        lines.append(f"\n{'=' * 60}")
        lines.append(f"  {title}")
        lines.append(f"{'=' * 60}")

    # 基本统计
    lines.append(f"\n📊 数据概览")
    lines.append(f"   总行数: {len(df)}")
    lines.append(f"   总列数: {len(df.columns)}")

    # 字段信息
    lines.append(f"\n📋 字段信息")
    header = f"   {'列名':<20} {'类型':<10} {'非空数':<8} {'示例值'}"
    lines.append(header)
    lines.append(f"   {'-' * 60}")
    for col in df.columns:
        dtype = _fmt_dtype(df[col].dtype)
        non_null = df[col].notna().sum()
        # 取前 3 个非空值作为示例
        samples = df[col].dropna().head(3).tolist()
        sample_str = ', '.join(str(s)[:20] for s in samples)
        if len(sample_str) > 40:
            sample_str = sample_str[:37] + '...'
        lines.append(f"   {str(col):<20} {dtype:<10} {non_null:<8} {sample_str}")

    # 前 N 行数据
    lines.append(f"\n📄 前 {min(n, len(df))} 行数据")
    preview_df = df.head(n)
    # 使用 pandas 的 to_string 格式化输出
    preview_str = preview_df.to_string(index=False, max_colwidth=25)
    lines.append(preview_str)

    lines.append(f"\n{'=' * 60}\n")
    return '\n'.join(lines)


def show_query_examples(df):
    """根据 DataFrame 的列生成查询示例。"""
    examples = []
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    text_cols = [c for c in df.columns if pd.api.types.is_string_dtype(df[c]) or df[c].dtype == object]

    if numeric_cols:
        col = numeric_cols[0]
        examples.append(f"   where {col} > 100")
    if text_cols:
        col = text_cols[0]
        # 取一个非空值作为示例
        sample_val = df[col].dropna().iloc[0] if not df[col].dropna().empty else '示例值'
        examples.append(f"   where {col} == '{sample_val}'")
        examples.append(f"   where {col} contains '关键词'")
    if numeric_cols and len(numeric_cols) >= 2:
        examples.append(f"   select {numeric_cols[0]},{numeric_cols[1]}")
    if numeric_cols:
        examples.append(f"   sort {numeric_cols[0]} desc")
    if len(df.columns) >= 2:
        # 交互模式 groupby 与 agg 是两条独立命令，与 CLI 的 --groupby/--agg 标志不同。
        examples.append(f"   groupby {df.columns[0]}")
        examples.append(f"   agg sum:{df.columns[1]}")
        examples.append(f"   run")

    return examples


INTERACTIVE_HELP = """
╔══════════════════════════════════════════════════════════════╗
║                    交互式查询命令说明                         ║
╠══════════════════════════════════════════════════════════════╣
║  where <条件>      设置筛选条件，如: age > 30                 ║
║  select <列>       选择列（逗号分隔），如: name,age           ║
║  exclude <列>      排除列（逗号分隔），如: 序号,id            ║
║  sort <规则>       排序，如: salary desc,age asc              ║
║  distinct <列>     按列去重，如: department                   ║
║  groupby <列>      分组列，如: department                     ║
║  agg <规则>        聚合规则，如: sum:salary,count:name        ║
║  limit <N>         限制返回行数                               ║
║  offset <N>        跳过前 N 行                                ║
║  run / 回车        执行当前查询并展示结果                     ║
║  save [路径]       保存结果为 CSV（默认同目录同名）           ║
║  reset             重置所有查询条件                           ║
║  preview           重新显示原始数据预览                       ║
║  skip / next       跳过当前 sheet，进入下一张（多 sheet 时）  ║
║  help              显示此帮助                                 ║
║  quit / exit       退出交互模式                               ║
╚══════════════════════════════════════════════════════════════╝
"""


def interactive_mode(input_path, df, sheet_name=None, fmt='csv'):
    """交互式查询模式：循环读取用户命令，执行查询，展示结果。"""
    print(f"\n📁 文件: {input_path}" + (f" [{sheet_name}]" if sheet_name else ""))
    print(preview_data(df, n=10, title="数据预览"))

    # 生成并显示查询示例
    examples = show_query_examples(df)
    if examples:
        print("💡 您可以尝试以下查询示例：")
        for ex in examples:
            print(ex)
        print()

    print("🔧 输入 'help' 查看命令说明，'quit' 退出交互模式。\n")

    # 当前查询状态
    state = {
        'where': None,
        'select': None,
        'exclude': None,
        'sort': None,
        'distinct': None,
        'groupby': None,
        'agg': None,
        'limit': None,
        'offset': None,
    }

    while True:
        try:
            user_input = input("query> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n👋 再见！")
            break

        if not user_input:
            # 回车 = run
            cmd = 'run'
            args = ''
        else:
            parts = user_input.split(None, 1)
            cmd = parts[0].lower()
            args = parts[1] if len(parts) > 1 else ''

        if cmd in ('quit', 'exit', 'q'):
            print("👋 再见！")
            break

        # A-3 P1-6: 多 sheet 模式下提供"跳过当前 sheet"命令，
        # 让用户不必跑完一张 sheet 的所有命令后才能进入下一张。
        if cmd in ('skip', 'next'):
            print("   ⏭  跳过当前 sheet")
            return  # 直接返回，让外层 dispatch 继续下一张

        elif cmd == 'help':
            print(INTERACTIVE_HELP)

        elif cmd == 'preview':
            print(preview_data(df, n=10, title="数据预览"))

        elif cmd == 'where':
            state['where'] = args
            print(f"   ✓ 筛选条件已设置: {args}")

        elif cmd == 'select':
            state['select'] = args
            print(f"   ✓ 列选择已设置: {args}")

        elif cmd == 'exclude':
            state['exclude'] = args
            print(f"   ✓ 排除列已设置: {args}")

        elif cmd == 'sort':
            state['sort'] = args
            print(f"   ✓ 排序规则已设置: {args}")

        elif cmd == 'distinct':
            state['distinct'] = args
            print(f"   ✓ 去重列已设置: {args}")

        elif cmd == 'groupby':
            state['groupby'] = args
            print(f"   ✓ 分组列已设置: {args}")

        elif cmd == 'agg':
            state['agg'] = args
            print(f"   ✓ 聚合规则已设置: {args}")

        elif cmd == 'limit':
            try:
                v = int(args)
            except ValueError:
                print("   ✗ limit 参数必须是整数")
                continue
            # A-1 P1-7: 与 CLI _apply_pipeline 一致；负数立即拦截而非延后到 run。
            if v < 0:
                print(f"   ✗ limit 不能为负数: {v}")
                continue
            state['limit'] = v
            print(f"   ✓ 行数限制已设置: {v}")

        elif cmd == 'offset':
            try:
                v = int(args)
            except ValueError:
                print("   ✗ offset 参数必须是整数")
                continue
            if v < 0:
                print(f"   ✗ offset 不能为负数: {v}")
                continue
            state['offset'] = v
            print(f"   ✓ 偏移量已设置: {v}")

        elif cmd == 'reset':
            state = {k: None for k in state}
            print("   ✓ 所有查询条件已重置")
            # A-1 P1-8: reset 后提示当前 sheet，避免多 sheet 模式下"上一张 sheet
            # 的 query" 与 "当前 sheet" 边界模糊。
            if sheet_name:
                print(f"   当前 sheet: {sheet_name}")

        elif cmd == 'run':
            result_df, error = _apply_query(df, state)
            if error:
                # A-3 P2-1: 透传 _apply_pipeline_safe 分类字符串（用户错误 vs 内部错误），
                # 与 CLI query_excel 入口的"参数或输入错误" / "内部错误"口径对齐。
                print(f"   ✗ {error}")
                continue

            print(f"\n📄 查询结果（共 {len(result_df)} 行）")
            if len(result_df) > 0:
                print(preview_data(result_df, n=10, title="结果预览"))
            else:
                print("   （无数据）")

            # 显示当前查询条件摘要
            active = [f"{k}={v}" for k, v in state.items() if v is not None]
            if active:
                print(f"   当前条件: {' | '.join(active)}")
            print()

        elif cmd == 'save':
            result_df, error = _apply_query(df, state)
            if error:
                # A-3 P2-1: 同上，透传分类字符串。
                print(f"   ✗ {error}")
                continue

            if args:
                # A-1 P0-1: 交互模式 save 路径与 CLI -o 同语义；用户输入 `save ../escape.csv`
                # 会越界到父目录写文件。`save` 与 CLI --tag 同源风险但语义不同（-o 是完整路径），
                # 这里只拦截路径穿越 `..` 而非字符白名单（`/` 是合法目录分隔符）。
                if '..' in Path(args).parts:
                    print(f"   ✗ save 路径不允许包含 '..'；非法: {args!r}")
                    continue
                out_path = Path(args)
            else:
                suffix = '.xlsx' if fmt == 'xlsx' else '.csv'
                base = Path(input_path).with_suffix('')
                if sheet_name:
                    out_path = Path(str(base) + f'_{sheet_name}{suffix}')
                else:
                    out_path = base.with_suffix(suffix)
                # 默认命名时自动防覆盖
                out_path = _ensure_unique(out_path)

            try:
                _write_output(result_df, out_path, fmt)
                print(f"   ✓ 结果已保存: {out_path}")
            except (OSError, ValueError) as e:
                # A-3 P2-1: 写文件错误也用"参数或输入错误"分类，与 CLI 对齐。
                print(f"   ✗ 参数或输入错误: 保存失败: {e}")
            except Exception as e:
                # 内部错误附 traceback，与 query_excel 入口一致。
                import traceback
                print(f"   ✗ 内部错误: 保存失败: {e}\n{traceback.format_exc()}")

        else:
            print(f"   ✗ 未知命令: {cmd}，输入 'help' 查看可用命令")


def _apply_pipeline(df, params):
    """统一的查询管线：where / exclude / select / distinct / groupby+agg / sort / offset / limit。

    A-2 P0-2: 异常透传给外层 caller，不在内部做 (df, error) 包装。
    一次性查询（_query_single）→ query_excel 入口 try/except 分类；
    交互模式（_apply_query）→ _apply_query_safe 包装为 (df, error) 元组。

    这样外层 query_excel 才能按异常类型正确分类"用户错误"与"内部错误"，
    避免之前 _query_single 把已分类 error 包成 RuntimeError 重抛后
    外层只能走"内部错误"分支、双层包装的 bug。
    """
    # 入口处统一校验 select / exclude / distinct / groupby 引用的列
    # 是否都存在；typo 的列名应被显式拒绝，而不是 silent drop。
    # A-2 P0-3: sort 推迟到 groupby+agg 之后再校验，因为聚合后才会出现
    # `salary_mean` / `count_name` 这类新列。
    _validate_columns(params, df)

    # limit / offset 负数校验：A-3 P1-2 修复。
    # pandas 默认行为 `df.head(-1)` 返回除最后一行外全部，`iloc[:-5]` 跳过最后 5 行，
    # 与用户预期相反，必须显式拦截。
    if params.get('limit') is not None and int(params['limit']) < 0:
        raise ValueError(f"--limit 不能为负数: {params['limit']}")
    if params.get('offset') is not None and int(params['offset']) < 0:
        raise ValueError(f"--offset 不能为负数: {params['offset']}")

    result = df.copy()

    if params.get('where'):
        mask = parse_where(params['where'], result)
        result = result[mask]

    if params.get('exclude'):
        exclude_cols = [c.strip() for c in params['exclude'].split(',')]
        # _validate_columns 已保证所有 exclude 列存在，直接 drop。
        result = result.drop(columns=exclude_cols, errors='ignore')

    if params.get('select'):
        cols = [c.strip() for c in params['select'].split(',')]
        # _validate_columns 已保证所有 select 列存在，直接索引。
        result = result[cols]

    if params.get('distinct'):
        cols = [c.strip() for c in params['distinct'].split(',')]
        # _validate_columns 已保证所有 distinct 列存在。
        result = result.drop_duplicates(subset=cols)

    if params.get('groupby'):
        group_cols = [c.strip() for c in params['groupby'].split(',')]
        # _validate_columns 已保证所有 groupby 列存在。
        if group_cols and params.get('agg'):
            agg_dict = parse_agg(params['agg'])
            # A-3 P0-1: 命名聚合 kwargs 模式让 groupby 产出
            # `salary_mean` / `count_name` 这种 `<col>_<func>` 形式的列名，
            # 与 SKILL.md / README 文档承诺一致；之前用 agg(dict) 模式
            # 列名是 `salary`（无 `_func` 后缀），与文档不符。
            result = result.groupby(group_cols).agg(**agg_dict).reset_index()

    # A-2 P0-3: sort 列名校验推迟到 groupby+agg 之后，使用 result.columns
    # （聚合后才会出现 salary_mean / count_name 这类新列）。
    if params.get('sort'):
        sort_cols, ascending = parse_sort(params['sort'])
        _validate_sort_columns(params['sort'], result)
        result = result.sort_values(by=sort_cols, ascending=ascending)

    # offset / limit 必须用 is not None 判断：0 是合法值（offset=0 等于不偏移，limit=0 要求 0 行）
    if params.get('offset') is not None:
        result = result.iloc[int(params['offset']):]

    if params.get('limit') is not None:
        result = result.head(int(params['limit']))

    return result


def _apply_pipeline_safe(df, params):
    """A-2 P0-2: 交互模式包装层。把 _apply_pipeline 的异常转为 (df, error) 元组，
    让交互循环能打印错误并继续接收下一条命令，而不是崩出整个 --interactive 会话。
    错误分类语义与 query_excel 入口一致：用户错误 vs 内部错误（带 traceback）。"""
    try:
        return _apply_pipeline(df, params), None
    except (ValueError, KeyError, FileNotFoundError) as e:
        return None, f"参数或输入错误: {e}"
    except Exception as e:
        # 内部错误：保留 traceback 供调试
        import traceback
        return None, f"内部错误: {e}\n{traceback.format_exc()}"


def _apply_query(df, state):
    """交互模式入口：把 state 灌进统一管线，异常包装为 (df, error) 元组。"""
    return _apply_pipeline_safe(df, state)


def parse_where(expr, df):
    """解析 where 条件表达式为 pandas mask。

    设计要点：
    - 括号优先：先把最内层括号求值并替换为占位符；
    - and / or 用 reduce 合并所有段，避免 N≥3 时丢失尾部条件；
    - 占位符与 df 解耦：放在 masks 字典里，不向 df 写列，杜绝副作用。
    """
    return _parse_where(expr, df, {})


def _parse_where(expr, df, masks):
    """递归解析 where 条件。masks 存储括号展开后的中间 mask 与字符串字面量占位符。"""
    if not expr:
        return pd.Series([True] * len(df), index=df.index)

    expr = expr.strip()

    # A-2 P0-1: 字符串字面量保护。先识别 '...' / "..." 用占位符替换，
    # 避免 _PRECEDENCE 切 ' or ' / ' and ' 时把字符串内的 ' or ' / ' and ' 错切
    # （典型场景：name == 'a or b'、name contains 'foo and bar'）。
    # 占位符与括号占位符复用同一个 masks 字典；解析时一并还原。
    str_re = re.compile(r"'[^']*'|\"[^\"]*\"")
    for m in reversed(list(str_re.finditer(expr))):
        placeholder = f"__STR_{len(masks)}__"
        masks[placeholder] = m.group(0)
        expr = expr[:m.start()] + placeholder + expr[m.end():]

    # 括号优先：把最内层括号替换为占位符
    while '(' in expr:
        match = re.search(r'\(([^()]+)\)', expr)
        if not match:
            break
        inner_mask = _parse_where(match.group(1), df, masks)
        placeholder = f"__MASK_{len(masks)}__"
        masks[placeholder] = inner_mask
        expr = (expr[:match.start()] + placeholder + expr[match.end():]).strip()

    # and / or：显式优先级表，按 _PRECEDENCE 顺序切分。
    # 列表顺序即优先级从高到低（高优先级先切 = 绑定更松）。
    # 与 Python / pandas.DataFrame.query() 默认语义一致。
    # A-3 P1-3: 大小写不敏感，AND/OR/And/Or 等都按同一运算符处理。
    for op_token, _, op_func in _PRECEDENCE:
        op_pattern = re.compile(re.escape(op_token), re.IGNORECASE)
        if op_pattern.search(expr):
            parts = op_pattern.split(expr)
            part_masks = [_parse_where(p.strip(), df, masks) for p in parts]
            return reduce(op_func, part_masks)

    # 占位符（括号展开后的中间结果）
    if expr in masks:
        return masks[expr]

    # 运算符优先级：多词、带空格的运算符必须先于单字符比较符匹配，
    # 避免 `name contains '>=10'` 被误切为 `name contains '` 与 `'10'`。
    for op in [' in ', ' contains ', ' startswith ', ' endswith ',
               '==', '!=', '>=', '<=', '>', '<']:
        if op in expr:
            left, right = expr.split(op, 1)
            left = left.strip().strip('"\'')
            right = right.strip()
            # A-2 P0-1: 还原字符串字面量占位符；占位符不在 masks 时
            # 退回剥离外层引号的旧行为（保持向后兼容）。
            if right in masks:
                right = masks[right]
            else:
                right = right.strip('"\'')
            if left not in df.columns:
                avail = ', '.join(f'{c!r}' for c in df.columns)
                raise KeyError(
                    f"where 引用了不存在的列: {left!r}；可用列: [{avail}]"
                )
            col = df[left]
            right_val = _parse_value(right)

            if op == '==':
                return col == right_val
            elif op == '!=':
                return col != right_val
            elif op == '>':
                return col > right_val
            elif op == '<':
                return col < right_val
            elif op == '>=':
                return col >= right_val
            elif op == '<=':
                return col <= right_val
            elif op == ' in ':
                if isinstance(right_val, str):
                    right_val = ast.literal_eval(right_val)
                return col.isin(right_val)
            elif op == ' contains ':
                return col.astype(str).str.contains(str(right_val), na=False)
            elif op == ' startswith ':
                return col.astype(str).str.startswith(str(right_val), na=False)
            elif op == ' endswith ':
                return col.astype(str).str.endswith(str(right_val), na=False)

    raise ValueError(f"无法解析的条件表达式: {expr}")


def _parse_value(val):
    """尝试将字符串解析为 Python 字面量。"""
    try:
        return ast.literal_eval(val)
    except (ValueError, SyntaxError):
        return val


def _validate_columns(params, df):
    """统一校验 select / exclude / sort / distinct / groupby / agg 引用的列是否存在。

    任意一个参数引用了不存在的列，立即抛 `ValueError`，错误信息包含参数名、
    缺失列名与可用列清单；不再 silent drop 也不让用户走完管线后看到空结果。
    """
    valid = set(df.columns)

    def _check_list(key):
        val = params.get(key)
        if not val:
            return
        cols = [c.strip() for c in val.split(',') if c.strip()]
        missing = [c for c in cols if c not in valid]
        if missing:
            avail = ', '.join(f'{c!r}' for c in valid)
            raise ValueError(
                f"--{key} 引用了不存在的列: {missing}；可用列: [{avail}]"
            )

    for k in ('select', 'exclude', 'distinct'):
        _check_list(k)

    # groupby 单独走一次
    _check_list('groupby')

    # A-2 P0-3: sort 列名校验推迟到 groupby+agg 之后。
    # 原因：聚合后才会出现 `salary_mean` / `count_name` 这类新列，
    # 用原始 df.columns 校验会误拒绝合法 sort。sort 校验逻辑搬到
    # `_validate_sort_columns` 与管线 groupby+agg 之后。

    # agg 的 col 字段也必须存在
    agg_val = params.get('agg')
    if agg_val:
        # parse_agg 返回 {output_name: pd.NamedAgg(col, func)}，
        # 原始 col 在 NamedAgg.column；校验 DataFrame 是否有这些 col。
        agg_dict = parse_agg(agg_val)
        agg_cols = [na.column for na in agg_dict.values()]
        missing = [c for c in agg_cols if c not in valid]
        if missing:
            avail = ', '.join(f'{c!r}' for c in valid)
            raise ValueError(
                f"--agg 引用了不存在的列: {missing}；可用列: [{avail}]"
            )


def parse_sort(sort_str):
    """解析排序字符串，如 'salary desc,age asc'。"""
    if not sort_str:
        return [], []
    cols = []
    orders = []
    for part in sort_str.split(','):
        part = part.strip()
        if not part:
            continue
        tokens = part.split()
        col = tokens[0]
        asc = True
        if len(tokens) > 1 and tokens[1].lower() in ('desc', 'descending'):
            asc = False
        cols.append(col)
        orders.append(asc)
    return cols, orders


def _validate_sort_columns(sort_str, df):
    """A-2 P0-3: sort 列名校验推迟到 groupby+agg 之后，使用 result.columns。

    与 `_validate_columns` 内被搬走的 sort 校验逻辑等价，但接受外部 df
    参数（聚合后的 DataFrame），让 sort 引用 `salary_mean` / `count_name`
    这类新列成为合法。
    """
    if not sort_str:
        return
    valid = set(df.columns)
    sort_cols = []
    for part in sort_str.split(','):
        part = part.strip()
        if not part:
            continue
        tokens = part.split()
        sort_cols.append(tokens[0])
    missing = [c for c in sort_cols if c not in valid]
    if missing:
        avail = ', '.join(f'{c!r}' for c in valid)
        raise ValueError(
            f"--sort 引用了不存在的列: {missing}；可用列: [{avail}]"
        )


def parse_agg(agg_str):
    """解析聚合字符串，如 'sum:salary,count:name,mean:age'。

    返回**命名聚合 kwargs dict**：`{f"{col}_{func}": pd.NamedAgg(column=col, aggfunc=func)}`。
    配套 `_apply_pipeline` 的 `df.groupby(...).agg(**agg_dict)` 会产出
    `salary_mean` / `count_name` 这类与 SKILL.md / README 文档承诺一致的
    列名（之前用 `agg(dict)` 模式时列名是 `salary`，与文档不符——A-3 P0-1）。

    校验：每段必须包含 `:`，且 func 与 col 都非空；空 func / 空 col 视为
    畸形条目并抛 `ValueError`，避免后续 groupby 拿到垃圾字典。
    """
    if not agg_str:
        return {}
    result = {}
    for part in agg_str.split(','):
        part = part.strip()
        if not part:
            continue
        if ':' not in part:
            raise ValueError(f"agg 条目缺少 ':' 分隔符: {part!r}")
        func, col = part.split(':', 1)
        func = func.strip()
        col = col.strip()
        if not func or not col:
            raise ValueError(f"agg 条目 func 或 col 为空: {part!r}")
        # A-1 P1-2: 检测重复 col；之前字典赋值会静默覆盖，多半是用户笔误。
        # 改成命名聚合后，重复检测改为"重复输出列名"以匹配实际行为。
        out_name = f"{col}_{func}"
        if out_name in result:
            raise ValueError(
                f"agg 重复列: {out_name!r}（已映射到 {result[out_name]!r}，又指定 {func!r}）"
            )
        result[out_name] = pd.NamedAgg(column=col, aggfunc=func)
    return result


def query_excel(input_path, sheet=None, header=0, where=None, select=None,
                sort=None, distinct=None, groupby=None, agg=None,
                limit=None, offset=None, output=None, tag=None, fmt='csv',
                preview=None, interactive=None, exclude=None):
    """执行 Excel 查询并输出 CSV，或预览数据，或进入交互模式。"""

    if not Path(input_path).exists():
        return {"error": f"文件不存在: {input_path}"}

    # B-2: CLI `--output` 与交互模式 `save <path>` 同语义；`save` 已拦截 `..` 越界
    # 写父目录（A-1 P0-1 修复），但 CLI `--output` 留了一个口子。`--output ../escape.csv`
    # 会越界到父目录写文件。在入口处对称拦截，错误信息走"参数或输入错误"分类。
    if output and '..' in Path(output).parts:
        return {"error": f"参数或输入错误: --output 路径不允许包含 '..': {output!r}"}

    # A-3 P2-3: 运行时只读保护 - 入口对 input_path 做 stat 快照，
    # 调度结束后再 stat 一次；若 mtime / size 改变则抛 RuntimeError，
    # 强制退出码 2（区别于普通错误码 1）。覆盖未来的"误把输出写回源文件"代码路径。
    input_path_obj = Path(input_path)
    try:
        _stat_before = input_path_obj.stat()
    except OSError as e:
        return {"error": f"无法读取源文件状态: {e}"}

    # 整个 read → preview/interactive/_query_single 流程统一 try/except，
    # 保证下游错误（KeyError、空 mask 解析、_query_single 抛错）都转为 {"error": ...}，
    # 避免 Python traceback 一路暴露到 CLI。
    # A-1 P0-2: 异常分类上提到入口；用户错误归"参数或输入错误"，
    # 其余内部错误附 traceback（与 _apply_pipeline 已有口径保持一致）。
    # LookupError 在 pandas 实际抛出场景里几乎不存在，参见 A-1 P1-1 移除。
    try:
        df = _read_input(input_path, sheet=sheet, header=header)
    except (ValueError, KeyError, FileNotFoundError) as e:
        return {"error": f"参数或输入错误: {e}"}
    except Exception as e:
        import traceback
        return {"error": f"内部错误: 读取输入失败: {e}\n{traceback.format_exc()}"}

    try:
        result = _query_excel_dispatch(df, input_path, sheet, where, select, sort,
                                       distinct, groupby, agg, limit, offset,
                                       output, tag, fmt, preview, interactive, exclude)
    except (ValueError, KeyError, FileNotFoundError) as e:
        return {"error": f"参数或输入错误: {e}"}
    except Exception as e:
        import traceback
        return {"error": f"内部错误: 查询失败: {e}\n{traceback.format_exc()}"}

    # 出口 stat 校验
    try:
        _stat_after = input_path_obj.stat()
        if (_stat_before.st_mtime_ns != _stat_after.st_mtime_ns
                or _stat_before.st_size != _stat_after.st_size):
            return {
                "error": f"运行时只读保护：源文件 {input_path} 在查询过程中被修改 "
                         f"(mtime/size 变化)，已中止。请检查脚本是否有写源文件的代码路径。",
                "_read_only_violation": True,
            }
    except OSError as e:
        return {"error": f"无法校验源文件状态: {e}"}

    return result


def _query_excel_dispatch(df, input_path, sheet, where, select, sort,
                          distinct, groupby, agg, limit, offset,
                          output, tag, fmt, preview, interactive, exclude):
    """query_excel 的主调度逻辑：从读取后到结果返回。"""

    # sheet_name=None 时 pandas 总是返回 dict（即使文件只有 1 张表），
    # 这里 unwrap：单 sheet dict 等同于单 DataFrame，多 sheet dict 走分片分支。
    if isinstance(df, dict) and sheet is None and len(df) == 1:
        df = next(iter(df.values()))

    # 多 sheet 处理
    if isinstance(df, dict):
        # A-2 P1-2: 用 `is not None` 判 `--preview` 而非 `if preview:`，
        # 避免 `--preview 0` 落入 Python falsy 陷阱（应走 preview 分支预览 0 行，
        # 而不是走 _query_single 写文件）。A-1 已对 `--limit` / `--offset` 负数
        # 做入口校验；这里把 falsy 边界补齐。
        if preview is not None:
            # 多 sheet + --json --preview：每个 sheet 仍走完整 preview，
            # 但每个 sheet 的 preview 文本只保留 5 行样本 + 字段信息，
            # 避免 JSON 输出体量爆炸。完整 preview 仍可通过普通 CLI 输出查看。
            results = {}
            for sheet_name, sheet_df in df.items():
                results[sheet_name] = preview_data(sheet_df, n=5)
            return {"success": True, "preview": results}

        if interactive:
            # 交互模式下，对每个 sheet 分别进入交互；本轮先在 banner 提示当前 sheet
            # 与剩余 sheet 数量，UX 优化留待后续轮次做完整 sheet 切换状态机。
            sheet_list = list(df.items())
            for idx, (sheet_name, sheet_df) in enumerate(sheet_list, 1):
                print(f"\n{'#' * 60}")
                print(f"# 工作表: {sheet_name}（{idx}/{len(sheet_list)}，输入 'quit' 进入下一张）")
                print(f"{'#' * 60}")
                interactive_mode(input_path, sheet_df, sheet_name, fmt)
            return {"success": True}

        results = {}
        for sheet_name, sheet_df in df.items():
            out_path = _query_single(sheet_df, input_path, sheet_name,
                                     where, select, sort, distinct,
                                     groupby, agg, limit, offset, output, tag, fmt, exclude)
            results[sheet_name] = out_path
        return {"success": True, "output_files": results}

    # 单 sheet 处理
    # A-2 P1-2: 同上，`--preview 0` 仍应走 preview 分支（输出 0 行 + 字段元信息），
    # 不应被 `if 0:` 当成 False 跳过而落进 _query_single 写文件。
    if preview is not None:
        return {"success": True, "preview": preview_data(df, n=preview)}

    if interactive:
        interactive_mode(input_path, df, fmt=fmt)
        return {"success": True}

    out_path = _query_single(df, input_path, None,
                             where, select, sort, distinct,
                             groupby, agg, limit, offset, output, tag, fmt, exclude)
    return {"success": True, "output_file": out_path}


def _query_single(df, input_path, sheet_name,
                  where, select, sort, distinct,
                  groupby, agg, limit, offset, output, tag=None, fmt='csv', exclude=None):
    """处理单个 DataFrame 的查询逻辑：先调统一管线得到 result，再决定输出路径。"""
    params = {
        'where': where,
        'exclude': exclude,
        'select': select,
        'distinct': distinct,
        'groupby': groupby,
        'agg': agg,
        'sort': sort,
        'offset': offset,
        'limit': limit,
    }
    # A-2 P0-2: _apply_pipeline 现在直接 raise 原始异常，不再返回 (df, error) 元组。
    # 错误分类（用户错误 vs 内部错误 + traceback）由外层 query_excel 的
    # except (ValueError, KeyError, FileNotFoundError) / except Exception 双分支
    # 负责。_query_single 不再重抛 RuntimeError，避免双层包装。
    result = _apply_pipeline(df, params)

    # 确定输出路径
    safe_sheet = _sanitize_sheet_name(sheet_name) if sheet_name else None
    if output:
        # --output 显式给出 + 多 sheet：按 sheet 名分片到相邻文件，
        # 避免所有 sheet 顺序写入同一文件、最后一个静默覆盖前面的结果。
        out_path = Path(output)
        if safe_sheet:
            out_path = Path(str(out_path.with_suffix('')) + f'_{safe_sheet}{out_path.suffix}')
    else:
        suffix = '.xlsx' if fmt == 'xlsx' else '.csv'
        base = Path(input_path).with_suffix('')
        if safe_sheet:
            out_path = Path(str(base) + f'_{safe_sheet}{suffix}')
        else:
            out_path = base.with_suffix(suffix)

        # 如果有 tag，追加到文件名
        if tag:
            safe_tag = _validate_tag(tag)
            out_path = out_path.with_suffix('')
            out_path = Path(str(out_path) + f'_{safe_tag}{suffix}')

        # 自动防覆盖（仅默认命名时）
        out_path = _ensure_unique(out_path)

    # 写入输出文件
    _write_output(result, out_path, fmt)
    return str(out_path)


def main():
    parser = argparse.ArgumentParser(description='Excel 查询工具')
    parser.add_argument('--version', action='version', version=f'%(prog)s {__version__}')
    parser.add_argument('-f', '--file', required=True, help='输入 Excel 文件路径')
    parser.add_argument('--sheet', help='指定工作表名或索引')
    parser.add_argument('--header', type=int, default=0, help='表头所在行号，0-based（默认0，若表头在第2行则传1）')
    parser.add_argument('--where', help='筛选条件，如 "age > 30"')
    parser.add_argument('--select', help='选择列，逗号分隔')
    parser.add_argument('--sort', help='排序规则，如 "salary desc,age asc"')
    parser.add_argument('--distinct', help='去重列，逗号分隔')
    parser.add_argument('--exclude', help='排除列，逗号分隔（与 --select 互斥，优先执行排除）')
    parser.add_argument('--groupby', help='分组列，逗号分隔')
    parser.add_argument('--agg', help='聚合规则，如 "sum:salary,count:name"')
    parser.add_argument('--limit', type=int, help='限制返回行数')
    parser.add_argument('--offset', type=int, help='跳过前 N 行')
    parser.add_argument('-o', '--output', help='输出路径（默认同目录同名，扩展名由 --format 决定）')
    parser.add_argument('--tag', help='输出文件名标签，如 "EastHigh" 会生成 sales_EastHigh.csv')
    parser.add_argument('--format', choices=['csv', 'xlsx'], default='csv',
                        help='输出格式：csv（默认）或 xlsx')
    parser.add_argument('--json', action='store_true', help='以 JSON 格式输出结果')
    parser.add_argument('-p', '--preview', type=int, nargs='?', const=10, default=None,
                        help='预览数据：打印前 N 行和字段信息（默认 10 行）')
    parser.add_argument('-i', '--interactive', action='store_true',
                        help='进入交互模式：先预览数据，然后逐条输入查询条件')

    args = parser.parse_args()

    # --json 与 --interactive 互斥：交互模式会从 stdin 读取 prompt，
    # 与 JSON 结构化输出会互相污染 stdout。早期拦截比返回部分 JSON 更友好。
    if args.json and args.interactive:
        parser.error("--json 与 --interactive 互斥；如需 JSON 输出请使用 --json 单次查询")

    result = query_excel(
        input_path=args.file,
        sheet=args.sheet,
        header=args.header,
        where=args.where,
        select=args.select,
        sort=args.sort,
        distinct=args.distinct,
        groupby=args.groupby,
        agg=args.agg,
        limit=args.limit,
        offset=args.offset,
        output=args.output,
        tag=args.tag,
        fmt=args.format,
        preview=args.preview,
        interactive=args.interactive,
        exclude=args.exclude,
    )

    if args.json:
        # --json 模式：统一走 JSON；错误也以 {"error": ...} 形式返回 stdout
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        # 非 --json 模式：成功信息走 stdout，错误也走 stderr 让 shell 区分。
        # 这是有意的双通道；--json 才是结构化输出的权威。
        if 'error' in result:
            print(f"错误: {result['error']}", file=sys.stderr)
            sys.exit(1)
        if 'output_file' in result:
            print(f"输出: {result['output_file']}")
        elif 'output_files' in result:
            for sheet, path in result['output_files'].items():
                print(f"[{sheet}] 输出: {path}")
        elif 'preview' in result:
            if isinstance(result['preview'], dict):
                for sheet, preview_text in result['preview'].items():
                    print(f"\n{'#' * 60}")
                    print(f"# 工作表: {sheet}")
                    print(f"{'#' * 60}")
                    print(preview_text)
            else:
                print(result['preview'])

    sys.exit(0 if result.get('success') else 1)


if __name__ == '__main__':
    main()
