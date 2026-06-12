#!/usr/bin/env python3
"""
两表去重合并脚本：
- 支持 CSV / XLSX 输入（可混合格式）
- 自动检测或用户指定关键匹配列
- 精确匹配 + 模糊匹配（可开关、可选预设）
- 合并同名记录，保留两表全部列；冲突列名自动加后缀
- 输出格式跟随输入（CSV→CSV，XLSX→XLSX，混合格式→XLSX）
- 关键列空值不参与匹配（保留为 left_only / right_only）
"""

import argparse
import difflib
import re
import sys
from pathlib import Path
from typing import cast

try:
    import pandas as pd
except ImportError as e:
    print("错误: 缺少 pandas。请安装: pip install pandas openpyxl", file=sys.stderr)
    sys.exit(1)


def read_table(path):
    """根据文件后缀读取 CSV 或 XLSX 为 DataFrame。"""
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix == '.csv':
        # A-2 Fix 4: NUL 字节在 CSV 写入时会被 pandas 默默截断 / 当作字段分隔符，
        # 导致关键值数据丢失；read 阶段在喂给 pandas 之前显式检测 raw bytes 中的
        # NUL 字节并 raise ValueError，让 main 走 stderr 单行错误退出
        # 注: 用 open(..., 'rb') 让 FileNotFoundError 透传，main 的
        # except FileNotFoundError 才能命中"不可读"消息（A-1 P1-2 兼容）
        with open(path, "rb") as _fb:
            raw_bytes = _fb.read()
        if b"\x00" in raw_bytes:
            raise ValueError(
                f"输入 CSV 包含 NUL (0x00) 字节: {path}。"
                f" 关键值含 NUL 时输出 CSV 会被 pandas 截断/错位丢数据，请先清洗源文件。"
            )
        # 编码回退：utf-8-sig → utf-8 → gb18030（简中）→ big5（繁中）→ shift_jis（日）→ euc-kr（韩）→ latin-1（单字节兜底）
        for encoding in ['utf-8-sig', 'utf-8', 'gb18030', 'big5', 'shift_jis', 'euc-kr', 'latin-1']:
            try:
                df = pd.read_csv(path, encoding=encoding, dtype=str, keep_default_na=False)
            except UnicodeDecodeError:
                continue
            return df
        raise ValueError(f"无法识别 CSV 编码: {path}")

    if suffix in ('.xlsx', '.xls', '.xlsm'):
        # A-2 Fix 3: openpyxl 在 cell 写入阶段对 NUL 等非法字符抛 IllegalCharacterError；
        # 但当源 xlsx 是手工拼装（zip + xml 注入 NUL 等控制字符）时，openpyxl XML
        # 解析器会抛 ParseError("not well-formed")。两种都表明"输入 .xlsx 含
        # openpyxl 拒绝的非法字符"，统一捕获并转 ValueError，让 main 现有
        # (ValueError, UnicodeDecodeError) 异常处理统一接管（单行 stderr 错误，
        # verbose 模式附 stack trace）
        try:
            return pd.read_excel(path, dtype=str, keep_default_na=False)
        except Exception as e:
            cls_name = type(e).__name__
            err_str = str(e).lower()
            if 'illegalcharacter' in cls_name.lower() or 'illegal character' in err_str:
                raise ValueError(
                    f"输入 .xlsx 包含 openpyxl 拒绝的非法字符 (NUL/控制字符): {path}。"
                    f" 原始错误: {e}"
                ) from e
            if cls_name == 'ParseError' or 'not well-formed' in err_str:
                raise ValueError(
                    f"输入 .xlsx 文件结构损坏或包含 openpyxl 拒绝的非法字符: {path}。"
                    f" 原始错误: {e}"
                ) from e
            raise

    raise ValueError(f"不支持的文件格式 '{suffix}'，仅支持 .csv / .xlsx / .xls / .xlsm: {path}")


def _sanitize_xlsx_value(v):
    """A-2 Fix 5: openpyxl 写入 .xlsx 时对 cell value 是 str 且以 =/+/-/@ 开头的
    加单引号前缀，防止 Excel 打开时执行公式（CSV/公式注入 / DDE 注入高危）。
    """
    if isinstance(v, str) and v and v[0] in ('=', '+', '-', '@'):
        return "'" + v
    return v


def write_table(df, path):
    """根据文件后缀输出 CSV 或 XLSX。"""
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix == '.csv':
        df.to_csv(path, index=False, encoding='utf-8-sig')
    elif suffix in ('.xlsx', '.xls', '.xlsm'):
        # .xlsm 宏由 pandas/openpyxl 写入会丢失，主动提示
        if suffix == '.xlsm':
            print("警告: 输出 .xlsm 时宏将丢失，请用 VBA/Excel 工具合并。", file=sys.stderr)
        # A-2 Fix 5: 对全部 cell value 做公式注入防护（仅 .xlsx/.xls/.xlsm 路径）
        # 注: pandas 2.x 已弃用 `DataFrame.applymap`，需用 `df.map` 或
        # `df.apply(lambda col: col.map(...))` 兼容
        sanitized = df.apply(lambda col: col.map(_sanitize_xlsx_value))
        sanitized.to_excel(path, index=False, engine='openpyxl')
    else:
        raise ValueError(f"不支持的输出格式 '{suffix}'，仅支持 .csv / .xlsx / .xls / .xlsm: {path}")


def _column_uniqueness_score(series):
    """计算列的唯一值比例（越高越适合作为关键列）。"""
    total = len(series)
    if total == 0:
        return 0.0
    unique = series.nunique(dropna=False)
    return unique / total


def _column_name_score(name):
    """根据列名语义给分（越高越可能是关键列）。

    high_keywords 覆盖中英文常见关键列命名：
    id/编号/名称/名字/全称/简称/name/title/key/标识；
    journal/publisher/author/product/sku/issn/isbn/category/subject 等。
    """
    name_lower = name.lower()
    # 高权重关键词（直接命中基本确定是关键列）
    high_keywords = [
        'id', '编号', '名称', '名字', '全称', '简称', 'name', 'title', 'key', '标识',
        'journal', 'publisher', 'author', 'product', 'product_name', 'sku',
        'issn', 'isbn', 'doi', 'category', 'subject', 'tag', 'member',
    ]
    # 低权重/负面关键词（通常不是关键列）
    low_keywords = ['网址', 'url', '网站', 'web', '地址', 'address', 'link', 'comment', '备注', 'desc', 'description']
    score = 0
    for kw in high_keywords:
        if kw in name_lower:
            score += 20
    for kw in low_keywords:
        if kw in name_lower:
            score -= 15
    return score


def _sequence_similarity(a, b):
    """基于 difflib 序列相似度，对中英文都更合理。"""
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()


def detect_key_column(df1, df2, key_hint=None, key2_hint=None, uniqueness_min=0.5):
    """自动检测两表的关键匹配列。

    策略：
    1. 若 key_hint 提供，直接使用
    2. 若 key2_hint 提供（仅指定了文件2的关键列），自动检测文件1的关键列，再用 key2_hint 锁定文件2
    3. 找两表共有的列名（忽略大小写），排除唯一值比例过低的列
    4. 在候选列中按语义得分+唯一性得分排序，取最优
    5. 若无共有列，找列名相似度最高的配对（用 difflib.SequenceMatcher）
    """
    cols1 = list(df1.columns)
    cols2 = list(df2.columns)

    if key2_hint is not None and key_hint is None:
        # 仅指定文件2关键列：自动检测文件1关键列，再用 key2_hint 锁定文件2
        if key2_hint not in cols2:
            lower_map2 = {c.lower(): c for c in cols2}
            if key2_hint.lower() in lower_map2:
                key2_hint = lower_map2[key2_hint.lower()]
            else:
                raise ValueError(f"文件2中未找到关键列 '{key2_hint}'，可用列: {cols2}")
        # 走普通共有列检测路径，结束时把 key2 覆盖为 key2_hint
        result = detect_key_column(df1, df2, key_hint=None, key2_hint=None, uniqueness_min=uniqueness_min)
        return result[0], key2_hint

    if key_hint:
        # key_hint 可能是文件1的列名，也可能是两表共有的列名
        key1 = key_hint
        if key1 not in cols1:
            # 尝试大小写不敏感匹配
            lower_map1 = {c.lower(): c for c in cols1}
            if key1.lower() in lower_map1:
                key1 = lower_map1[key1.lower()]
            else:
                raise ValueError(f"文件1中未找到关键列 '{key_hint}'，可用列: {cols1}")

        # 文件2优先找同名列，否则也尝试大小写不敏感
        if key1 in cols2:
            return key1, key1
        lower_map2 = {c.lower(): c for c in cols2}
        if key1.lower() in lower_map2:
            return key1, lower_map2[key1.lower()]
        raise ValueError(f"文件2中未找到关键列 '{key_hint}'（或大小写变体），可用列: {cols2}")

    # 找共有列（忽略大小写）
    lower1 = {c.lower(): c for c in cols1}
    lower2 = {c.lower(): c for c in cols2}
    common = set(lower1.keys()) & set(lower2.keys())

    if common:
        candidates = []
        for col_lower in common:
            col1 = lower1[col_lower]
            col2 = lower2[col_lower]
            # 唯一值比例：要求至少 50% 的值是唯一的（排除"学科"、"类别"等大量重复的列）
            uniq1 = _column_uniqueness_score(df1[col1])
            uniq2 = _column_uniqueness_score(df2[col2])
            min_uniq = min(uniq1, uniq2)
            if min_uniq < uniqueness_min:
                continue  # 排除重复值过多的列
            # 语义得分 + 唯一性得分 - 长度惩罚（鼓励较短的列名）
            # A-1 P1-1 fix: length 惩罚从 0.1 降到 0.01；name_score 差 ≥ 10 时按
            # name_score 决胜（避免长但语义强的列名被短但语义空的列名反超）
            name_score = _column_name_score(col1)
            length_penalty = len(col1) * 0.01
            total_score = name_score + min_uniq * 10 - length_penalty
            candidates.append((total_score, name_score, col1, col2))

        if candidates:
            # name_score 决胜护栏：当最高与次高 name_score 差 ≥ 10 时，按 name_score
            # 决胜，避免 length 修正让语义弱列名反超语义强列名
            name_scores = sorted({c[1] for c in candidates}, reverse=True)
            if len(name_scores) >= 2 and (name_scores[0] - name_scores[1]) >= 10:
                candidates.sort(key=lambda x: (-x[1], -x[0]))
            else:
                candidates.sort(key=lambda x: -x[0])
            return candidates[0][2], candidates[0][3]

        # A-3 P0-1 fix: 上一段已经 return，下面的 `if candidates:` 是历史残留死代码
        # （且 `candidates[0][1]` / `candidates[0][2]` 引用的是旧元组结构，已被 `(total,
        # name_score, col1, col2)` 取代）。直接进入"放宽条件再试一次"路径。

        # 如果所有共有列都被排除，放宽条件再试一次（只看语义得分）
        candidates = []
        for col_lower in common:
            col1 = lower1[col_lower]
            col2 = lower2[col_lower]
            name_score = _column_name_score(col1)
            if name_score > 0:
                # A-1 P1-1 fix: length 惩罚从 0.1 降到 0.01，与主路径保持一致
                total_score = name_score - len(col1) * 0.01
                candidates.append((total_score, col1, col2))

        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0][1], candidates[0][2]

    # 无共有列或全部排除：找最相似的列名配对（difflib 序列相似度）
    best_score = 0.0
    best_pair = None
    for c1 in cols1:
        for c2 in cols2:
            score = _sequence_similarity(c1, c2)
            if score > best_score:
                best_score = score
                best_pair = (c1, c2)

    if best_pair and best_score >= 0.5:
        return best_pair

    raise ValueError(
        f"无法自动检测关键列。两表列名无重叠或共有列均不适合作为关键列。\n"
        f"文件1列: {cols1}\n"
        f"文件2列: {cols2}\n"
        f"请使用 --key1 / --key2 手动指定。"
    )


def normalize_key(value):
    """标准化关键值用于精确匹配。

    P0-2 fix: 不再用哨兵字符串替换空值；空值（NaN/None/空串/纯空白）经本函数返回空字符串，
    是否参与匹配由调用方用 `_is_null_key` mask 决定。这样即使真实关键值恰好是
    `__NULL_KEY_SENTINEL__` 字符串，也不会被误判为空。
    """
    if value is None:
        return ''
    if isinstance(value, float) and pd.isna(value):
        return ''
    s = str(value).strip()
    if not s:
        return ''
    s = re.sub(r'\s+', ' ', s)  # 压缩连续空格
    s = s.rstrip('.')  # 去除末尾句点
    # 不对关键值整体 upper（保护数字前导零、符号）
    return s


def _is_null_key(value):
    """判定关键值是否为空（NaN/None/空串/纯空白）；用于 merge_tables 的 mask 过滤。

    显式返回 ``bool``，避免上游 ``.apply`` 推断成 ``object`` dtype 后
    ``Series.sum()`` 退化为字符串拼接而崩溃（A-1 P0-1 fix）。
    """
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    s = str(value).strip()
    return bool(not s)


def fuzzy_normalize(value, preset='academic'):
    """更宽松的标准化，用于模糊匹配兜底。

    P0-4 fix: academic 预设的 `THE ` 前缀与 `JOURNALS OF` 规则改为大小写不敏感：
    在归一化副本上做 `.upper()` 后再应用替换；原文（merge_tables 输出 key1 时仍用原文）
    保留用户原始大小写与前导零。

    P1-3 fix: 'none' 预设由外层守卫（`fuzzy and fuzzy_preset != 'none'`）拦截，
    函数内部不再保留 `none` 分支，避免误导。

    A-3 P2-1 fix: 移除历史"防御性 raise"分支；`preset='none'` 时直接返回
    `normalize_key(value)`（与 `--no-fuzzy` 等价），让函数本身可独立测试。
    """
    if preset == 'none':
        # A-3 P2-1 fix: 防御 raise 改为直通 `normalize_key`；外层不再需要守卫
        return normalize_key(value)
    s = normalize_key(value)
    # 所有预设：连字符变空格、AND→&（与大小写无关）
    s = re.sub(r'[\-–—]', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    s = s.replace(' AND ', ' & ')
    if preset == 'cjk':
        # 仅做连字符与 AND 替换（大小写敏感，与原文一致）
        return s.strip()
    # academic（默认）：THE 前缀、JOURNALS OF→JOURNAL OF（大小写不敏感）
    norm = s.upper()
    norm = norm.replace('JOURNALS OF', 'JOURNAL OF')
    if norm.startswith('THE '):
        norm = norm[4:]
    return norm.strip()


def _deduplicate_columns(df):
    """去重 DataFrame 中的重复列名，保留首次出现。"""
    seen = set()
    keep = []
    for col in df.columns:
        if col not in seen:
            seen.add(col)
            keep.append(col)
    return df[keep]


def merge_tables(df1, df2, key1, key2, fuzzy=True, fuzzy_preset='academic',
                 sort_by_key=False):
    """按关键列合并两个 DataFrame，返回合并结果和统计信息。

    关键列空值不参与匹配：merge 前从两表剔除空值行，并在统计中独立报告。
    """
    df1 = df1.copy()
    df2 = df2.copy()

    # A-2 Fix 7: 关键值内部重复时 pandas merge 按 Cartesian 展开（pandas 默认行为），
    # 统计与预览可能误导；显式打印警告让用户立刻看到
    norm1 = df1[key1].apply(normalize_key)
    norm2 = df2[key2].apply(normalize_key)
    dup1 = int(norm1[norm1 != ''].duplicated(keep=False).sum())
    dup2 = int(norm2[norm2 != ''].duplicated(keep=False).sum())
    if dup1 > 0:
        print(
            f"警告: 关键列 '{key1}' 在文件1中有 {dup1} 个重复值（非空），"
            f" 匹配结果按 Cartesian 展开（pandas 默认行为）。",
            file=sys.stderr,
        )
    if dup2 > 0:
        print(
            f"警告: 关键列 '{key2}' 在文件2中有 {dup2} 个重复值（非空），"
            f" 匹配结果按 Cartesian 展开（pandas 默认行为）。",
            file=sys.stderr,
        )

    # P0-2 fix: 用 mask 跟踪空值行，不再用哨兵字符串
    # A-1 P0-1 fix: 显式 astype(bool) 防止 apply 返回 object dtype 导致 sum() 字符串拼接
    df1['__zm_dedup_merge_is_null_key'] = df1[key1].apply(_is_null_key).astype(bool)
    df2['__zm_dedup_merge_is_null_key'] = df2[key2].apply(_is_null_key).astype(bool)

    # 创建标准化关键列用于匹配（空值经 normalize_key 返回空串）
    df1['__zm_dedup_merge_key_norm'] = df1[key1].apply(normalize_key)
    df2['__zm_dedup_merge_key_norm'] = df2[key2].apply(normalize_key)

    # 统计空值行数
    null_count_1 = int(df1['__zm_dedup_merge_is_null_key'].sum())
    null_count_2 = int(df2['__zm_dedup_merge_is_null_key'].sum())

    # 剔除空值行参与匹配；同时丢掉内部 mask 列避免泄漏到 merge 结果的列名
    df1_for_merge = df1[~df1['__zm_dedup_merge_is_null_key']].drop(columns=['__zm_dedup_merge_is_null_key']).copy()
    df2_for_merge = df2[~df2['__zm_dedup_merge_is_null_key']].drop(columns=['__zm_dedup_merge_is_null_key']).copy()

    # 为避免 key 列在 merge 时产生冲突后缀，临时重命名
    df1_for_merge = df1_for_merge.rename(columns={key1: '__zm_dedup_merge_key_left'})
    df2_for_merge = df2_for_merge.rename(columns={key2: '__zm_dedup_merge_key_right'})

    # 收集非关键列用于冲突检测
    other_cols1 = [c for c in df1_for_merge.columns if c not in ('__zm_dedup_merge_key_left', '__zm_dedup_merge_key_norm')]
    other_cols2 = [c for c in df2_for_merge.columns if c not in ('__zm_dedup_merge_key_right', '__zm_dedup_merge_key_norm')]

    # P1-1 fix: 冲突列名后缀用 `_<N>` 顺序枚举；left 与 right 各自独立计数，
    # 避免累积下划线（`_1` → `__` → `___`）产生 `a_1_1_` / `a_1_2_` 等丑陋列名。
    # left 后缀从 _1 开始，right 后缀在 left 已用过的 N 之后继续编号。
    all_other = other_cols1 + other_cols2
    used_suffixes: set = set()

    def _next_suffix(cols):
        n = 1
        while any(c.endswith(f"_{n}") for c in cols) or n in used_suffixes:
            n += 1
        used_suffixes.add(n)
        return f"_{n}"

    # 先确定 left 后缀，再确定 right 后缀（确保两者不同）
    suffix1 = _next_suffix(all_other)
    suffix2 = _next_suffix(all_other)

    # 重命名可能冲突的非关键列
    rename1 = {c: f"{c}{suffix1}" for c in other_cols1 if c in other_cols2}
    rename2 = {c: f"{c}{suffix2}" for c in other_cols2 if c in other_cols1}

    if rename1:
        df1_for_merge = df1_for_merge.rename(columns=rename1)
    if rename2:
        df2_for_merge = df2_for_merge.rename(columns=rename2)

    # 精确匹配
    # A-1 P1-5 fix: 显式 indicator 列名与代码引用一致（避免与用户列名冲突）
    merged = pd.merge(
        df1_for_merge, df2_for_merge, on='__zm_dedup_merge_key_norm', how='outer',
        indicator='__zm_dedup_merge_merge',
        suffixes=(suffix1, suffix2)
    )
    # P0-1 fix: 精确匹配阶段所有行都不是 fuzzy；提前初始化避免后续 pd.concat 留下 NaN
    # 导致 'if is_fuzzy:' 对 NaN 求值为 True 误标精确匹配
    merged['__zm_dedup_merge_fuzzy_match'] = False

    exact_both = int((merged['__zm_dedup_merge_merge'] == 'both').sum())
    exact_left = int((merged['__zm_dedup_merge_merge'] == 'left_only').sum())
    exact_right = int((merged['__zm_dedup_merge_merge'] == 'right_only').sum())

    fuzzy_both = 0

    if fuzzy and fuzzy_preset != 'none':
        left_cols = [c for c in df1_for_merge.columns if c != '__zm_dedup_merge_key_norm']
        right_cols = [c for c in df2_for_merge.columns if c != '__zm_dedup_merge_key_norm']

        left_unmatched = cast(
            pd.DataFrame,
            merged[merged['__zm_dedup_merge_merge'] == 'left_only'][left_cols].copy(),
        )
        right_unmatched = cast(
            pd.DataFrame,
            merged[merged['__zm_dedup_merge_merge'] == 'right_only'][right_cols].copy(),
        )

        if not left_unmatched.empty and not right_unmatched.empty:
            left_unmatched['__zm_dedup_merge_key_fuzzy'] = left_unmatched['__zm_dedup_merge_key_left'].apply(
                lambda v: fuzzy_normalize(v, preset=fuzzy_preset)
            )
            right_unmatched['__zm_dedup_merge_key_fuzzy'] = right_unmatched['__zm_dedup_merge_key_right'].apply(
                lambda v: fuzzy_normalize(v, preset=fuzzy_preset)
            )

            fuzzy_merged = cast(
                pd.DataFrame,
                pd.merge(
                    left_unmatched, right_unmatched,
                    on='__zm_dedup_merge_key_fuzzy', how='outer',
                    indicator='__zm_dedup_merge_merge',
                    suffixes=(suffix1, suffix2),
                ),
            )

            fuzzy_both_count = (fuzzy_merged['__zm_dedup_merge_merge'] == 'both').sum()

            if fuzzy_both_count > 0:
                fuzzy_both = int(fuzzy_both_count)
                fuzzy_norms_set: set = set(
                    fuzzy_merged[fuzzy_merged['__zm_dedup_merge_merge'] == 'both']['__zm_dedup_merge_key_fuzzy'].tolist()
                )

                # 向量化：先 map 一列再 isin
                left_only_mask = cast(pd.Series, merged['__zm_dedup_merge_merge']) == 'left_only'
                right_only_mask = cast(pd.Series, merged['__zm_dedup_merge_merge']) == 'right_only'
                left_fuzzy = cast(
                    pd.Series, merged['__zm_dedup_merge_key_left']
                ).map(lambda v: fuzzy_normalize(v, preset=fuzzy_preset))
                right_fuzzy = cast(
                    pd.Series, merged['__zm_dedup_merge_key_right']
                ).map(lambda v: fuzzy_normalize(v, preset=fuzzy_preset))

                keep = ~(left_only_mask & left_fuzzy.isin(fuzzy_norms_set)) & ~(
                    right_only_mask & right_fuzzy.isin(fuzzy_norms_set)
                )
                merged = cast(pd.DataFrame, merged[keep].copy())

                fuzzy_both_rows = fuzzy_merged[fuzzy_merged['__zm_dedup_merge_merge'] == 'both'].drop(
                    columns=['__zm_dedup_merge_key_fuzzy', '__zm_dedup_merge_merge']
                )
                fuzzy_both_rows['__zm_dedup_merge_merge'] = 'both'
                fuzzy_both_rows['__zm_dedup_merge_fuzzy_match'] = True
                merged = cast(
                    pd.DataFrame,
                    pd.concat([merged, fuzzy_both_rows], ignore_index=True),
                )

                exact_left = int((merged['__zm_dedup_merge_merge'] == 'left_only').sum())
                exact_right = int((merged['__zm_dedup_merge_merge'] == 'right_only').sum())

    # 合并后的关键列：取非空值（优先左表/文件1）
    key_left_series = cast(pd.Series, merged['__zm_dedup_merge_key_left'])
    key_right_series = cast(pd.Series, merged['__zm_dedup_merge_key_right'])
    merged[key1] = key_left_series.fillna(key_right_series)

    # 构造 match_type 列（exact / fuzzy / left_only / right_only）
    # 精确阶段已在 P0-1 修复中预置 _fuzzy_match=False；fuzzy 阶段只把新增的 both 标 True
    match_type: list = []
    for merge_val, is_fuzzy in zip(
        cast(pd.Series, merged['__zm_dedup_merge_merge']).tolist(),
        cast(pd.Series, merged['__zm_dedup_merge_fuzzy_match']).tolist(),
    ):
        if merge_val == 'both':
            match_type.append('fuzzy' if is_fuzzy else 'exact')
        elif merge_val == 'left_only':
            match_type.append('left_only')
        elif merge_val == 'right_only':
            match_type.append('right_only')
        else:
            match_type.append(str(merge_val))
    merged['match_type'] = match_type

    # 清理辅助列
    for col in ['__zm_dedup_merge_key_left', '__zm_dedup_merge_key_right', '__zm_dedup_merge_key_norm', '__zm_dedup_merge_merge', '__zm_dedup_merge_key_fuzzy', '__zm_dedup_merge_fuzzy_match']:
        if col in merged.columns:
            merged = merged.drop(columns=[col])

    # 去重列名
    merged = _deduplicate_columns(merged)

    # 补回空值行：作为 left_only / right_only 直接进入结果
    if null_count_1 > 0:
        null_rows_1 = df1[df1['__zm_dedup_merge_is_null_key']].drop(columns=['__zm_dedup_merge_key_norm', '__zm_dedup_merge_is_null_key']).copy()
        # 重命名 key1 列保持结果列名一致
        null_rows_1 = null_rows_1.rename(columns={key1: key1})
        null_rows_1['match_type'] = 'null_key_1'
        merged = pd.concat([merged, null_rows_1], ignore_index=True, sort=False)
    if null_count_2 > 0:
        # P0-3 fix: 右表非关键列统一追加 _2 后缀（不再依赖 `c in merged.columns` 漏检），
        # 然后用 _deduplicate_columns 兜底收敛列名
        null_rows_2 = df2[df2['__zm_dedup_merge_is_null_key']].drop(columns=['__zm_dedup_merge_key_norm', '__zm_dedup_merge_is_null_key']).copy()
        null_rows_2 = null_rows_2.rename(columns={key2: key1})  # 与左表统一为 key1
        # 右表所有非 key1 列统一追加 suffix2 后缀
        rename_null_2 = {
            c: f"{c}{suffix2}"
            for c in list(null_rows_2.columns)
            if c != key1
        }
        if rename_null_2:
            null_rows_2 = null_rows_2.rename(columns=rename_null_2)
        null_rows_2['match_type'] = 'null_key_2'
        merged = pd.concat([merged, null_rows_2], ignore_index=True, sort=False)
        # 兜底：concat 后再做一次列名收敛，避免出现 `V` 与 `V_2` 共存的孤立列
        merged = _deduplicate_columns(merged)

    # 默认不排序；用户显式 --sort-by-key 时按关键列大写排序
    if sort_by_key:
        merged = merged.sort_values(  # type: ignore[call-overload]
            by=key1, key=lambda s: s.astype(str).str.upper(), na_position='last'
        ).reset_index(drop=True)
    else:
        merged = merged.reset_index(drop=True)  # type: ignore[attr-defined]

    stats = {
        'total': len(merged),
        'both': int(exact_both + fuzzy_both),
        'fuzzy': int(fuzzy_both),
        'left_only': int(exact_left + null_count_1),
        'right_only': int(exact_right + null_count_2),
        'null_key_1': null_count_1,
        'null_key_2': null_count_2,
    }

    return merged, stats


def infer_output_path(path1, path2, output=None):
    """推断默认输出路径：默认落到当前工作目录。

    规则：
    - 两表后缀相同（csv/xlsx/xls/xlsm）→ 沿用同后缀
    - 两表后缀不同 → 优先 .xlsx
    """
    if output:
        return Path(output)

    p1 = Path(path1)
    p2 = Path(path2)

    s1 = p1.suffix.lower()
    s2 = p2.suffix.lower()

    if s1 == s2:
        ext = s1
    else:
        ext = '.xlsx'

    return Path.cwd() / f"dedup_merged{ext}"


def print_stats(stats, key1, key2):
    """打印合并统计信息。"""
    print(f"\n{'='*40}")
    print("合并统计")
    print(f"{'='*40}")
    print(f"  总计行数:        {stats['total']}")
    print(f"  两表都匹配:      {stats['both']}（含模糊匹配 {stats['fuzzy']}）")
    print(f"  仅文件1有:       {stats['left_only']}（关键列: {key1}）")
    print(f"  仅文件2有:       {stats['right_only']}（关键列: {key2}）")
    if stats.get('null_key_1', 0) or stats.get('null_key_2', 0):
        print(f"  其中空关键值行:  文件1={stats['null_key_1']}，文件2={stats['null_key_2']}（未参与匹配）")
    print(f"{'='*40}\n")


def write_match_log(merged, key1, output_path):
    """dry-run 时输出匹配明细到 output_path。

    真实字段：关键值 + match_type（exact/fuzzy/left_only/right_only/null_key_1/null_key_2）

    A-2 Fix 6: 写文件失败时 raise SystemExit，与 read_table / write_table 错误处理
    一致；silent failure 让 dry-run 失去 '真实预览' 意义。
    """
    if 'match_type' not in merged.columns:
        print("警告: match_type 列缺失，跳过匹配明细。", file=sys.stderr)
        return
    try:
        merged[[key1, 'match_type']].to_csv(output_path, index=False, encoding='utf-8-sig')
    except Exception as e:
        raise SystemExit(f"错误: 写入匹配明细失败: {e}")


def _script_version():
    """从 VERSION.yaml 读版本号；读取失败时回落到 'unknown'。"""
    try:
        version_file = Path(__file__).resolve().parent.parent / "VERSION.yaml"
        if version_file.is_file():
            for raw in version_file.read_text(encoding="utf-8").splitlines():
                line = raw.strip()
                if line.startswith("version:"):
                    return line.split(":", 1)[1].strip().strip('"').strip("'")
    except Exception:
        pass
    return "unknown"


def main():
    parser = argparse.ArgumentParser(
        prog="dedup_merge.py",
        description="两表去重合并：按关键列匹配，合并同名记录，保留两表全部列。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 自动检测关键列并合并
  python dedup_merge.py -1 tableA.csv -2 tableB.csv

  # 指定关键列
  python dedup_merge.py -1 tableA.csv -2 tableB.csv --key1 "期刊名称" --key2 "期刊全称"

  # 关闭模糊匹配，仅精确匹配
  python dedup_merge.py -1 tableA.csv -2 tableB.csv --no-fuzzy

  # 中文学术表，使用 cjk 预设
  python dedup_merge.py -1 tableA.csv -2 tableB.csv --fuzzy-preset cjk

  # 预览模式（不写入文件）
  python dedup_merge.py -1 tableA.csv -2 tableB.csv --dry-run
        """.strip(),
    )
    parser.add_argument('--version', action='version', version=f"%(prog)s {_script_version()}")
    parser.add_argument('-1', '--file1', required=False, default=None, help='第一个输入文件（CSV/XLSX）')
    parser.add_argument('-2', '--file2', required=False, default=None, help='第二个输入文件（CSV/XLSX）')
    parser.add_argument('--key1', help='文件1的关键列名（默认自动检测）')
    parser.add_argument('--key2', help='文件2的关键列名（默认与 --key1 相同或自动检测；可单独指定）')
    parser.add_argument(
        '--key-uniqueness-min', type=float, default=0.5, metavar='FLOAT',
        help='自动检测关键列时要求的最低唯一值比例（0-1，默认 0.5）',
    )
    parser.add_argument('-o', '--output', help='输出文件路径（默认当前工作目录下 dedup_merged.<ext>）')
    parser.add_argument('--no-fuzzy', action='store_true', help='关闭模糊匹配（默认开启）')
    parser.add_argument(
        '--fuzzy-preset', choices=['academic', 'cjk', 'none'], default='academic',
        help='模糊匹配预设（默认 academic）',
    )
    parser.add_argument(
        '--sort-by-key', action='store_true',
        help='合并后按关键列大写排序（默认关闭，保留两表原行序）',
    )
    parser.add_argument('--dry-run', action='store_true', help='预览模式：输出匹配统计与明细，不写入文件')
    # P1-2 fix: dry-run 默认**不**写匹配明细文件；用 --match-log PATH 显式指定
    parser.add_argument(
        '--match-log', metavar='PATH', default=None,
        help='dry-run 模式下的匹配明细输出路径（缺省时不写文件；与 --dry-run 配合使用）',
    )
    # A-1 P0-2/3 fix: 默认拒绝覆盖已存在输出；--force 显式允许
    parser.add_argument(
        '--force', action='store_true',
        help='允许覆盖已存在输出文件（与默认输出覆盖保护配合使用）',
    )
    parser.add_argument('-v', '--verbose', action='store_true', help='显示详细日志')
    # A-3 P1-2 fix: 自检 conda env 一致性（不需要 -1/-2；专供 evals 自检用）
    parser.add_argument(
        '--check-conda-env-consistency', action='store_true',
        help='自检：核对 SKILL.md frontmatter compatibility.runtime[*].name '
             '与 agents/openai.yaml system_requirements.conda_env 一致；返回 0 表示一致，1 表示不一致',
    )

    args = parser.parse_args()

    # A-3 P1-2 fix: 自检 conda env 一致性，先于常规流程返回
    if args.check_conda_env_consistency:
        import re as _re_consistency
        skill_md = (Path(__file__).resolve().parent.parent / 'SKILL.md').read_text(encoding='utf-8')
        openai_yaml_text = (Path(__file__).resolve().parent.parent / 'agents' / 'openai.yaml').read_text(encoding='utf-8')
        runtime_names = _re_consistency.findall(
            r"compatibility:.*?runtime:\s*([\s\S]*?)(?=\n[a-zA-Z_]|\Z)", skill_md
        )
        runtime_block = _re_consistency.search(r"runtime:\s*([\s\S]*?)(?=\n[a-zA-Z_]|\Z)", skill_md)
        runtime_match_names = []
        if runtime_block:
            runtime_match_names = _re_consistency.findall(r"-\s*name:\s*(\S+)", runtime_block.group(1))
        conda_match = _re_consistency.search(r"conda_env:\s*(\S+)", openai_yaml_text)
        conda_env_name = conda_match.group(1) if conda_match else None
        ok = bool(runtime_match_names) and conda_env_name in runtime_match_names
        if ok:
            print(f"OK: conda env consistency check passed; conda_env={conda_env_name}, runtime_names={runtime_match_names}")
            sys.exit(0)
        print(
            f"FAIL: conda env 不一致；"
            f"openai.yaml.conda_env={conda_env_name!r}，"
            f"SKILL.md.compatibility.runtime[*].name={runtime_match_names!r}",
            file=sys.stderr,
        )
        sys.exit(1)

    if not (args.file1 and args.file2):
        raise SystemExit("错误: 缺少必需参数 -1/--file1 与 -2/--file2。")

    # A-1 P0-2/3/P1-4 fix: 早期护栏——dry-run 与正式运行都受同样保护
    # 顺序：先看 --match-log 与默认主表冲突（影响 --dry-run 行为），
    # 再看 -o 与输入文件重名 / 输出已存在（影响正式写入）
    if args.match_log and not args.dry_run:
        # 冲突检查依赖 infer_output_path；这里只给警告，写表前再统一保护
        print(
            f"提示: --match-log ({args.match_log}) 必须与 --dry-run 同用，已忽略。",
            file=sys.stderr,
        )

    # A-1 P0-2 fix: 拒绝让输出覆盖输入文件
    output_path_early = infer_output_path(args.file1, args.file2, args.output)
    output_resolved_early = Path(output_path_early).resolve()
    for inp in (args.file1, args.file2):
        try:
            if output_resolved_early == Path(inp).resolve():
                raise SystemExit(
                    f"错误: 输出路径与输入文件重名 ({output_path_early})，拒绝覆盖输入。"
                    f" 请用 -o 显式指定不同的输出路径。"
                )
        except OSError:
            pass

    # A-1 P0-3 fix: 默认拒绝覆盖已存在输出；--force 显式允许
    if output_path_early.exists() and not args.force:
        raise SystemExit(
            f"错误: 输出文件已存在 ({output_path_early})，拒绝覆盖。"
            f" 用 -o 指定新路径，或传 --force 显式允许覆盖。"
        )

    # 读取输入
    if args.verbose:
        print(f"读取文件1: {args.file1}")
    try:
        df1 = read_table(args.file1)
    except FileNotFoundError as e:
        # A-1 P1-2 fix: 不在用户面前抛 traceback；verbose 模式仍附 stack trace
        msg = f"错误: 文件 1 ({args.file1}) 不存在或不可读: {e}"
        if args.verbose:
            import traceback as _tb
            msg += "\n" + _tb.format_exc()
        raise SystemExit(msg)
    except (ValueError, UnicodeDecodeError) as e:
        msg = f"错误: 文件 1 ({args.file1}) 读取失败: {e}"
        if args.verbose:
            import traceback as _tb
            msg += "\n" + _tb.format_exc()
        raise SystemExit(msg)
    if args.verbose:
        print(f"  行数: {len(df1)}, 列: {list(df1.columns)}")

    if args.verbose:
        print(f"读取文件2: {args.file2}")
    try:
        df2 = read_table(args.file2)
    except FileNotFoundError as e:
        msg = f"错误: 文件 2 ({args.file2}) 不存在或不可读: {e}"
        if args.verbose:
            import traceback as _tb
            msg += "\n" + _tb.format_exc()
        raise SystemExit(msg)
    except (ValueError, UnicodeDecodeError) as e:
        msg = f"错误: 文件 2 ({args.file2}) 读取失败: {e}"
        if args.verbose:
            import traceback as _tb
            msg += "\n" + _tb.format_exc()
        raise SystemExit(msg)
    if args.verbose:
        print(f"  行数: {len(df2)}, 列: {list(df2.columns)}")

    # 检测关键列
    key_hint = args.key1
    key2_hint = args.key2
    if key_hint and key2_hint:
        # 两个都指定了
        key1, key2 = args.key1, args.key2
        # 验证存在性
        if key1 not in df1.columns:
            lower_map = {c.lower(): c for c in df1.columns}
            if key1.lower() in lower_map:
                key1 = lower_map[key1.lower()]
            else:
                raise SystemExit(f"错误: 文件1中未找到列 '{args.key1}'，可用列: {list(df1.columns)}")
        if key2 not in df2.columns:
            lower_map = {c.lower(): c for c in df2.columns}
            if key2.lower() in lower_map:
                key2 = lower_map[key2.lower()]
            else:
                raise SystemExit(f"错误: 文件2中未找到列 '{args.key2}'，可用列: {list(df2.columns)}")
    else:
        # A-2 Fix 3 (DET-1): detect_key_column 退化路径（列名相似度 < 0.5 等）抛
        # ValueError 时 main 未捕获会暴露 traceback；统一在关键调用点包裹
        # try/except → SystemExit 单行错误，verbose 模式附 stack trace
        try:
            key1, key2 = detect_key_column(
                df1, df2,
                key_hint=key_hint, key2_hint=key2_hint if (key2_hint and not key_hint) else None,
                uniqueness_min=args.key_uniqueness_min,
            )
        except ValueError as e:
            msg = f"错误: 自动检测关键列失败: {e}"
            if args.verbose:
                import traceback as _tb
                msg += "\n" + _tb.format_exc()
            raise SystemExit(msg)

    if args.verbose:
        print(f"关键列: 文件1='{key1}', 文件2='{key2}'")

    # 合并
    fuzzy = not args.no_fuzzy
    merged, stats = merge_tables(
        df1, df2, key1, key2,
        fuzzy=fuzzy, fuzzy_preset=args.fuzzy_preset,
        sort_by_key=args.sort_by_key,
    )

    # 输出统计
    print_stats(stats, key1, key2)

    if args.dry_run:
        print("[dry-run] 不写入合并结果文件。")
        print(f"合并后预览（前10行）:")
        print(merged.head(10).to_string(index=False))

        # A-1 P1-3 fix: --match-log 在非 dry-run 下被静默忽略的提示
        # A-1 P1-4 fix: --match-log 路径与默认主表重名时报错退出
        # A-2 Fix 2 (PATH-1 / MATCHLOG-2): 把 P1-4 检查改用 args.output 解析后的
        # 实际路径（与 write_table 用的同一变量 `output_path_early`），而非
        # infer_output_path 默认值；同时覆盖 dry-run 模式（用户传 -o + --match-log
        # 同名时也拒绝）
        if args.match_log:
            match_log = Path(args.match_log)
            try:
                if Path(match_log).resolve() == output_path_early.resolve():
                    raise SystemExit(
                        f"错误: --match-log ({match_log}) 与主表路径 ({output_path_early}) 重名；"
                        f" 请改为不同的路径。"
                    )
            except OSError:
                pass
            write_match_log(merged, key1, match_log)
            print(f"[dry-run] 匹配明细已写入: {match_log}")
        else:
            print("[dry-run] 未指定 --match-log，跳过匹配明细文件输出。")
        return

    # 推断输出路径（早期护栏已处理 P0-2/P0-3，此处仅复用路径变量）
    output_path = output_path_early

    print(f"写入: {output_path}")
    # A-2 Fix 3 (DIR-1 / PERM-1): write_table 写表前未做目录/权限检查；
    # -o 路径所在目录不存在时抛 OSError、-o 路径无写权限时抛 PermissionError，
    # main 未捕获会暴露 traceback；统一在关键调用点包裹
    # try/except → SystemExit 单行错误，verbose 模式附 stack trace
    try:
        write_table(merged, output_path)
    except (OSError, PermissionError) as e:
        msg = f"错误: 写入主表失败 ({output_path}): {e}"
        if args.verbose:
            import traceback as _tb
            msg += "\n" + _tb.format_exc()
        raise SystemExit(msg)
    print("完成。")


if __name__ == '__main__':
    main()
