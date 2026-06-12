# 实现说明

本文件汇总 `scripts/query_excel.py` 的内部实现约束，仅供维护者参考；不在用户使用层面提及。

## 输入读取

- 使用 `pandas` 读取 Excel，`openpyxl` 为默认引擎
- `.csv` / `.xlsx` / `.xlsm` / `.xls` 走对应解析路径
- 未知后缀显式抛 `ValueError`（A-1 P0-3 修复；之前静默 fallback 会掩盖真实错误）

## 输出编码

- 输出 CSV 使用 `utf-8-sig` 编码（含 BOM，Excel 直接打开中文不乱码）
- 日期列保持 ISO 格式字符串（`YYYY-MM-DD`）
- 不输出 DataFrame 索引（`index=False`）
- NaN 在 CSV 中输出为空字符串

## 查询管线

- 一次性查询（`_query_single`）与交互模式（`_apply_query`）共用同一管线 `_apply_pipeline(df, params)`，保证两边行为一致
- 统一管线返回 `(df, error)` 元组；外层 `main()` 与交互模式据此决定是否 `sys.exit(1)` 或打印错误
- 解析器（`parse_where`）不向入参 `df` 写入临时列；占位符走 `masks` 字典，杜绝副作用
- `_parse_where` 切 `and` / `or` 改为 `re.IGNORECASE`，与 pandas.query 习惯一致
- `_PRECEDENCE` 显式优先级表决定 `and` 绑定比 `or` 紧；改顺序前请同步 `where-expression.md`

## 路径安全

- `_validate_tag` 与 `_sanitize_sheet_name` 把 `tag` / `sheet_name` 中的路径危险字符拦截或替换
- 交互模式 `save <路径>` 拒绝 `..` 路径穿越（A-1 P0-1 修复）
- 入口 / 出口 `os.stat` 快照对比保证源文件不被改

## 错误处理

- 入口与 `_apply_pipeline` 一致：用户错误（`ValueError` / `KeyError` / `FileNotFoundError`）走"参数或输入错误"；其余走"内部错误"并附 traceback
- 未知后缀文件、非数字 `--limit` / `--offset`、缺失列等均归"参数或输入错误"
- 内部错误（pandas 配置、TypeError、AttributeError）保留 traceback 供调试
- A-1 P1-1 移除 `LookupError`（在 pandas 实际抛出场景里几乎不存在）

## 互斥与边界

- `--json` 与 `--interactive` 互斥：互斥时由 `main()` 入口 `parser.error()` 拦截
- 预览和交互模式输出应适合终端展示
- `parse_agg` 重复 col 抛 `ValueError`（A-1 P1-2 修复；之前字典赋值静默覆盖）
- `_load_skill_version` 找不到 `VERSION.yaml` 时抛 `RuntimeError`，主流程不中断（A-1 P1-3 修复；之前硬编码 "0.3.0" 与 VERSION.yaml 漂移）
