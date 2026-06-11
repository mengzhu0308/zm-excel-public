---
name: zm-excel-query
description: >-
  从 Excel 中查询、筛选、搜索、预览或提取数据，并保存为同目录同名 CSV 或 XLSX。
  适用于：按条件过滤行、选择列、排序、去重、分组聚合、预览 Excel 结构、
  把 Excel 转换为 CSV、提取数据子集，或任何涉及 .xlsx/.xls/.csv/.xlsm
  的结构化数据查询任务。源 Excel 只读，不会被修改。

  触发表述举例：'查询 Excel'、'筛选表格'、'过滤数据'、'从 Excel 导出 CSV'、
  '提取满足条件的数据'、'查看 Excel 数据结构'、'先预览再查询'。

  **不要**触发：创建复杂 Excel 公式、财务建模、生成图表、原位修改单元格格式——
  使用对应的 xlsx skill。
license: MIT
metadata:
  skill_mode: hybrid
compatibility:
  runtime:
    - name: agent-skills
      call_command: "conda run -n agent-skills python \"$SKILL_DIR/scripts/query_excel.py\" [args]"
---

# zm-excel-query

从 Excel 查询数据，结果保存为同目录同名 CSV 或 XLSX。支持预览数据结构后再查询。源 Excel 只读，不会被修改。

## 触发条件

- 筛选/过滤满足条件的数据
- 选择特定列、排序、去重
- 聚合统计（分组求和、计数、平均值等）
- 将 Excel 查询结果导出为 CSV
- 提取数据子集用于后续分析
- **预览 Excel 数据结构**（列名、类型、前 N 行）
- **先查看表格内容再决定如何查询**
- **交互式逐步构建查询条件**

## 核心原则

- **先预览后查询**：不熟悉的数据先预览结构（列名、类型、示例值），确认字段后再查询
- **查询即转换**：以 pandas DataFrame 为中间载体，最终输出 CSV
- **多操作组合，一步完成**：`--where`、`--select`、`--exclude`、`--sort`、`--distinct`、`--groupby`、`--agg`、`--limit`、`--offset` 可在单次查询中任意组合，AI 推理执行时应优先单条命令完成
- **"分组"不等于"聚合"**：用户提到"按XX分组"但未要求聚合统计时，理解为**按该列排序**，使用 `--sort`；`--groupby` 仅用于真正聚合，必须与 `--agg` 配合。不要擅自添加 `--select`
- **零公式依赖**：不涉及 Excel 公式写入或单元格格式化
- **运行时只读保护**：脚本仅调用 pandas 只读接口（`read_excel` / `read_csv`），并对 `input_path` 做 `os.stat` 快照对比（mtime / size），若查询过程中源文件被改则返回 `{"_read_only_violation": True}` 强制 exit 2；源 Excel 始终保持原始状态
- **输出即输入的映射**：输出与源 Excel 同目录、同名（仅扩展名不同）；可通过 `--tag` 附加标识（**仅允许字母数字下划线连字符**，`../` 等路径字符会被拒绝）；默认命名时若目标文件已存在则自动追加序号避免覆盖

## 与 /goal 配合使用

`/goal` 是会话层的任务跟踪能力（Codex CLI、Claude Code 等平台均提供），不是本 skill 的脚本参数。明确的一次性查询可直接执行；当用户需要先探索字段、多轮调整条件、批量导出多个结果或中断后续跑时，建议先开启 `/goal`。

- **目标记录**：写清源文件、目标 sheet、查询意图、输出格式、标签命名和结果验收标准。
- **阶段检查点**：按“预览结构 → 确认字段和条件 → 执行查询 → 抽样校验结果 → 保存输出”推进。
- **恢复点**：续跑时优先使用已确认的字段、条件、`--tag` 和输出路径，不从自然语言重新推断。
- **完成条件**：输出文件存在，列集合、筛选行样本、排序/聚合结果符合目标。

## 数据预览

### 预览输出内容

| 信息项 | 说明 |
|--------|------|
| 总行数 | 数据表的总行数 |
| 总列数 | 数据表的总列数 |
| 字段信息 | 每列的名称、数据类型（整数/浮点数/文本/日期时间/布尔）、非空值数量、前 3 个示例值 |
| 前 N 行 | 实际数据样本 |

### 预览用法

```bash
# 预览前 10 行（默认）
conda run -n agent-skills python "$SKILL_DIR/scripts/query_excel.py" \
  -f data.xlsx --preview

# 预览前 20 行
conda run -n agent-skills python "$SKILL_DIR/scripts/query_excel.py" \
  -f data.xlsx --preview 20
```

## 支持的查询操作

| 操作 | 说明 | 示例 |
|------|------|------|
| `--where` | 条件筛选，支持比较运算符 | `age > 30`, `name == '张三'` |
| `--select` | 选择列，逗号分隔；与 `--exclude` 同时使用时后执行 | `name,age,salary` |
| `--sort` | 按列排序，逗号分隔；在 `--groupby` 之后执行，支持对聚合结果排序 | `salary desc,age asc` |
| `--distinct` | 按指定列去重 | `department` |
| `--groupby` | 分组聚合；聚合完成后若同时指定 `--sort`，排序作用于聚合结果 | `department` + `--agg sum:salary` |
| `--exclude` | 排除列，逗号分隔；与 `--select` 同时使用时先执行 | `序号,id` |
| `--limit` | 限制返回行数 | `100` |
| `--offset` | 跳过前 N 行 | `10` |

## 输出规范

**默认规则**：输出文件与源 Excel 同目录、同名，扩展名由 `--format` 决定（默认 `.csv`，可选 `.xlsx`）。

```
输入: /path/to/sales.xlsx
输出: /path/to/sales.csv        # --format csv（默认）
输出: /path/to/sales.xlsx       # --format xlsx
```

多表输出（如按 sheet 分片）时，在原文件名后追加 sheet 名：`data_Sheet1.csv`、`data_Sheet2.csv`。

**未指定 `--sheet` 时的多 sheet 行为**：当 `--sheet` 缺省且文件含多个工作表时，`pd.read_excel` 返回 dict，脚本会对每个 sheet 各跑一次查询并按 sheet 名分片输出；与显式传 `--sheet` 时只跑单 sheet 的行为不同。

通过 `--tag` 附加标识区分不同查询：`--tag EastHigh` → `sales_EastHigh.csv`。

**自动防覆盖**：默认命名（未指定 `-o`）时，若目标文件已存在，自动追加序号：`sales.csv` → `sales_1.csv` → `sales_2.csv`。自定义 `-o` 路径时不触发。

## 使用方式

### 方式一：直接调用脚本（确定性操作）

```bash
# 基础查询
conda run -n agent-skills python "$SKILL_DIR/scripts/query_excel.py" \
  -f data.xlsx \
  --where "age > 30 and department == 'Engineering'" \
  --select "name,age,salary" \
  --sort "salary desc"

# 一步完成多操作（筛选 + 列选择 + 分组聚合 + 排序）
conda run -n agent-skills python "$SKILL_DIR/scripts/query_excel.py" \
  -f data.xlsx \
  --where "department in ['Sales', 'Marketing']" \
  --select "department,name,salary" \
  --groupby "department" \
  --agg "mean:salary,count:name" \
  --sort "salary_mean desc"
```

### 方式二：数据预览

```bash
conda run -n agent-skills python "$SKILL_DIR/scripts/query_excel.py" \
  -f data.xlsx --preview
```

### 方式三：交互式查询

```bash
conda run -n agent-skills python "$SKILL_DIR/scripts/query_excel.py" \
  -f data.xlsx --interactive
```

交互模式命令：

| 命令 | 作用 |
|------|------|
| `where <条件>` | 设置筛选条件 |
| `select <列>` | 选择列（逗号分隔） |
| `sort <规则>` | 设置排序规则 |
| `limit <N>` | 限制返回行数 |
| `run` / 回车 | 执行当前查询并展示结果前 10 行 |
| `save [路径]` | 保存结果为 CSV |
| `reset` | 重置所有查询条件 |
| `preview` | 重新显示原始数据预览 |
| `help` | 显示命令帮助 |
| `quit` / `exit` | 退出交互模式 |

### 方式四：AI 推理执行

```python
import pandas as pd

df = pd.read_excel('data.xlsx')
print(df.head(10))
print(df.dtypes)
result = df[df['age'] > 30][['name', 'age', 'salary']].sort_values('salary', ascending=False)
result.to_csv('data.csv', index=False, encoding='utf-8-sig')
```

## 条件表达式语法

`--where` 参数支持的运算符：

| 运算符 | 含义 | 示例 |
|--------|------|------|
| `==` | 等于 | `status == 'active'` |
| `!=` | 不等于 | `status != 'deleted'` |
| `>` `<` `>=` `<=` | 比较（数值 / 日期） | `age > 18` / `date == '2024-01-01'` |
| `and` `or` | 逻辑组合 | `age > 18 and status == 'active'` |
| `in` | 包含于列表 | `department in ['Sales', 'Marketing']` |
| `contains` | 字符串包含 | `name contains '张'` |
| `startswith` `endswith` | 字符串前缀/后缀 | `email startswith 'admin'` |

> 日期列筛选：用 ISO 字符串与 pandas datetime 列比较会自动转换。例如 `date == '2024-01-01'` 会匹配 datetime 值等于 `2024-01-01` 的行。
>
> 更细的运算符优先级、边界陷阱（`in` 列表引号、字符串内含 `>=`、3+ 条件 `and`/`or` 链路）与 `pandas.DataFrame.query()` 的差异，参见 [where 表达式语法](references/where-expression.md)。

## 代码规范

- 使用 `pandas` 读取 Excel，`openpyxl` 为默认引擎
- 输出 CSV 使用 `utf-8-sig` 编码
- 日期列保持 ISO 格式字符串（`YYYY-MM-DD`）
- 不输出 DataFrame 索引（`index=False`）
- 一次性查询和交互模式共用同一查询管线 `_apply_pipeline(df, params)`，保证两边行为一致
- 统一管线返回 `(df, error)` 元组；外层 `main()` 与交互模式据此决定是否 `sys.exit(1)` 或打印错误
- 解析器（`parse_where`）不向入参 `df` 写入临时列；占位符走 `masks` 字典，杜绝副作用
- `_parse_where` 切 `and` / `or` 改为 `re.IGNORECASE`，与 pandas.query 习惯一致
- `_PRECEDENCE` 显式优先级表决定 `and` 绑定比 `or` 紧；改顺序前请同步 `references/where-expression.md`
- `_validate_tag` 与 `_sanitize_sheet_name` 把 `tag` / `sheet_name` 中的路径危险字符拦截或替换
- 错误处理分两类：用户错误（`ValueError` / `KeyError` / `FileNotFoundError` / `LookupError`）返回 `"参数或输入错误: ..."`；其他 `Exception` 返回 `"内部错误: ..."` 并附 traceback 供调试
- `--json` 与 `--interactive` 互斥：互斥时由 `main()` 入口 `parser.error()` 拦截
- 预览和交互模式输出应适合终端展示

## 注意事项

- **源 Excel 文件只读**：查询不修改源文件，所有操作在内存中进行，结果输出到新文件；入口 / 出口 `os.stat` 快照对比确保源文件 mtime / size 不被改写
- 空值（NaN）在 CSV 中输出为空字符串
- 大数据量时，优先使用 `--select` 减少内存占用
- 若用户未指定输出路径，自动推导为同目录同名 CSV
- 交互模式适合探索性查询；确定性批量任务应使用参数化命令
- 若 Excel 表头不在第1行，可通过 `--header N` 指定（0-based，如表头在第2行则传 `--header 1`）
- 普通查询保持原始列名和列顺序；分组聚合会改变数据结构，列名遵循 pandas 原生行为
- 预览模式只读取和展示数据，不生成 CSV 输出文件
- `--tag` 仅允许字母、数字、下划线、连字符；含 `..` `/` `\\` `:` `*` `?` 等路径或 shell 危险字符会被 `ValueError` 拦截
- 多 sheet 文件中，sheet 名含路径分隔符或 `..` 会被自动替换为 `_`，避免输出路径越界
