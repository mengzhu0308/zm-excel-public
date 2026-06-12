---
name: zm-excels-merge
description: >-
  当用户需要合并、组合或拼接多个 Excel（.xlsx / .xls / .xlsm）或 CSV 文件为单个文件
  时，使用此 skill。

  触发场景：合并目录下所有 Excel 文件；将多个表格合并为一个；将多个 CSV 的行
  追加到单个文件；整合多个工作簿的数据；批量合并销售报表/月度数据/导出表格；
  或任何涉及多个文件垂直拼接表格数据的任务。也包括用户提到 '合并Excel'、
  '合并多个表格'、'把多个Excel合并成一个'、'批量合并csv'、'汇总多个文件的数据'
  等。

  **不要**触发：
  - 单个文件内的数据查询、筛选、排序、去重：使用 `zm-excel-query`、`zm-excel-sort`
    或 `zm-excel-dedup-merge`（按需）
  - 创建/写入公式、追加单条记录、按条件删除若干行：使用对应的写操作 skill
    （`zm-excel-add-one-row`、`zm-excel-del-multi-rows`）
  - 原地修改单元格格式（字体、对齐、列宽等）：使用 `zm-excel-formalization`
  - 单文件格式转换（xlsx↔csv）：使用 `zm-csv2xlsx` 或 `zm-xlsx2csv`
license: MIT
metadata:
  skill_mode: hybrid
compatibility:
  runtime:
    - name: agent-skills
      call_command: "conda run -n agent-skills python \"$SKILL_DIR/scripts/merge_excels.py\" [args]"
---

# zm-excels-merge

合并多个 Excel（.xlsx / .xls / .xlsm）和 CSV 文件，支持字段兼容性分析、智能分组和合并清单工作流。

## 触发条件

- 合并目录下所有 Excel/CSV 文件为一个
- 垂直拼接多个结构相似的数据文件
- 按同名 sheet 合并多个 xlsx 文件
- 将单个文件内的多个 sheet 合并为一个 sheet
- 分析字段兼容性并给出合并建议

## 核心原则

- **智能字段分组**：默认模式下走 `analyze_compatibility` 路径——先按后缀分组、再按 sheet 收集列、并集相似度判断、贪心单链扩展；列名集合完全相同的文件合入同组，否则分到不同 sheet
- **单文件多 sheet 合并**：`--merge-sheets` 模式下，将文件内所有 sheet 按字段兼容性分组后垂直合并
- **自动列对齐**：列名不一致时缺失列自动补空值，不丢弃数据
- **字段兼容性分析**：`--preview` 分析列名重叠度并生成分组清单
- **按后缀分组合并**：不同后缀（`.xlsx`、`.csv` 等）各自独立成组
- **多 sheet 默认首 sheet**：目录/多文件输入时，Excel 文件默认只读第一个 sheet（`--merge-sheets` 模式除外）
- **来源可追溯**：可选添加来源列标注原始文件
- **输出目录自动推断**：未指定 `-o` 时自动推断输出路径
- **空文件安全跳过**：空文件或无法读取的文件自动跳过并记录警告
- **灵活输入**：支持目录扫描或直接指定多个文件路径
- **单文件保护**：输入路径指向单个文件时自动识别为单文件模式

## 长任务 / 多轮合并

本 skill 支持”一次性合并”和”多轮合并”两种使用模式。当合并涉及多个文件、字段兼容性判断、清单调整或中断后续跑时，建议按”发现文件 → 预览字段兼容性 → 调整合并清单 → 执行合并 → 抽样验证输出”五个阶段推进。

- **目标记录**：写清输入范围、目标输出文件、是否递归、是否添加来源列、字段相似度阈值和验收标准。
- **阶段检查点**：发现文件 → 预览字段兼容性 → 调整合并清单 → 执行合并 → 抽样验证输出。
- **恢复点**：若已生成 `合并清单.md`，续跑时优先读取该清单和当前会话上下文，不重新猜测分组意图。
- **完成条件**：输出文件存在，sheet/行数/来源列符合目标，警告与跳过文件已向用户说明。

> 本节不绑定任何会话层目标跟踪工具；不同 AI 客户端的会话层能力各不相同，使用本 skill 时不依赖特定工具存在。

## 支持的文件类型

| 扩展名  | 说明                |
| ------- | ------------------- |
| `.xlsx` | Excel 2007+ 格式    |
| `.xls`  | Excel 97-2003 格式  |
| `.xlsm` | Excel 宏启用格式    |
| `.csv`  | CSV 文本格式        |

## 合并行为

### 多文件输入默认策略

1. **按后缀分组**：`.xlsx`、`.xls`、`.xlsm`、`.csv` 各自独立成组。同组内按 sheet 名和字段兼容性合并，不同后缀输出到不同 sheet（如 `Sheet1_xlsx_组1`）。
2. **默认只取首 sheet**：Excel 文件默认只读第一个 sheet；如需全部 sheet，用 `--merge-sheets`。

### 三种工作模式

#### 模式一：智能合并（默认）

不提供 `--preview` 和 `--plan` 时自动分组合并：
1. 按后缀分组 → 读取文件（Excel 默认首 sheet）→ 计算 Jaccard 相似度
2. 经 `analyze_compatibility` 字段兼容性分析：若某 sheet 下所有文件列名完全一致（`total_groups == total_sheets`），按同名 sheet 直接合并；否则按 `--similarity-threshold`（默认 0.8）分组输出到独立 sheet
3. 每组缺失列自动补空值

强制按同名 sheet 合并（不自动分组）：`--no-auto-group`

#### 模式二：预览分析（--preview）

分析字段兼容性并输出 `合并清单.md` 分组建议。

#### 模式三：按清单合并（--plan）

读取用户调整后的合并清单执行合并。

#### 模式四：单文件多 sheet 合并（--merge-sheets）

将每个文件内的所有 sheet 按字段兼容性分组后垂直合并：
- 字段兼容的 sheet 合并为一个输出 sheet
- 不兼容的分到不同组（如 `Merged_组1`）
- 配合 `--add-source` 可追踪数据来源文件和 sheet

与默认模式的区别：默认模式按 sheet 名跨文件合并；`--merge-sheets` 按文件内合并。

### 字段兼容性分析

对每个 sheet：
1. 收集每个文件的列名集合
2. 计算文件两两之间的 Jaccard 相似度：`|A∩B| / |A∪B|`
3. 使用贪心单链扩展分组：取一个未分组文件作为 seed，依次把所有与组内任一文件相似度 >= 阈值的文件并入该组，直到没有更多文件可并入
4. 对剩余未分组文件重复上述过程

> 详细算法（贪心单链扩展、阈值选择、边界情况）见 [references/grouping-algorithm.md](references/grouping-algorithm.md)。

### CSV 文件处理

CSV 文件视为只有一个 "sheet"（默认名为 `Sheet1`）。由于多文件输入时按后缀分组，CSV 文件只会与同后缀（即其他 `.csv`）文件的 `Sheet1` 合并，不会与 `.xlsx` 文件的 `Sheet1` 混在一起。

### 列对齐规则

1. 收集所有文件中该 sheet 的所有列名（并集）
2. 每个文件缺失的列补空值
3. 最终列顺序以第一个遇到该 sheet 的文件的列顺序为基准

> 列顺序属于**隐式约定**：当一个分组内多个文件的列顺序不一致时，输出列顺序以该组内第一个被读取到的文件为准；新出现的列追加到末尾。脚本不提供 `--column-order` 参数来覆盖该约定——如需自定义列顺序，请先用 `--preview` 生成清单并人工调整列顺序（或在合并后用 `zm-excel-formalization` / `zm-excel-sort` 等 skill 二次处理）。

### 输出格式

- 输出 `.xlsx`：多 sheet 分别写入不同工作表
- 输出 `.csv`：所有 sheet 的数据按顺序垂直拼接为一个 CSV（sheet 之间无分隔）

## 使用方式

### 方式一：直接调用脚本

**智能合并（默认模式）：**

```bash
# 合并目录
conda run -n agent-skills python "$SKILL_DIR/scripts/merge_excels.py" -d ./data

# 指定多个文件（逗号/空格/中文逗号/顿号分隔）
conda run -n agent-skills python "$SKILL_DIR/scripts/merge_excels.py" \
  -f "report_01.xlsx, report_02.xlsx, summary.csv"

# 指定输出、递归、来源列、CSV 输出、表头行
conda run -n agent-skills python "$SKILL_DIR/scripts/merge_excels.py" \
  -d ./data -o ./merged.xlsx -r --add-source --header 1

# 单文件多 sheet 合并
conda run -n agent-skills python "$SKILL_DIR/scripts/merge_excels.py" \
  -f "report.xlsx" --merge-sheets

# 强制按同名 sheet 合并（不自动分组）
conda run -n agent-skills python "$SKILL_DIR/scripts/merge_excels.py" \
  -d ./data --no-auto-group
```

**预览分析模式：**

```bash
conda run -n agent-skills python "$SKILL_DIR/scripts/merge_excels.py" \
  -d ./data --preview --similarity-threshold 0.6 --plan-output ./my_plan.md
```

**按清单合并模式：**

```bash
conda run -n agent-skills python "$SKILL_DIR/scripts/merge_excels.py" \
  --plan ./合并清单.md -o ./merged.xlsx
```

### 合并清单格式

`合并清单.md` 由 `--preview` 自动生成，用户可手动编辑后通过 `--plan` 执行：

```markdown
# Excel 合并清单

## 分组 1: Sheet1
- **文件列表**: ./data/A.xlsx, ./data/B.xlsx
- **Sheet 名**: Sheet1
- **共同列** (3 个): 姓名, 年龄, 城市
- **全部列** (4 个): 姓名, 年龄, 城市, 备注
- **组内相似度**: 0.75
- **输出 sheet 名**: 合并_Sheet1_组1
```

可编辑项：文件列表、输出 sheet 名；删除分组段落可跳过该组合并。

### 方式二：AI 推理执行

```python
import pandas as pd
from pathlib import Path

input_path = Path('./data/report.xlsx')
output_path = Path('./output/merged.xlsx')  # 未指定时自动推断

# 自动推断输出路径
if output_path is None:
    if input_path.is_file():
        output_path = input_path.parent / f"{input_path.stem}_merged.xlsx"
    else:
        output_path = input_path.parent / f"{input_path.name}-excel-merging" / "merged.xlsx"
        output_path.parent.mkdir(parents=True, exist_ok=True)

if input_path.is_file():
    # 单文件：读取所有 sheet（pd.read_excel 不指定 sheet_name 时返回 dict；示例用 concat 合并）
    all_sheets = pd.read_excel(input_path, sheet_name=None, dtype=str, keep_default_na=False)
    df = pd.concat(all_sheets.values(), ignore_index=True)
    df['来源文件'] = input_path.name
    df.to_excel(output_path, index=False)
else:
    # 目录：按后缀分组，Excel 默认只取首 sheet
    files = (
        list(input_path.glob('*.xlsx'))
        + list(input_path.glob('*.xls'))
        + list(input_path.glob('*.xlsm'))
        + list(input_path.glob('*.csv'))
    )
    files_by_suffix = {}
    for f in files:
        files_by_suffix.setdefault(f.suffix.lower(), []).append(f)
    merged_sheets = {}
    for suffix, suffix_files in files_by_suffix.items():
        dfs = []
        for f in suffix_files:
            reader = (
                pd.read_csv
                if f.suffix == '.csv'
                else lambda p: pd.read_excel(p, sheet_name=0, dtype=str, keep_default_na=False)
            )
            df = reader(f)
            # f.name 是同后缀组内的来源文件名；若用户开启 --add-source，才作为"来源列"追加
            df['来源文件'] = f.name
            dfs.append(df)
        if dfs:
            merged = pd.concat(dfs, ignore_index=True)
            merged_sheets[f"Sheet1_{suffix.lstrip('.')}"] = merged
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df in merged_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
```

## 脚本参数说明

| 参数 | 简写 | 说明 | 默认值 |
|------|------|------|--------|
| `--directory` | `-d` | 输入目录路径 | 无 |
| `--files` | `-f` | 直接指定多个文件路径，逗号/空格/中文逗号/顿号分隔 | 无 |
| `--output` | `-o` | 输出文件路径（目录输入时默认自动推断） | 自动推断 |
| `--recursive` | `-r` | 递归搜索子目录 | `False` |
| `--pattern` | `-p` | 文件匹配模式，逗号分隔 | `*.xlsx,*.xls,*.xlsm,*.csv` |
| `--sheets` | `-s` | 只合并指定 sheet 名，逗号分隔 | 全部 |
| `--header` | | 表头行号（0-based） | `0` |
| `--add-source` | | 添加来源列 | `False` |
| `--source-col` | | 来源列名 | `来源文件` |
| `--preview` | | 预览模式：分析兼容性并生成分组清单 | `False` |
| `--plan` | | 执行模式：按合并清单执行合并 | 无 |
| `--similarity-threshold` | | 字段相似度阈值（Jaccard） | `0.8` |
| `--plan-output` | | 预览模式下清单输出路径 | `合并清单.md` |
| `--no-auto-group` | | 禁用默认模式下的自动字段分组，强制按同名 sheet 合并 | `False` |
| `--merge-sheets` | | 单文件多 sheet 合并：将每个文件内的所有 sheet 按字段兼容性分组后合并 | `False` |
| `--force` | | 强制覆盖已存在的输出文件 | `False` |
| `--log-level` | | 日志级别（`DEBUG` / `INFO` / `WARNING` / `ERROR`） | `INFO` |

## 代码规范

- 使用 `pandas` 读取 Excel/CSV，用 `openpyxl` 作为默认引擎
- **数据类型保持**：读取时统一使用 `dtype=str`，保持原始文本形式，避免大数字等文本形式数据被 pandas 自动推断为数值类型导致精度丢失
- **空值处理**：读取时设置 `keep_default_na=False`，防止 pandas 将 "NA"、"N/A" 等字符串误识别为空值；缺失列补 `pd.NA`，写入时显示为空单元格
- 输出 xlsx 使用 `openpyxl` 引擎
- 输出 CSV 使用 `utf-8-sig` 编码，确保 Excel 打开中文不乱码
- 不输出 DataFrame 索引（`index=False`）
- 错误处理显式捕获并打印警告信息，不中断整个合并流程
- 空文件或无法读取的文件自动跳过并记录警告
- 不存在的路径自动跳过并记录警告

## 注意事项

- **源文件只读**：合并过程不修改任何源文件
- **输出自动推断**：未指定 `-o` 时，单文件生成 `文件名_merged.xlsx`，目录生成 `原目录名-excel-merging/merged.xlsx`
- **默认阈值 0.8**：较为保守，避免语义不同的字段被强行合并；如需强制按同名 sheet 合并用 `--no-auto-group`
- **大数据量**：文件数量极多或单个文件极大时，建议先检查内存
- **跨后缀合并**：`.xlsx`、`.csv` 等各自独立成组，需跨后缀合并时请分别处理后再手动整合
- **重复表头**：默认保留每个文件的表头行；如需去重，合并后手动处理
- **输入优先级**：`-d` 和 `-f` 可同时使用，文件会去重合并

## 环境与依赖

- 脚本依赖：`pandas`、`openpyxl`，以及 `xlrd>=2.0`（仅在需要读取 `.xls` 时使用；若仅处理 `.xlsx` / `.csv` 可不装）
- 安装：`pip install 'xlrd>=2.0' pandas openpyxl` 或使用 `conda` 等价命令
- 运行入口以 SKILL.md frontmatter `compatibility.runtime.call_command` 为准；本仓库默认推荐 `conda run -n agent-skills python ...`。脚本本身不主动选择解释器；若本机无 conda 或无 `agent-skills` 环境，AI 推理或评测 runner（`evals/run_smoke.py`）可降级到当前 Python 解释器并打印 `WARNING`
- 脚本默认不写 `__pycache__/`（入口处设置 `sys.dont_write_bytecode = True`）；如果想强制不写字节码缓存，可显式用 `python3 -B ...`（`python3 -B` 的语义是“关闭字节码生成”，与脚本默认行为一致）
