---
name: zm-csv2xlsx
description: >-
  将 CSV 文件转换为 Excel（.xlsx）格式。当用户需要转换、导入 CSV 数据到 Excel
  时，**务必优先使用此 skill**，即使只是随口一提。

  触发场景：单个 .csv 转 .xlsx；批量转换目录下所有 CSV 文件；转换文件列表
  （空格、逗号、顿号分隔）；将多个 CSV 合并为一个 Excel，每个 CSV 作为独立
  sheet；导入 CSV 数据到 Excel；或用户提到 'CSV转Excel'、'csv转xlsx'、
  '把CSV转成Excel'、'CSV导入Excel'、'批量转换CSV'、'多个CSV合并成一个Excel'
  等。

  **不要**触发：涉及 Excel 内数据查询/筛选、格式化样式、创建图表、写入公式、
  或将 Excel 转 CSV 时——使用对应的 query、formatting、conversion 或 merge
  skill。
license: MIT
metadata:
  skill_mode: hybrid
compatibility:
  runtime:
    - name: agent-skills
      call_command: "conda run -n agent-skills python \"$SKILL_DIR/scripts/csv2xlsx.py\" [args]"
---

# zm-csv2xlsx

将 CSV 文件转换为 Excel 格式（.xlsx）。支持单文件、目录批量、文件列表三种输入模式，以及多 CSV 合并到一个 Excel 不同 sheet 中。

## 触发条件

当用户需要以下操作时触发本 skill：

- 将单个 CSV 文件转换为 Excel
- 将目录下的所有 CSV 文件批量转换为 Excel
- 将多个 CSV 文件（以空格、逗号、顿号分隔的列表）批量转换为 Excel
- 将多个 CSV 合并为一个 Excel，每个 CSV 作为一个 sheet
- 将 CSV 数据导入到 Excel 中
- 处理 CSV 编码问题（自动检测 utf-8-sig / utf-8 / gb18030）

## 核心原则

- **零修改源文件**：转换过程只读取 CSV，不修改源文件
- **编码自动检测**：自动尝试 utf-8-sig、utf-8、gb18030 读取 CSV
- **灵活输入解析**：自动识别单文件、目录、文件列表；列表支持空格、逗号、顿号、分号、换行混用分隔
- **两种输出模式**：
  - **独立模式**（默认）：每个 CSV 输出为独立的 xlsx 文件
  - **合并模式**（`--combine`）：多个 CSV 合并为一个 xlsx，每个 CSV 作为一个 sheet
- **sheet 名安全处理**：文件名中的非法字符自动替换为下划线，长度超过 31 字符自动截断（openpyxl 限制）
- **不输出索引**：Excel 不包含 DataFrame 行索引

## 与 /goal 配合使用

`/goal` 是会话层的任务跟踪能力（Codex CLI、Claude Code 等平台均提供），不是本 skill 的脚本参数。单个 CSV 转换通常直接执行；当任务包含批量转换、合并多个 CSV、需要确认编码或跨会话继续时，建议先开启 `/goal`，记录输入范围、输出模式、目标文件和完成标准。

## 使用方式

### 方式一：脚本调用

```bash
# 单文件
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv

# 目录批量
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" ./data/

# 文件列表（空格、逗号、顿号、分号、换行混用分隔）
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" "a.csv, b.csv、c.csv"

# 指定输出路径
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv -o ./output/result.xlsx

# 多 CSV 合并为一个 Excel（每个 CSV 一个 sheet）
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" \
  "data1.csv data2.csv data3.csv" -o combined.xlsx --combine

# 指定 sheet 名
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv -n "销售数据"

# 无表头 CSV
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv --no-header

# 指定编码读取
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv -e gb18030
```

### 方式二：AI 推理执行

```python
import pandas as pd
from pathlib import Path

# 自动检测编码
def read_csv_auto(path):
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return pd.read_csv(path, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(f"无法检测编码: {path}")

# 单文件
df = read_csv_auto("data.csv")
df.to_excel("data.xlsx", sheet_name="Sheet1", index=False, engine="openpyxl")

# 批量转换目录
for f in Path("./data/").glob("*.csv"):
    read_csv_auto(f).to_excel(
        f.with_suffix(".xlsx"), sheet_name=f.stem[:31], index=False, engine="openpyxl"
    )

# 多 CSV 合并到一个 Excel
with pd.ExcelWriter("combined.xlsx", engine="openpyxl") as writer:
    for f in Path("./data/").glob("*.csv"):
        read_csv_auto(f).to_excel(writer, sheet_name=f.stem[:31], index=False)
```

## 输入解析规则

按以下优先级识别输入：

1. **目录**：遍历目录下所有 `.csv`
2. **单文件**：直接转换
3. **文件列表**：按逗号、顿号、分号、换行拆分；文件名中的空格会被保留；无效路径打印警告但不中断

## 输出命名规则

| 场景 | 输入 | 输出 |
|------|------|------|
| 单文件 | `data.csv` | `data.xlsx` |
| 指定输出文件名 | `data.csv` + `-o result.xlsx` | `result.xlsx` |
| 目录批量 | `./data/` 下有 `a.csv`, `b.csv` | `./data/a.xlsx`, `./data/b.xlsx` |
| 合并模式 | `a.csv, b.csv` + `--combine -o out.xlsx` | `out.xlsx`（含 a, b 两个 sheet，sheet 名使用 CSV 文件名） |
| 指定 sheet 名 | `data.csv` + `-n 销售数据` | `data.xlsx`（sheet 名为"销售数据"）；`-n` 仅在单文件模式下生效 |

## 编码处理

CSV 文件编码自动检测优先级：

1. **utf-8-sig**（带 BOM 的 UTF-8，Excel 默认保存的编码）
2. **utf-8**（纯 UTF-8）
3. **gb18030**（中文 Windows 常用编码）

若自动检测失败，可通过 `-e` 参数手动指定编码（支持任何 Python 内置编码名，如 `gb18030`、`latin1`、`shift_jis` 等）。

## 代码规范

- 使用 `pandas.read_csv()` 读取，`openpyxl` 输出
- 不输出 DataFrame 索引（`index=False`）
- 错误处理显式捕获，失败时打印错误信息并退出码 1
