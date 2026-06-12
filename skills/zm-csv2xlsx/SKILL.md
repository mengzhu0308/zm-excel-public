---
name: zm-csv2xlsx
description: >-
  将 CSV 文件转换为 Excel（.xlsx）格式。关键词：CSV转Excel、csv转xlsx、
  把CSV转成Excel、CSV导入Excel、批量转换CSV、多个CSV合并成一个Excel。
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
- 将多个 CSV 文件（以逗号、顿号、分号、换行分隔的列表）批量转换为 Excel
- 将多个 CSV 合并为一个 Excel，每个 CSV 作为一个 sheet
- 将 CSV 数据导入到 Excel 中
- 处理 CSV 编码问题（自动检测 utf-8-sig / utf-8 / gb18030）

## 不触发条件

以下场景**不**应触发本 skill，应改用其他专门 skill：

- Excel 内数据查询/筛选 → 使用 `zm-excel-query`
- 将 Excel 转 CSV（反向转换） → 使用 `zm-xlsx2csv`
- 跨表合并、跨文件合并 → 使用 `zm-excel-dedup-merge` 或 `zm-excels-merge`

> 注意：`formatting` / `chart` / `formula` 类 skill 在当前 `zm-excel` 项目中尚未提供；如需 Excel 格式化、图表或公式写入，请使用对应专业工具（如 openpyxl 原生 API）

## 核心原则

- **零修改源文件**：转换过程只读取 CSV，不修改源文件
- **不静默覆盖**：目标 xlsx 已存在时默认报错，需传 `--force` 才覆盖
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

# 文件列表（逗号、顿号、分号、换行混用分隔）
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" "a.csv, b.csv、c.csv"

# 指定输出路径
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv -o ./output/result.xlsx

# 强制覆盖已存在的输出文件
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv -o result.xlsx --force

# 多 CSV 合并为一个 Excel（每个 CSV 一个 sheet）
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" \
  "data1.csv data2.csv data3.csv" -o combined.xlsx --combine

# 指定 sheet 名
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv -n "销售数据"

# 无表头 CSV
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv --no-header

# 指定编码读取
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv -e gb18030

# 自定义编码检测行数
conda run -n agent-skills python "$SKILL_DIR/scripts/csv2xlsx.py" data.csv --encoding-detect-lines 200
```

### 方式二：AI 推理执行

需要自行实现转换逻辑时（不调用脚本），核心要素是：

- **读取**：按 `utf-8-sig → utf-8 → gb18030` 顺序试编码，避免大文件全量读取
- **写出**：`df.to_excel(out, sheet_name=..., index=False, engine="openpyxl")`；sheet 名非法字符替换为下划线、长度上限 31
- **批量**：遍历 `Path(dir).glob("*.csv")` 逐文件写出
- **合并**：用 `pd.ExcelWriter(path, engine="openpyxl")` 上下文，循环内 `df.to_excel(writer, sheet_name=...)`

完整可运行模板见 [README.md](README.md) "## 备选用法 > 两种入口差异" 段。

## 输入解析规则

按以下优先级识别输入：

1. **文件列表**：若输入字符串含列表分隔符（`,` / `，` / `、` / `;` / `；` / 换行），优先按文件列表解析
2. **目录**：不含分隔符时，若输入是已存在目录则遍历目录下所有 `.csv`
3. **单文件**：不含分隔符且不是目录时，按单文件处理
4. 文件列表中的无效路径会打印警告但不会中断处理

## 路径安全（--project-root）

启用 `--project-root <绝对路径>` 后（**必须是绝对路径**，相对路径则 `parser.error` 退出），所有输入（按 `collect_csv_files` 解析后的实际文件）和 `-o` 输出路径必须**解析后**位于该项目根目录之下；越界则 `parser.error` 退出（退出码 2）。

- 默认关闭（不传时完全不校验）
- 与批量/合并模式兼容：会先收集所有实际文件再逐个校验，不再误判文件列表输入
- 与 `-o` 一起使用时建议在备选用法/批处理场景显式传入
- 不替代用户文件系统权限，仅是"自报家门"式的输入域约束

## 输出命名规则

| 场景 | 输入 | 输出 |
|------|------|------|
| 单文件 | `data.csv` | `data.xlsx` |
| 指定输出文件名 | `data.csv` + `-o result.xlsx`（若 `-o` 无后缀且非目录，脚本会按文件名追加 `.xlsx` 并通过 stderr 打印提示） | `result.xlsx` |
| 目录批量 | `./data/` 下有 `a.csv`, `b.csv` | `./data/a.xlsx`, `./data/b.xlsx` |
| 合并模式 | `a.csv, b.csv` + `--combine -o out.xlsx`（**必须**配合 `-o` 指定输出文件路径） | `out.xlsx`（含 a, b 两个 sheet，sheet 名使用 CSV 文件名） |
| 指定 sheet 名 | `data.csv` + `-n 销售数据` | `data.xlsx`（sheet 名为"销售数据"）；`-n` 仅在单文件模式下生效 |

## 编码处理

CSV 文件编码自动检测优先级：

1. **utf-8-sig**（带 BOM 的 UTF-8，Excel 默认保存的编码）
2. **utf-8**（纯 UTF-8）
3. **gb18030**（中文 Windows 常用编码）

若自动检测失败，可通过 `-e` 参数手动指定编码（支持任何 Python 内置编码名，如 `gb18030`、`latin1`、`shift_jis` 等）。自动检测失败时会打印"建议使用 -e 手动指定编码"并退出（退出码 1）；更多 `-e` / `--encoding` 参数语义与边界示例见 [README.md](README.md) "## 备选用法 > -e/--encoding" 段。
