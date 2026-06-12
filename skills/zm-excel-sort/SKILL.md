---
name: zm-excel-sort
description: >-
  对 Excel/CSV 文件按指定字段规则排序。支持多字段排序、升序/降序、数据类型自动识别、自定义排序顺序。输出为 CSV 或 XLSX，XLSX 完整保留原始单元格样式（字体、填充、边框、对齐、列宽、行高、合并单元格、冻结窗口[按当前坐标复制，排序后引用可能错位]、表格[按当前坐标复制，排序后引用可能错位]、命名区域[按当前坐标复制，排序后引用可能错位]、数据验证[按当前坐标复制，排序后 sqref 引用可能错位]、条件格式[基于固定区域，排序后可能错位]）。

  当用户要求对 Excel/CSV/XLSX 数据进行排序（按某列排、从高到低、整理顺序、把某类放前面等）时，**务必优先使用此 skill**。

  **不要**直接用 pandas/openpyxl 手写排序脚本——本 skill 能正确处理混合数据类型、保留 XLSX 原始样式，并支持自定义排序顺序（如优先级 P0→P1→P2），手写脚本会丢失格式且容易出错。

  **输入限制**：仅支持 .xlsx / .xlsm / .csv。.xls 不在本 skill 支持范围内（openpyxl 限制），请先在 Excel/WPS/LibreOffice 中另存为 .xlsx 再调用。
metadata:
  skill_mode: hybrid
  version: 0.1.3
compatibility:
  runtime:
    - name: agent-skills
      call_command: "conda run -n agent-skills python \"$SKILL_DIR/scripts/sort_excel.py\" [args]"
---

# zm-excel-sort

## 概述

- **多字段排序**：多个字段依次排序，每字段独立指定方向
- **数据类型感知**：自动识别数字、日期、文本，按对应语义排序
- **自定义排序**：支持非字母/数字的自定义顺序（如优先级、状态流转）
- **格式保留**：XLSX 输出保留原始单元格样式（字体、颜色、边框、列宽）
- **空值处理**：可配置空值排在最前或最后

## 与 /goal 配合使用

`/goal` 是会话层的任务跟踪能力（Codex CLI、Claude Code 等平台均提供），不是本 skill 的脚本参数。

**最小判据——以下情况一般不需要 `/goal`：** 单字段、单方向、单 Sheet、规则明确（无需多轮确认）、不在批量处理流程中。

**建议使用 `/goal` 的场景：** 自定义顺序较长（≥3 项）、批量处理多个文件、跨会话续跑、规则需多轮与用户对齐、需要在失败时回滚到中间产物。`/goal` 至少应包含：输入文件、目标 Sheet、结构化排序规则、输出格式与完成标准。

## 输入规范

### 输入文件

1. **Excel**（`.xlsx` / `.xlsm`）：需指定 Sheet 名称或索引（默认第 1 个）。注意：本 skill 不支持 `.xls`（openpyxl 限制）
2. **CSV**（`.csv`）：编码自动检测，优先 UTF-8，失败时回退常见中文编码；仅在 `latin1` 兜底时强制 stderr 提示乱码风险，正常编码在 `--verbose` 模式回显到 stderr

### 排序规则

用户用自然语言描述需求，AI 解析为结构化 JSON 后传递给脚本。

```json
{
  "columns": [
    {
      "name": "字段名",
      "direction": "asc" | "desc",
      "custom_order": ["可选", "自定义", "顺序"]
    }
  ],
  "null_position": "last" | "first",
  "case_sensitive": false | true
}
```

#### 解析示例

| 用户输入 | 解析后的规则 |
|---------|------------|
| "按销售额降序" | `[{"name":"销售额","direction":"desc"}]` |
| "先按地区升序，再按销售额降序" | `[{"name":"地区","direction":"asc"},{"name":"销售额","direction":"desc"}]` |
| "按状态：待处理→处理中→已完成" | `[{"name":"状态","direction":"asc","custom_order":["待处理","处理中","已完成"]}]` |

> **边界行为**：未在 `custom_order` 列表中的值，按 pandas Categorical 默认行为排在已定义值之后（即 NaN/未分类值统一置后，且不影响其他字段排序）。若需要在自定义顺序中包含所有可能值，应将所有取值显式列入 `custom_order` 列表。

#### 字段名匹配

- AI 层（`agents/openai.yaml` 描述）负责字段名模糊匹配（如 "销售额" 可匹配 "销售金额"）
- 脚本（`sort_data`）只做严格匹配：列名与表头不一致时直接 `ValueError`
- 匹配失败时由脚本报错，AI 据此回退并向用户确认，不擅自猜测

## 执行流程

### 步骤 1：解析排序规则

将自然语言需求转为结构化 JSON。需确认：

- 目标字段是否存在（模糊匹配）
- 排序方向是否明确（未明确时先向用户确认，不擅自默认升序）
- 是否有自定义排序需求、空值位置偏好

规则存在歧义时先向用户确认，不擅自决定——与 `agents/openai.yaml` system 段"排序规则未明确或存在歧义时先向用户确认"口径一致。

### 步骤 2：调用排序脚本

```bash
conda run -n agent-skills python "$SKILL_DIR/scripts/sort_excel.py" \
  --input "输入文件路径" \
  --output "输出文件路径" \
  --rules '{"columns":[...],"null_position":"last"}' \
  [--sheet "Sheet名称"] \
[--format csv|xlsx]
```

> 上述命令使用统一的 `agent-skills` runtime：`conda run -n agent-skills python "$SKILL_DIR/scripts/sort_excel.py"` —— 该环境已预装 `openpyxl` 与 `pandas`。

**输出格式**默认从 `--output` 扩展名推断（`.csv` → csv；`.xlsx` / `.xlsm` → xlsx）；仅在显式需要时使用 `--format`，且传入值必须与扩展名一致（冲突时脚本报 `ValueError` 并退出）。如需将 XLSX 强制输出为 CSV（或反之），请改写 `--output` 路径的扩展名，去掉 `--format` 让脚本按扩展名自动推断。

### 步骤 3：验证结果

检查输出文件是否存在、数据行数是否与输入一致（仅顺序改变，无增删）、排序结果是否符合规则。

## 输出规范

### XLSX 输出

- 保留原始单元格样式（字体、填充色、边框、对齐方式）
- 保留列宽、行高
- 保留冻结窗口（freeze_panes）
- 保留批注（cell comment）与超链接（cell hyperlink）
- 保留条件格式（**注意**：条件格式基于固定区域，排序后可能错位；脚本会原样复制，并在 verbose 模式下提示"存在错位风险"）
- 保留单行合并（自动跟随行重排）；多行合并仅在排序后该组行仍连续时整体平移，否则跳过并日志提示；跨表头+首行合并（A1:A2 类）按"首行跟随重排、整体保持表头在第 1 行"处理
- 保留数据验证（data validation）：按当前坐标复制到新工作表；sqref 引用基于固定行号，排序后该规则可能不再对应原行，**请人工评估是否需要重新设定**
- 不保留项：工作表保护（sheet protection）、打印设置（print options）、宏（macro）、外部链接（external links）等

### CSV 输出

- 编码：UTF-8 with BOM
- 分隔符：逗号，含逗号字段用双引号包裹
- 表头：保留原始表头
- 换行符：LF

## 错误处理

| 场景 | 处理方式 |
|------|---------|
| 输入文件不存在 | 报错并提示检查路径 |
| 输入为 `.xls` | 直接拒绝（提示仅支持 .xlsx/.xlsm，请先另存为 .xlsx） |
| `--format` 与 `--output` 扩展名冲突 | 报错并提示明确选择（请保持一致） |
| CSV 输入附带 `--sheet` 参数 | 报错（CSV 没有 Sheet 概念，请去掉 `--sheet`） |
| 表头含重复列名 | 报错并指出重复列名 |
| 指定 Sheet 不存在 | 列出可用 Sheet 名称供选择；按名匹配优先，回退到数字索引 |
| 排序字段不存在 | 列出所有可用字段，提示确认 |
| 排序规则 JSON 解析失败 | 报错并保留原始 json 错误信息 |
| 排序规则 schema 错误（缺 `columns`、`columns=[]`、非法 `direction`、非布尔 `case_sensitive`、非法 `null_position`） | 提前拒绝并打印明确错误 |
| `custom_order` 含重复值 | 提前拒绝并指出重复值 |
| 数据类型混合（如文本和数字混在同一列） | 走 pandas object 列默认排序（字典序），并通过 stderr 提示用户——脚本不显式做类型转换，行为取决于 pandas |
| CSV 字段以 `=` / `+` / `-` / `@` / Tab / CR 开头 | 写入时自动加 `'` 前缀，避免 Excel/WPS 打开时被强制执行公式 |
| 排序执行失败 | 报错并保留原始异常信息 |
| 输出写入失败 | 报错并保留原始异常信息 |
| 空文件或仅有表头 | 直接输出空结果，不报错；空 XLSX 默认 Sheet 名 `"Sheet1"` |
| 输出路径不可写 | 自动创建父目录；其他失败时报错并建议其他路径 |

## 脚本接口规范

```
--input,  -i    输入文件路径（必填）
--output, -o    输出文件路径（必填）
--rules,  -r    排序规则 JSON（必填）
--sheet,  -s    Sheet 名称或索引（Excel 时可选，默认 0）
--format, -f    输出格式：csv 或 xlsx（默认从扩展名推断）
--verbose, -v   显示详细日志
```

## 示例

### 单字段排序

用户：把 `sales_data.xlsx` 的 Sheet1 按销售额从高到低排，输出 `sorted.xlsx`

```bash
conda run -n agent-skills python "$SKILL_DIR/scripts/sort_excel.py" \
  --input sales_data.xlsx --sheet Sheet1 --output sorted.xlsx \
  --rules '{"columns":[{"name":"销售额","direction":"desc"}],"null_position":"last"}'
```

### 多字段排序 + 自定义顺序

用户：把 `orders.csv` 按状态（未付款→已付款→已发货→已完成）排，同状态下按下单时间从早到晚

```bash
conda run -n agent-skills python "$SKILL_DIR/scripts/sort_excel.py" \
  --input orders.csv --output sorted.csv \
  --rules '{"columns":[{"name":"状态","direction":"asc","custom_order":["未付款","已付款","已发货","已完成"]},{"name":"下单时间","direction":"asc"}],"null_position":"last"}'
```
