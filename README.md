# zm-excel-public

`zm-excel-public` 是一个面向 Codex CLI/Claude Code 的 Excel 处理 Agent Skill 集合。装好之后，你可以按任务直接调用不同 skill，处理 Excel 格式转换、数据查询、表格合并、格式化排版等常见工作。

本文档默认由 `project-write-readme/` 自动生成和维护。

## 项目目标

- 帮你把 `zm-excel-public` 这组常用 Agent Skill 一次安装到本机，减少到处找零散配置的时间。
- 帮你按任务快速找到合适的 skill，而不是先翻一圈仓库结构和说明文档。
- 让你在不依赖 skill 级 README 的情况下，也能先跑通第一次使用。

## 安装与使用

如果你是第一次使用这套 skill，先安装到本机，再按任务挑一个最接近的 skill 开始用。若只是想把已经装过的一批 skill 清掉，也可以直接走卸载入口。

### 常用入口

- 先用 `project-install/main.py` 把这套 skill 安装到本机；这是第一次使用时最常见的入口。
- 如果你只是想清理本机运行态，直接用 `project-uninstall/main.py` 反向卸载就行。
- 安装完成后，直接看下面的 `Skills 用途一览` 和 `附录：各 Skill Prompt 示例`，就能开始试用。
- 如果你暂时拿不准该用哪个 skill，优先看“更适合你在什么时候用”这一列。

### 脚本运行环境

#### `project-*` 项目级脚本

- `project-*` 下的项目级脚本沿用统一优先级：`uv run python`（项目存在 `pyproject.toml` / `uv.lock` 时）> `python`（已激活 conda 环境）> `python3`。
- 这套入口规则只用于根级 `project-*` 自动化脚本，不直接外推到 skill 自身的脚本回退入口。
- 安装 skill 到本机时，统一写入 `~/.agent-skills/.zm/`，并把本轮选中的工具入口同步到顶层 `skills/` 目录。
- 卸载时沿用同一套目标规则：先清理工具顶层入口，再删除 `~/.agent-skills/.zm/` 中的实际 skill 目录。

### 安装到本机

```bash
# uv 环境
uv run python project-install/main.py
uv run python project-install/main.py --skill zm-excel-query --skill zm-excel-dedup-merge
uv run python project-install/main.py --pattern 'zm-excel-*'

# conda 环境
python project-install/main.py
python project-install/main.py --skill zm-excel-query --skill zm-excel-dedup-merge
python project-install/main.py --pattern 'zm-excel-*'

# 系统级
python3 project-install/main.py
python3 project-install/main.py --skill zm-excel-query --skill zm-excel-dedup-merge
python3 project-install/main.py --pattern 'zm-excel-*'
```

### 从本机卸载

```bash
# uv 环境
uv run python project-uninstall/main.py
uv run python project-uninstall/main.py --skill zm-excel-query --skill zm-excel-dedup-merge
uv run python project-uninstall/main.py --pattern 'zm-excel-*'

# conda 环境
python project-uninstall/main.py
python project-uninstall/main.py --skill zm-excel-query --skill zm-excel-dedup-merge
python project-uninstall/main.py --pattern 'zm-excel-*'

# 系统级
python3 project-uninstall/main.py
python3 project-uninstall/main.py --skill zm-excel-query --skill zm-excel-dedup-merge
python3 project-uninstall/main.py --pattern 'zm-excel-*'
```

## Skills 用途一览

这张表更关心“你遇到什么任务时该想到它”，而不是展示仓库内部文件怎么组织。

| Skill | 更适合你在什么时候用 | 能帮你做什么 | 分类 |
| --- | --- | --- | --- |
| zm-csv2xlsx | `zm-csv2xlsx` 将 CSV 文件转换为 Excel 格式（.xlsx）。 | `zm-csv2xlsx` 将 CSV 文件转换为 Excel 格式（.xlsx）。支持单文件、目录批量、文件列表三种输入方式；单文件可指定输出路径和 sheet 名，多文件可合并到一个 Excel 的不同 sheet 中。自动检测 utf-8-sig / utf-8 / gb18030 编码，确保中文内容不丢失。 | 工作流 |
| zm-excel-add-one-row | 将 Excel 表格的字段提取为 Markdown 填写模板，用户填写后自动追加一行到表格末尾。 | 将 Excel 表格的字段提取为 Markdown 填写模板，用户填写后自动追加一行到表格末尾。支持预览确认与多工作表，适用于登记表、销售表、库存表、项目表、考勤表等数据录入场景。 | 工作流 |
| zm-excel-dedup-merge | `zm-excel-dedup-merge` 用于把两个 CSV / XLSX 表格按关键列去重合并：同名记录合并为一行，保留两表全部列，并自动处理列名冲突与关键列空值。 | `zm-excel-dedup-merge` 用于把两个 CSV / XLSX 表格按关键列去重合并：同名记录合并为一行，保留两表全部列，并自动处理列名冲突与关键列空值。 | 工作流 |
| zm-excel-del-multi-rows | 按关键词搜索并删除 Excel 或 CSV 中匹配的多行数据，结果输出到新文件，原始文件不会被修改。 | 按关键词搜索并删除 Excel 或 CSV 中匹配的多行数据，结果输出到新文件，原始文件不会被修改。支持多关键词、子串/精确匹配、大小写敏感开关、多 sheet 全处理或指定单个 sheet，以及先预览再执行的干跑模式。 | 工作流 |
| zm-excel-formalization | 将现有 Excel 文件按规范格式化为正式文档样式。 | 将现有 Excel 文件按规范格式化为正式文档样式。自动识别单元格内容的中英文字符并分别应用宋体与 Times New Roman，设置垂直居中与自动换行，自动按内容调整列宽至最优，支持单文件与批量处理。 | 工作流 |
| zm-excel-query | `zm-excel-query` 从 Excel 文件中查询数据，支持条件筛选、列选择、排序、去重、分组聚合等操作，将结果保存为 CSV 或 XLSX 文件。 | `zm-excel-query` 从 Excel 文件中查询数据，支持条件筛选、列选择、排序、去重、分组聚合等操作，将结果保存为 CSV 或 XLSX 文件。源 Excel 只读，不会被修改。输出文件与源 Excel 同目录、同名（扩展名由格式参数决定，默认 `.csv`）。 | 工作流 |
| zm-excel-sort | `zm-excel-sort` 是一个**混合型** skill，支持通过自然语言触发 AI 执行排序，也提供命令行脚本作为回退入口。 | `zm-excel-sort` 是一个**混合型** skill，支持通过自然语言触发 AI 执行排序，也提供命令行脚本作为回退入口。它能对 Excel（`.xlsx`/`.xlsm`）和 CSV 文件按字段规则排序，XLSX 输出完整保留原始单元格样式。 | 工作流 |
| zm-excels-merge | 合并多个 Excel（.xlsx / .xls / .xlsm）和 CSV 文件，支持按后缀分组、默认首 sheet、字段兼容性分析、智能分组和合并清单工作流。 | 合并多个 Excel（.xlsx / .xls / .xlsm）和 CSV 文件，支持按后缀分组、默认首 sheet、字段兼容性分析、智能分组和合并清单工作流。 | 工作流 |
| zm-xlsx2csv | 将 Excel 文件（.xlsx / .xls / .xlsm）转换为 CSV 格式。 | 将 Excel 文件（.xlsx / .xls / .xlsm）转换为 CSV 格式。支持单文件、目录批量、文件列表三种输入模式，多 sheet 时每个 sheet 输出为独立 CSV。 | 工作流 |

## 推荐工作流

下面这些路径更像“第一次用的时候该怎么起手”，不是维护仓库时的内部流程。

- 当前仓库还没有命中预设的使用路径；你可以先从 `Skills 用途一览` 和附录里的起手 Prompt 开始。

## 附录：各 Skill Prompt 示例

下面这些示例尽量直接沿用各个 skill README 里的“用法”代码块；你可以直接复制，再按自己的任务改几个关键词。

### zm-csv2xlsx

```
请使用 zm-csv2xlsx skill 将 CSV 文件转换为 Excel
输入：单个 CSV 文件路径（如 data.csv）
输出：与源文件同名的 xlsx 文件
另外，还有下列参数约束：
- 合并模式：默认关闭，启用后多个 CSV 合并到一个 Excel 不同 sheet 中
- 编码：默认自动检测（utf-8-sig / utf-8 / gb18030），也可手动指定
- sheet 名：单文件时默认使用文件名，可自定义
- 表头：默认 CSV 包含表头行，可指定无表头
```

### zm-excel-add-one-row

```
请使用 zm-excel-add-one-row skill 在 Excel 表格末尾新增一行
输入：./employees.xlsx
输出：自动生成 employees_增加一行.xlsx（原文件保持不变；同名输出已存在时自动递增到 _2、_3…直至 _9999）
另外，还有下列参数约束：
- 工作表选择：不指定时处理活动工作表；多工作表文件可用 --sheet 指定目标工作表
- 输出路径：默认落在 Excel 同目录；可用 --output / -o 自定义 Markdown 模板或回填结果路径
- 强制覆盖：默认拒绝覆盖已存在的模板或同名输出文件；可用 --force 显式允许
- 预览确认：可用 `add_one_row.py write --dry-run` 预览待追加数据后再正式写入
```

### zm-excel-dedup-merge

```
请使用 zm-excel-dedup-merge skill 把两个表格按关键列去重合并
输入：tableA.csv、tableB.xlsx
输出：当前工作目录下的 dedup_merged.xlsx（CSV+CSV→CSV，XLSX+XLSX→XLSX，混合→XLSX）
另外，还有下列参数约束：
- 关键列：默认自动检测同名列；如需手动指定请告知列名
- 模糊匹配：默认 academic 预设（适合英文学术期刊名册）；中文学术表与产品清单用 cjk，要求 0 误匹配用 none
- 预览建议：建议先以 dry-run 模式预览匹配统计，确认后再正式合并
```

### zm-excel-del-multi-rows

```
请使用 zm-excel-del-multi-rows skill 帮我删除 Excel/CSV 中匹配关键词的行
输入：/path/to/data.xlsx（Excel .xlsx/.xlsm 或 CSV .csv）
输出：同目录下自动命名为 data_删除多行.xlsx 的新文件
另外，还有下列参数约束：
- 关键词："作废"、"删除"（匹配任一关键词即删除该行）
- 匹配模式：子串匹配（默认），即单元格内容包含关键词即命中
- sheet：不指定，默认处理所有 sheet（仅 Excel 有效）
```

### zm-excel-formalization

```
请使用 zm-excel-formalization skill 格式化以下 Excel 文件
输入：/path/to/report.xlsx
输出：同目录保存为 report_副本.xlsx
另外，还有下列参数约束：
- 列宽调整：默认自动按内容调整列宽；如需保留原文件的列宽设置，请显式声明禁用
```

### zm-excel-query

```
请使用 zm-excel-query skill 查询以下 Excel 文件，筛选出满足条件的数据并导出为 CSV
输入：./sales.xlsx
输出：同目录下 sales.csv
另外，还有下列参数约束：
- 筛选条件：region == 'East' and amount > 1000
- 列选择：date, region, product, amount
- 排序规则：amount desc
- 工作表：默认第一个工作表；多工作表时可指定名称
- 预览：先查看数据结构（列名、类型、示例值），确认后再查询
- 交互模式：逐步验证查询逻辑，确认结果后再保存
- 查询标识：EastHigh（对同一文件多次查询时，用标签区分不同结果）
- 输出格式：csv（默认，可选 xlsx）
- 排除字段：序号
- 表头偏移：1（表头在第2行时指定）
```

### zm-excel-sort

```
请使用 zm-excel-sort skill 对 /path/to/sales_data.xlsx 按销售额降序排列。
输入：/path/to/sales_data.xlsx（Excel 文件）
输出：/path/to/sales_sorted.xlsx（XLSX 格式）
另外，还有下列参数约束：
- 排序规则：按"销售额"字段从高到低排列
- 目标 Sheet：Sheet1（Excel 多 Sheet 时指定，默认第 1 个 Sheet）
- 输出格式：xlsx（也可选择 csv，默认从输出文件扩展名自动推断）
- 空值位置：last（默认空值排最后，可选 first 将其放最前）
```

### zm-excels-merge

```
请使用 zm-excels-merge skill 合并以下 Excel 和 CSV 文件
输入：path/to/your/data 目录下的所有 .xlsx、.xls、.xlsm 和 .csv 文件
输出：path/to/your/merged.xlsx
另外，还有下列参数约束：
- 递归搜索：False
- 文件匹配模式：*.xlsx,*.xls,*.xlsm,*.csv
- 指定 sheet 名：全部
- 表头行号：0
- 添加来源列：False
- 来源列名称：来源文件
```

### zm-xlsx2csv

```
请使用 zm-xlsx2csv skill 将 Excel 文件转换为 CSV 格式。
输入：`data.xlsx`
输出：同目录下的 `data.csv`
另外，还有下列参数约束：
- 指定 sheet：不指定时，转换所有 sheet；每个 sheet 输出为独立 CSV
```
