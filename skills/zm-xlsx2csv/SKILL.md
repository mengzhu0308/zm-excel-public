---
name: zm-xlsx2csv
description: >-
  将 Excel 文件（.xlsx/.xls/.xlsm）转换为 CSV。批量或单个文件转换，支持自动编码检测，多 sheet 时各 sheet 输出为独立 CSV。
license: MIT
metadata:
  skill_mode: workflow
compatibility:
  runtime:
    - name: agent-skills
      call_command: "conda run -n agent-skills python \"$SKILL_DIR/scripts/excel2csv.py\" [args]"
---

# zm-xlsx2csv

将 Excel 文件转换为 CSV。支持单文件、目录批量、文件列表三种输入，多 sheet 时每个 sheet 输出独立 CSV。

## 触发条件

- 单个 Excel 文件转 CSV
- 目录下所有 Excel 文件批量转 CSV
- 多个 Excel 文件（空格/逗号/顿号/分号/换行分隔的列表）批量转 CSV
- 指定 sheet 导出为 CSV
- 从 Excel 提取数据为纯文本 CSV

## 不应触发

涉及以下场景时，请改用对应 skill，不要使用本 skill：

- 数据查询、筛选、条件取值 → `zm-excel-query`
- 样式、字体、颜色、图表、公式写入 → `zm-excel-formalization`
- 合并多个 Excel 文件（横向/纵向）→ `zm-excels-merge`
- 排序 / 去重合并 → `zm-excel-sort` / `zm-excel-dedup-merge`
- 增删单行/多行 → `zm-excel-add-one-row` / `zm-excel-del-multi-rows`

## 核心原则

- **零修改源文件**：只读取 Excel，不修改源文件
- **多 sheet 自动拆分**：多 sheet 时每个 sheet 输出为 `原文件名_Sheet名.csv`；单 sheet 输出为 `原文件名.csv`
- **灵活输入解析**：自动识别单文件、目录、文件列表三种输入形式
- **编码安全**：输出 CSV 使用 `utf-8-sig` 编码，确保 Excel 打开中文不乱码
- **不输出索引**：CSV 不含 DataFrame 行索引
- **类型保守**：默认按字符串读取，避免前导零 / 工号被自动推断丢失
- **不静默覆盖**：默认跳过已存在的同名 CSV，需 `--overwrite` 才覆盖

## 使用方式

### 方式一：脚本调用（推荐）

`${SKILL_DIR}` 由调用本 skill 的运行时（Codex CLI / Claude Code 等）注入，指向 skill 安装根目录；纯 shell 环境请用绝对路径替换。

```bash
# 单文件
conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" data.xlsx

# 目录批量
conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" ./data/

# 指定 sheet
conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" data.xlsx -s "Sheet2"

# 文件列表（空格/逗号/顿号/分号/换行分隔均可）
conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" "a.xlsx, b.xlsx、c.xlsx"

# 指定输出目录
conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" data.xlsx -o ./output/

# 强制覆盖已有 CSV
conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" data.xlsx --overwrite

# 目录递归
conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" ./data/ --recursive

# 严格模式（任何失败 → exit 1）
conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" ./data/ --strict

# 单 sheet 读取超时（秒）
conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" data.xlsx --timeout 30
```

AI 推理执行时，若用户需求明确且参数可一次性确定，优先使用脚本调用。

完整 runtime 模板与 `$SKILL_DIR` 来源见 [`references/runtime-env.md`](references/runtime-env.md)。

### 方式二：脚本回退（无独立 AI 推理实现）

本 skill 不提供"脱离脚本的 AI 推理实现"——所有能力（多 sheet 拆分、非法字符替换、批量容错、覆盖保护、原子写入）都依赖 `scripts/excel2csv.py`。如确需在无法调用脚本的环境中执行，请改用 `python -c '...'` 内联调用脚本入口（见方式一）。

## 输入解析规则

脚本按以下优先级识别输入：

1. **目录**：路径存在且为目录时，遍历所有 `.xlsx` / `.xls` / `.xlsm` 文件
2. **单文件**：路径存在且为 Excel 文件时，直接转换
3. **文件列表**：按空格、逗号、顿号、分号、换行拆分，逐个识别有效 Excel 文件；无效路径打印警告但不中断

## 输出命名规则

| 场景 | 输入 | 输出 |
|------|------|------|
| 单文件，单 sheet，未指定 sheet | `sales.xlsx`（1 个 sheet） | `sales.csv` |
| 单文件，多 sheet，未指定 sheet | `sales.xlsx`（Sheet1, Sheet2） | `sales_Sheet1.csv`, `sales_Sheet2.csv` |
| 单文件，指定 sheet | `data.xlsx` + `-s Sheet2` | `data_Sheet2.csv` |
| 目录批量 | `./data/` 下有 `a.xlsx`, `b.xlsx` | 每个 sheet 对应一份 CSV；命名遵循上述规则 |
| 指定输出目录 | `data.xlsx` + `-o ./out/` | `./out/data.csv`（或带 sheet 后缀） |

**单一真相来源（伪代码）**：

```text
if sheet_name is None and len(sheet_names) == 1:
    out_name = "{base_stem}.csv"
else:
    out_name = "{base_stem}_{safe_sheet_name}.csv"
```

任何修改必须同步更新此表与脚本 `convert_single_file` 中的 `single_sheet_no_spec` 判断。

## 实现规范

实现规范（读取 / 写入 / 命名 / 批量容错 / 退出码）见 [`references/implementation.md`](references/implementation.md)。

## 排错参考

完整排错表见 [`references/runtime-env.md` 故障排查](references/runtime-env.md#故障排查)。以下是常见速查：

- 中文乱码 → 确认 CSV 头三字节为 `EF BB BF`（utf-8-sig BOM）
- 前导零丢失 → 确认 `dtype=str` 生效（脚本已默认）；如脱离脚本手动调用 `pd.read_excel`，需自行加 `dtype=str`
- `.xls` 读取失败 → 提示安装 `xlrd<2.0`；当前 conda `agent-skills` 已含
- 批处理意外中断 → 检查脚本是否支持"批量跳过失败"（见 [`references/implementation.md`](references/implementation.md)）；如使用旧版，可手动将每个文件单独跑
- 公式 cell 值陈旧 → openpyxl 不计算公式，CSV 反映 Excel 上次保存时的值；如需最新值请先在 Excel 中 `Ctrl+Alt+F9` 强制重算后保存
- 隐藏文件被误读 → Linux/macOS 上 `.hidden.xlsx` 与普通文件等价，会被 `--recursive` 拾取；Windows 上由文件系统隐藏属性决定，可先在资源管理器取消隐藏再跑
- `$SKILL_DIR` 未定义 → 调用方未注入；改用绝对路径或从 `references/runtime-env.md` 查回退
