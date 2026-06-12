---
name: zm-excel-add-one-row
description: >-
  将 Excel 列名提取为 Markdown 填写模板，用户填完后自动回填到表格末尾新增一行。
  当用户提到填表、加一行、录入数据、追加记录、补一条数据、往 Excel 写内容，
  或已有表格需要新增一行时，使用此 skill。
  适用于登记表、销售表、库存表、项目表、考勤表等数据录入场景。
metadata:
  skill_mode: hybrid
compatibility:
  runtime:
    - name: agent-skills
      call_command: 'conda run -n agent-skills python "$SKILL_DIR/scripts/add_one_row.py" [args]'
      requires:
        - openpyxl>=3.1
---

# zm-excel-add-one-row

将 Excel 字段提取为 Markdown 模板，用户填写后自动追加一行到表格末尾。

## 依赖

- **Python**：3.9 及以上
- **openpyxl**：`>= 3.1`（读取与写入 `.xlsx` / `.xlsm`；**不支持 `.xls`**，遇到 `.xls` 输入会直接报"Unsupported file extension"并退出码 1；如需处理 `.xls` 请先用其他工具转换为 `.xlsx`）
- 安装方式：`pip install "openpyxl>=3.1"`
- conda 路径（与 `compatibility.runtime[0].call_command` 一致）：`conda create -n agent-skills python=3.9 && conda run -n agent-skills pip install "openpyxl>=3.1"`

## 数据安全承诺

- **不修改源 Excel**；新行始终写入新文件
- 多次追加同源文件时，输出文件名自动递增（`_增加一行.xlsx` → `_增加一行_2.xlsx` → …），**不会覆盖前一次的结果**
- 模板文件默认拒绝覆盖（`--force` 显式允许），保护 `/goal` 续跑与跨会话恢复时的用户填写内容

## 工作流程

### 阶段一：提取模板

读取 Excel，自动从前 3 行中挑出非空单元格最多的那一行作为表头；再从后往前找“至少一半表头列有值”的最后一行作为示例值，生成 `<原文件名>_row_template.md`：
- 每个字段名为一级标题 `#`
- 示例值用 HTML 注释 `<!-- 示例: ... -->`
- 工作表名记录为 `> 工作表: `SheetName``
- 用户将 `[在此填写]` 替换为实际值
- 多工作表文件可用 `--sheet` 指定
- 可用 `--output` / `-o` 自定义 Markdown 模板的输出路径（不指定时落在 Excel 同目录）
- 若目标模板文件已存在，默认拒绝覆盖以保护用户填写内容；需用 `--force` 显式允许覆盖

### 阶段二：用户填写

打开模板文件，在各字段下填入实际内容。

### 阶段三：预览确认

运行 `conda run -n agent-skills python "$SKILL_DIR/scripts/add_one_row.py" write <excel> <template.md> --dry-run` 预览待追加数据。用户检查字段、值、留空项，确认无误后再执行写入。

### 阶段四：写入 Excel

去掉 `--dry-run` 执行正式写入。新行追加到已有数据之后，**原文件不变**，在同目录生成 `<原文件名>_增加一行.xlsx`。如果该输出文件已存在（例如上一轮追加未改源文件），脚本会自动改名为 `<原文件名>_增加一行_2.xlsx`，依此类推，**保证前次追加的行不被覆盖**。自动递增上限为 `_增加一行_9999.xlsx`；超过此上限时脚本会提示"已达自动递增上限 9999，请清理目标目录或改用 `--output` 显式指定"，并以退出码 2 兜住。新行自动复制末行样式。

如需写到 Excel 同目录之外的位置，可用 `--output` / `-o` 显式指定 xlsx 输出路径；指定的输出文件已存在时需要 `--force` 才覆盖。`--force` 在 extract / write 的全部边界细节见下一节表格。

### `--force` 在 extract / write 的边界

`extract` 与 `write` 的 `--force` 边界不同，使用前请明确：

| 子命令 | `--force` 默认值 | `--force=True` 时行为 |
| --- | --- | --- |
| `extract` | `False`：目标模板文件已存在则抛 `FileExistsError` 并退出码 2 | 允许覆盖目标模板文件（会清空用户已填写内容） |
| `write` 走 `--output` 路径 | `False`：显式输出文件已存在则抛 `FileExistsError` 并退出码 2 | 允许覆盖显式输出文件 |
| `write` 走 auto-increment 路径 | `False`：`<原文件名>_增加一行.xlsx` 已存在则递增到 `_2`、`_3`… | **不生效**：脚本会继续递增，**不会**覆盖第一个 `_增加一行.xlsx`；如需覆盖请改用 `--output` 显式指定路径 |

简言之：`extract --force` 用于"我同意重新生成模板"，`write --force` 用于"我同意覆盖显式输出文件"，两者**不会**触发 auto-increment 分支的覆盖。

## 与 /goal 配合使用

`/goal` 是会话层的任务跟踪能力（Codex CLI、Claude Code 等平台均提供），不是本 skill 的脚本参数。普通单次填表可直接按四阶段工作流执行；当填写内容需要用户补充、多次确认、跨会话继续时，建议先开启 `/goal`。

- **目标记录**：写清源 Excel、目标工作表、要新增的业务记录、模板路径和输出文件验收标准。
- **阶段检查点**：按“提取模板 → 用户填写 → dry-run 预览 → 用户确认 → 正式写入”推进。
- **恢复点**：续跑时优先读取已生成的 Markdown 模板和 `/goal` 中记录的工作表选择，不重新生成覆盖用户填写内容。
- **完成条件**：输出 Excel 存在，新增行字段和值与 dry-run 预览一致，源文件保持不变。

## 模板格式

`extract_template` 输出的 Markdown 模板有两级标题：

```markdown
## Excel 数据录入模板

> 来源文件: `<filename>.xlsx`
> 工作表: `<SheetName>`

# 字段名 1
<!-- 示例: 示例值 1 -->
[在此填写]

# 字段名 2
<!-- 示例: 示例值 2 -->
[在此填写]
```

- `## Excel 数据录入模板`：模板级标题，由 `extract_template.py` 生成时固定写入一份（标识"这是回填模板"）
- `# 字段名`：每个 Excel 表头对应一个一级标题，**这才是脚本解析字段的依据**

回填脚本（`write_back.parse_markdown`）只识别一级标题 `#` 来匹配字段；模板级 `##` 不参与字段解析，仅作视觉锚点。手动维护模板时不要把字段名错写成 `##`——脚本会把它当成模板标题忽略，导致整列无法匹配。

替换 `[在此填写]` 为实际内容，保留字段名（一级标题）和示例注释。

## 工作表选择

优先级：**`--sheet` 参数** > **模板记录的工作表名** > **活动工作表**。

若 `--sheet` 指定的工作表不存在，报错并列出所有可用名称。

## 字段匹配

- 按 `# 字段名` 与 Excel 表头匹配（去空格，大小写敏感）
- 未填写（`[在此填写]` 或空）则留空
- 不存在于表头的字段被忽略并提示用户

## 使用示例

**示例 1：提取模板**
用户提供 Excel 文件路径，读取表头和末行示例，生成 `<原文件名>_row_template.md`。多工作表文件可用 `--sheet` 指定目标工作表。

**示例 2：回填数据**
用户填好模板后：
1. `add_one_row.py write --dry-run` 预览待追加数据
2. 用户确认无误后去掉 `--dry-run` 执行写入
3. 生成 `<原文件名>_增加一行.xlsx`，新行复制末行样式

**示例 3：覆盖工作表**
`add_one_row.py write data.xlsx template.md --sheet Sheet2`：命令行 `--sheet` 覆盖模板记录的工作表名。

## 注意事项

- 不修改原文件，写入时生成 `<原文件名>_增加一行.xlsx`
- 新行自动复制末行样式（字体、对齐、边框、填充、数字格式、保护、**行高**）
- 数字字符串的推断结果与示例行类型挂钩：示例行是数字且新值能无损转换时才推断；超长数字（`> LONG_DIGIT_THRESHOLD`、脚本中默认 10）、带前导零的纯数字、含非数字字符的值都保持字符串，日期保持字符串；`LONG_DIGIT_THRESHOLD` 是 `_convert_value` 内单一真相来源
- Markdown 中的 `[在此填写]` 占位符整行或前缀匹配时视为”未填写”，对应单元格留空
- Excel 被占用时输出明确的”关闭 Excel 后重试”提示
- 处理大型文件时注意内存
- **错误信息本地化（i18n 未来扩展）**：当前所有面向用户的中文错误文案（9999 上限、文件占用、无法解析 Excel 等）都集中由 `scripts/_common.py:localize_error` 渲染；未来若需支持英文或其他语言，只需把 `localize_error` 的中文文案迁出到语言资源文件并按 `LANG`/`LC_ALL` 选择，**调用方无需改动**。该函数是错误信息国际化的唯一扩展入口。
