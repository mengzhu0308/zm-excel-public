# 脚本拆分与设计说明

简要记录 `scripts/` 下三个文件的职责划分与设计取舍，供维护者参考。

## 文件职责

| 文件 | 职责 | 是否可被 import |
| ---- | ---- | --------------- |
| `scripts/_common.py` | 共用 helper：表头检测、示例行查找、工作表选择、样式复制、`load_workbook` 策略封装、扩展名校验 | 是（被另两个脚本 import） |
| `scripts/extract_template.py` | 阶段一：从 Excel 生成 Markdown 填写模板 | 否（仅作为 CLI 入口） |
| `scripts/write_back.py` | 阶段三/四：解析已填模板，预览或回填到 Excel | 否（仅作为 CLI 入口） |

## 关键设计点

### 为什么抽 `_common.py`

`extract_template.py` 和 `write_back.py` 都需要：

- `detect_header_row(ws)`：在前 3 行中挑非空单元格最多的那一行
- `find_sample_row(ws, header_cols)`：从后往前找"≥ 一半表头列有值"的最后一行；`header_cols` 由调用方先从 `read_headers` 计算并传入，避免在 `find_sample_row` 内部重复 IO + `normalize_header`
- `select_worksheet(wb, sheet_name)`：按名字或回退到 active sheet
- `copy_styles(src, dst)`：复制单元格样式
- `read_headers(ws, header_row_idx)` / `normalize_header(value)`：把表头规范成字符串

如果两边各写一份，后续一处修改会忘记同步另一处，导致“提取模板的表头”和“回填的表头”不一致的隐蔽 bug。统一抽到 `_common.py` 后，行为可单点维护。

### 为什么 `load_workbook` 要显式 `data_only` 策略

- `extract_template` 只需读示例行的“计算后的值”，所以 `data_only=True`
- `write_back` 需要保留原工作簿的公式与样式，所以 `data_only=False`

`load_workbook` helper 把这个差异显式化，避免两个脚本用同一份配置互相干扰。

### 为什么 `write_back` 改完共享同一个 `Workbook`

旧实现会先在 `_build_row_data` 里 `load_workbook` 一次（丢弃），再在主函数里 `load_workbook` 一次（用于写入）。两次读盘：

- 浪费 I/O
- 理论上同一份 `.xlsx` 在两次读盘之间被外部修改会拿到不同快照（虽然实际少见）

新实现把 `wb` / `ws` 沿用到预览、追加、复制样式和保存的全过程，单一 `Workbook` 实例。

### 为什么 `extract_template` 默认拒绝覆盖

`SKILL.md` 明确说“恢复点：续跑时优先读取已生成的 Markdown 模板，不重新生成覆盖用户填写内容”。但旧版 `extract_template` 会无条件 `write_text` 覆盖。修复后：

- 默认 `force=False`，若目标模板已存在则抛 `FileExistsError` 并退出码 2
- 加 `--force` 才允许显式覆盖
- 错误信息明确告诉用户加 `--force`

### 为什么 `_convert_value` 要与示例行类型挂钩

旧版对所有纯数字字符串统一尝试 `int` → `float`，会导致：

- 身份证 18 位 → `int` 成功 → Excel 15 位精度限制下变成 `110101199001011000`
- 工号 `007` → `int` 成功 → 写入 Excel 后变成 `7`，前导零丢失
- 订单号 `2024011500000123`（> 10 位）→ 同身份证问题

修复策略（保守）：

- 空值 → `None`
- 超长纯数字（> 10 位）→ 字符串
- 前导零的纯数字 → 字符串
- 纯数字 + 示例行是 `int` → 才推断为 `int`
- 纯数字 + 示例行不是 `int` → 字符串
- 含非数字字符 + 示例行是 `float` → 才推断为 `float`
- 其他 → 字符串

最小爆炸面：典型数字推断路径（示例行是数字，新值也是数字）保持原样。

### 数字推断的非对称分支

`_convert_value` 的两条"示例行是数字 + 新值是纯数字"路径都走推断：

- 示例行 `int` + 新值纯数字 → `int`（保留前缀零 / 超长护栏）
- 示例行 `float` + 新值纯数字 → `float`（同上）
- 其他情况 → 字符串

这意味着"价格"列示例行是 `1.5`、用户在 md 中填 `2` 时，新行 `价格` 列存的是 `2.0` 而非字符串 `"2"`，与典型用户的预期一致。`SKILL.md` 注意事项与 `README.md` Q&A 已同步重写。

### `find_sample_row` 与 `read_headers` 共享具名列概念

`find_sample_row` 的"具名列"集合与 `read_headers` 共享归一化结果（`normalize_header`），因此：

- whitespace-only 表头（如 `" "`）会被当作空列，不计入阈值分母；
- 边界场景（表头中段有空列 / 全 whitespace 列）下 sample row 判定更稳定。

### 扩展名校验

`_common.validate_excel_extension()` 集中处理"非 `.xlsx` / `.xlsm` 一律拒绝"的护栏；`extract_template.py` 与 `write_back.py` 都在 `main()` 入口调用，遇到 `.xls` 等不支持的扩展名会抛 `ValueError` 并被 `main()` 的通用 `except` 兜住（退出码 1），不让 openpyxl 抛非友好的 `zipfile.BadZipFile` 错误。`SUPPORTED_EXCEL_EXTENSIONS` 元组是唯一真相来源，新增支持扩展名时改这里即可。

### 模板两级标题的语义分工

`extract_template.py` 输出的 Markdown 模板同时使用 `##` 与 `#` 两级标题，两者职责不同：

- `## Excel 数据录入模板`：模板级标题，只在文件首部写一次，作为"这是回填模板"的视觉锚点。`write_back.parse_markdown` 的 `_HEADING_RE` 只匹配一级标题，不会把它当成字段名去匹配。
- `# 字段名`：每个 Excel 表头对应一个一级标题，是回填脚本字段解析的**唯一**依据；写错成 `##` 会被脚本忽略，导致整列无法匹配。

两级分工的好处：

- 模板首部可以加元信息块（`> 来源文件` / `> 工作表`）、占位符说明、视觉锚点，但不会污染字段解析
- 用户手动写模板时，知道"改 `#` 影响字段，改 `##` 不影响字段"

新增字段或修改元信息时，`extract_template.py` 的 `lines.append(...)` 系列调用是单一真相来源（见 `extract_template.py:101-121`），不要绕过它直接拼字符串。
