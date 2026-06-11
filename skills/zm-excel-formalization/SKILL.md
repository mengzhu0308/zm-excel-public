---
name: zm-excel-formalization
description: >-
  将现有 Excel 文件按正式/商务文档样式格式化：垂直居中、水平两端对齐、自动
  换行、中文宋体、英文 Times New Roman、自动调整最优列宽。支持单文件与批量
  处理，原地覆盖或输出到新路径。

  当用户需要格式化、美化、标准化 Excel 表格时，**务必优先使用此 skill**。触发
  场景包括：格式化 Excel 文件；应用正式文档样式；统一单元格外观；表格美化；
  调整 .xlsx/.xlsm 中的字体与对齐；批量格式化；调整列宽/自动列宽/最优
  列宽；或用户提到"格式化Excel""表格格式化""Excel排版""规范Excel样式"
  "统一格式""调整字体""设置对齐""美化表格""正式样式""调整列宽""自动列宽"
  "最优列宽"等。

  **不要**触发：涉及数据查询/筛选、创建图表、写入公式/计算、数据分析、或格式
  转换时——使用对应的 query 或 conversion skill。
license: MIT
metadata:
  skill_mode: hybrid
compatibility:
  runtime:
    - name: agent-skills
      call_command: "conda run -n agent-skills python \"$SKILL_DIR/scripts/format_excel.py\" [args]"
---

# zm-excel-formalization

将现有 Excel 文件按规范格式化为正式文档样式。

## 格式化规则

| 属性 | 规则 |
|------|------|
| **垂直对齐** | 居中 (`vertical="center"`) |
| **水平对齐** | 单行居中 (`horizontal="center"`)，多行两端均匀分布 (`horizontal="distributed"`) |
| **自动换行** | 开启 (`wrap_text=True`) |
| **CJK 字体** | 宋体（含中文、日文假名、韩文等 CJK 字符） |
| **西文字体** | Times New Roman（拉丁字母、数字、符号等） |
| **字号** | 保留原字号；原 `None` 时回退到 11；原 `0` 保持为 0 |
| **列宽** | 自动按内容调整（CJK 字符宽度计为 2，西文字符计为 1；上限 50，下限 8，边距 +2） |

## 与 /goal 配合使用

`/goal` 是会话层的任务跟踪能力（Codex CLI、Claude Code 等平台均提供），不是本 skill 的脚本参数。单文件格式化通常直接执行；当任务包含批量格式化、原地覆盖前确认、输出目录约束或跨会话继续时，建议先开启 `/goal`，记录输入范围、覆盖策略、列宽选项和完成标准。

## 使用方式

### 方式一：脚本直接调用（确定性操作）

```bash
conda run -n agent-skills python "$SKILL_DIR/scripts/format_excel.py" data.xlsx
```

默认同目录保存副本：`data.xlsx` → `data_副本.xlsx`；副本已存在则自动递增编号。

### 方式二：AI 推理执行（灵活场景）

> 完整参数与默认行为请参见 [脚本参数](#脚本参数) 与 [输出冲突行为](#输出冲突行为)；本节只演示核心逻辑。

```python
import openpyxl
from openpyxl.styles import Alignment, Font

# ⚠️ AI 推理路径请直接 `from format_excel import is_cjk_char`，
# 不要复制下面 7 区间版本——以下仅为教学片段，扩展 B / Hangul jamo 等会漏判
def is_cjk_char(ch):
    cp = ord(ch)
    return (0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or
            0x3000 <= cp <= 0x303F or 0xFF00 <= cp <= 0xFFEF or
            0x3040 <= cp <= 0x309F or 0x30A0 <= cp <= 0x30FF or
            0xAC00 <= cp <= 0xD7AF)

def get_font_name(text):
    return "宋体" if text and any(is_cjk_char(c) for c in text) else "Times New Roman"

wb = openpyxl.load_workbook('data.xlsx')
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            text = str(cell.value) if cell.value is not None else ""
            horizontal = "distributed" if "\n" in text else "center"
            # 所有单元格均设对齐，防止空行/空单元格被裁剪
            cell.alignment = Alignment(vertical="center", horizontal=horizontal, wrap_text=True)
            if cell.value is not None:
                # 注意：cell.font.size == 0 是合法值（表示继承默认），
                # 用 `or 11` 会把 0 误判并改成 11，因此用 `is not None` 区分
                current_size = cell.font.size
                new_size = current_size if current_size is not None else 11
                cell.font = Font(name=get_font_name(text), size=new_size,
                                 bold=cell.font.bold, italic=cell.font.italic,
                                 underline=cell.font.underline, strike=cell.font.strike,
                                 color=cell.font.color)
wb.save('data_副本.xlsx')  # 与脚本默认副本模式一致；冲突时手动加编号
```

## 脚本参数

| 参数 | 说明 |
|------|------|
| `input` | 输入文件、目录或通配符，可指定多个 |
| `--in-place` | 直接覆盖原文件（先写临时文件再原子替换，最大化降低中断风险） |
| `--output, -o` | 输出目录或文件路径，与 `--in-place` 互斥 |
| `--no-adjust-width` | 禁用自动调整列宽（默认启用） |
| `--verbose, -v` | 显示详细处理日志 |
| `--copy-suffix` | 默认副本模式下的文件名后缀，默认 `_副本` |
| `--max-file-size` | 单文件最大字节数（默认 200MB），超出则跳过并提示 |
| `--max-sheets` | 工作簿最大工作表数（默认 50），超出则跳过并提示 |
| `--dry-run` | 只打印将处理的文件与输出路径，不实际写入 |

## 输出冲突行为

- 单文件 + `-o <file>`：直接写入该文件路径
- 多文件 + `-o <file>`：报错并退出（避免静默覆盖）
- 多文件 + `-o <dir>`：先按 `f.name` 落到输出目录；同名冲突时把父目录名拼到 stem 后消歧（`report__dir1.xlsx` / `report__dir2.xlsx`）
- `--in-place` + 多文件：每个文件原地写临时文件再原子替换；替换失败保留原文件
- 默认副本模式：使用 `generate_copy_path` 自动生成 `原名<copy_suffix>.<ext>`，冲突时递增 `<copy_suffix>1..N`

## 注意事项

- 仅处理 `.xlsx`、`.xlsm` 格式文件（大小写不敏感；openpyxl 不支持 Excel 97-2003 的 `.xls` 二进制格式）
- 目录遍历仅处理当前层，不递归子目录；如需递归处理，请显式指定每个子目录
- 目录遍历默认跳过以 `.` 开头的隐藏文件
- 所有单元格均设置对齐样式（防止空行/空单元格被裁剪导致行数减少）；仅含值的单元格才设置字体
- 字号为 `None` 时回退到 11；字号为 `0`（合法值，表示继承默认）保持原值
- 批量处理时若单个文件失败，继续处理其余文件（仅单文件时失败则退出码非零）
- 所有写入路径（`--in-place`、默认副本模式、`-o` 输出）均采用临时文件+原子替换：先写 `.{name}.{pid}.{idx}.tmp` 再 `os.replace`，避免极端中断（磁盘满 / 编码异常 / 进程被杀）留下半成品
- 列宽自动调整默认启用，按内容显示宽度计算：CJK 字符计为 2 单位，西文字符计为 1 单位；处理显式换行时取最长行的宽度
- 依赖 `openpyxl >= 3.1, < 4`，运行 Python `>= 3.9` 的环境
