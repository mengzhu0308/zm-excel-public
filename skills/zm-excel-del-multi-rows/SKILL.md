---
name: zm-excel-del-multi-rows
description: >-
  按关键词删除 Excel/CSV 中匹配的多行数据，输出到新文件，源文件只读不修改。
  当用户提到删除 Excel 行、删除含关键词/状态/标记的行、清理某类记录、剔除特定行、按条件过滤掉部分行时，**使用此 skill**。
  适用于 .xlsx / .xlsm / .csv 输入；支持多关键词 OR、子串/精确匹配、大小写敏感开关、多 sheet 全处理或指定单个 sheet（名称或 0-based 整数索引）、先预览再执行的干跑模式。
  不触发：仅查询/筛选不删除（用 `zm-excel-query`）、创建公式、财务建模、生成图表、修改格式（用 `zm-excel-formalization`）。
license: MIT
metadata:
  skill_mode: hybrid
compatibility:
  runtime:
    - name: agent-skills
      call_command: "conda run -n agent-skills python \"$SKILL_DIR/scripts/del_multi_rows.py\" [args]"
---

# zm-excel-del-multi-rows

按关键词搜索并删除 Excel/CSV 中匹配的多行数据，输出到新文件。源文件只读，不会被修改。

## 核心原则

- **源文件只读**：删除在内存中进行，结果输出到新文件
- **同目录输出**：文件名自动附加 `_删除多行` 后缀
- **多 sheet 全处理**：Excel 多 sheet 时每个 sheet 输出独立文件
- **默认子串匹配、不区分大小写**：任意列搜索，可用 `--match-mode exact` / `--case-sensitive` 切换
- **关键词可叠加**：多个关键词用 `-k` 多次指定，匹配任一即删除

## 与 /goal 配合使用

详见 [README.md](README.md)「与 /goal 配合使用」段；本 skill 不在脚本层接受 `/goal` 相关参数。

## 支持的输入格式

| 格式 | 扩展名 | 多 sheet |
|------|--------|----------|
| Excel | .xlsx, .xlsm | 支持 |
| CSV | .csv | 无（单表） |

> 注：本 skill 不支持 `.xls` 输入；遇到 `.xls` 会直接报错并提示先用 Excel 另存为 `.xlsx`，与同级 skill `zm-excel-add-one-row` 口径一致。

## 使用方式

### 方式一：直接调用脚本

> `$SKILL_DIR` 是安装态约定变量，由 `project-install` 注入；在 `skills/<skill-name>/scripts/...` 源码态下应替换为对应 skill 根目录。

```bash
conda run -n agent-skills python "$SKILL_DIR/scripts/del_multi_rows.py"   -f data.xlsx   -k "作废"   -k "删除"
```

输出：`data_删除多行.xlsx`

**常用选项：**

| 选项 | 说明 |
|------|------|
| `--sheet Sheet2` | 指定 sheet（仅 Excel） |
| `--match-mode exact` | 精确匹配（默认子串匹配） |
| `--case-sensitive` | 区分大小写 |
| `--dry-run` | 预览，只打印将删除的行数 |
| `-o path` | 自定义输出路径 |
| `--format csv` | 指定输出格式 |

### 方式二：AI 推理执行（灵活删除）

> **何时用脚本 vs 何时用 AI 推理？**
>
> - **用脚本**（方式一）：关键词简单、列无限制、单一 OR 逻辑、`--dry-run` 先确认范围——这覆盖 90% 场景
> - **用 AI 推理**（方式二）：需要 AND 多条件、列级过滤、复杂正则、按多列组合判断、或输出格式需自定义（如保留特定 sheet 不动）
>
> AI 解析意图后调用脚本或生成 Python 代码执行：

```python
import pandas as pd

df = pd.read_excel('data.xlsx')
# 删除包含 "作废" 或 "删除" 的行
mask = df.apply(lambda row: row.astype(str).str.contains('作废|删除', case=False, na=False).any(), axis=1)
result = df[~mask].reset_index(drop=True)
result.to_excel('data_删除多行.xlsx', index=False)
```

## 参数说明

| 参数 | 简写 | 说明 |
|------|------|------|
| `--file` | `-f` | 输入文件路径（必填） |
| `--keyword` | `-k` | 搜索关键词，可多次指定（必填） |
| `--sheet` | | 指定处理的 sheet 名称或 0-based 整数索引（仅 Excel），不指定则处理所有 sheet |
| `--match-mode` | | `contains`（子串匹配，默认）或 `exact`（精确匹配） |
| `--case-sensitive` | | 区分大小写 |
| `--output` | `-o` | 自定义输出文件路径 |
| `--format` | | 输出格式：`xlsx` 或 `csv`（默认与输入格式相同） |
| `--dry-run` | | 预览模式，只打印将删除的行数，不生成文件 |
| `--header-row` | | 表头所在行（0-based，默认 0）；xlsx 前几行是注释/合并标题时使用 |

## 输出规范

- 默认输出到源文件同目录，文件名附加 `_删除多行` 后缀
- **CSV**：输出 `<原文件名>_删除多行.csv`，无 sheet 名后缀
- **单 sheet Excel（不指定 `--sheet`）**：输出 `<原文件名>_删除多行.xlsx`，**不追加** sheet 名后缀
- **多 sheet Excel（不指定 `--sheet`）** 或 **显式指定 `--sheet`**：每个 sheet 输出独立文件，文件名追加 `_删除多行_<sheet名>`，如 `data_删除多行_Sheet1.xlsx`；含特殊字符的 sheet 名会被替换为下划线
- **`--output` 显式路径防护**：单 sheet 时若 `-o` 与 `-f` 指向同一文件，脚本立即报错退出（不进入任何行处理、不打印任何删除统计），避免覆盖源文件；多 sheet + `--output` 同样报错退出，提示省略 `-o` 让脚本按 sheet 自动命名
- 目标文件已存在时自动追加序号（`_1`、`_2`...）

## 匹配行为

- **搜索范围**：所有列的字符串化单元格
- **多关键词**：`OR` 逻辑，匹配任一即删除
- **空值**：NaN/空单元格不参与匹配
- **`contains`**：子串匹配（默认）；**`exact`**：完全相等

## 注意事项

- 多关键词用 `-k` 多次指定，不用逗号分隔
- `--dry-run` 适合先确认范围再执行
- 大文件一次加载到内存，极端场景需分段处理
- 代码实现细节（pandas + openpyxl 读写、`utf-8-sig` 输出 CSV、`reset_index(drop=True)`、显式错误处理）见 [README.md](README.md)「设计理念」段

## 评测

评测样本见 `evals/evals.json`；系统化测试由 `zm-auto-test-skill` 在 `tests-plans/zm-excel-del-multi-rows-tests/` 分桶目录按 A-1/A-2/A-3 + B 轮队列执行，详见 [README.md](README.md)「评测与测试」段。
