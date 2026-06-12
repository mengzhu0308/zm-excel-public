# 多 sheet 命名规则与非法字符处理

本文件沉淀 `zm-xlsx2csv` 处理多 sheet 文件的命名规则与非法字符替换表，作为 `SKILL.md` 表格的细节补充。

## 单一真相来源（伪代码）

```text
if sheet_name is None and len(sheet_names) == 1:
    out_name = "{base_stem}.csv"
else:
    out_name = "{base_stem}_{safe_sheet_name}.csv"
```

任何命名修改必须**同步**更新本伪代码与 `SKILL.md` 输出命名规则表、`scripts/excel2csv.py:convert_single_file` 的 `single_sheet_no_spec` 判断，以及 `README.md` 输出文件表。

## 4 个边界矩阵

| 原文件 sheet 数 | 是否指定 sheet | 期望输出 |
| --- | --- | --- |
| 1 | 否 | `原文件名.csv` |
| 1 | 是（名或索引） | `原文件名_Sheet名.csv` |
| ≥ 2 | 否 | `原文件名_Sheet1名.csv`, `原文件名_Sheet2名.csv`, ... |
| ≥ 2 | 是 | `原文件名_指定Sheet名.csv` |

## sheet 名 → 文件名片段：非法字符替换

仅当用户显式指定 sheet，或原文件多 sheet 时，sheet 名会拼到输出文件名。脚本对以下 Windows / POSIX 通用非法字符做下划线替换：

| 字符 | 含义 |
| --- | --- |
| `\` | Windows 路径分隔 |
| `/` | POSIX 路径分隔 |
| `:` | Windows 盘符 |
| `*` | 通配符 |
| `?` | 通配符 |
| `"` | 引号 |
| `<` | 重定向 |
| `>` | 重定向 |
| `\|` | 管道 |

替换后两侧空白 strip，中间空白保留。例如 sheet 名 `Q1: 销售 / 渠道?` 会变成 `Q1_ 销售 _ 渠道_`（连续非法字符不会合并，每个字符独立替换为 `_`）。

**已知限制（中间空白保留）**：

- sheet 名中间的空白会原样保留在文件名中；如 `销售  渠道`（中间两个空格）会生成 `xxx_销售  渠道.csv`，文件名内含连续空格
- 当前**不**折叠连续空白为单下划线；如确需折叠，可在调用前用 Excel 重命名 sheet 后再跑
- 此边界由 `scripts/excel2csv.py:_sheet_name_for_filename` 的 `re.sub(r'[\\/:*?"<>|]', '_', ...)` 行为决定；A-3 P0-1 把 evals case 23 重新指给 chmod 000 边界后，本边界未单独列入 evals cases 编号；验证可手工调用 `convert_single_file("带有  中间空格的 sheet.xlsx")` 确认 CSV 文件名含连续空格

## sheet 名查找失败的回退

`-s` 参数既接受 sheet 名也接受数字索引：

1. 先按 sheet 名精确匹配（区分大小写）
2. 不匹配则按 `int(sheet_name)` 解析为索引
3. 索引必须满足 `0 <= idx < len(sheet_names)`
4. 以上均不满足 → stderr 输出 "找不到 sheet 'xxx'。可用 sheet: [...]"，**批量场景下该文件计入失败但继续**

## 多 sheet + 数字索引

`-s 0` 表示第 1 个 sheet；`-s 1` 表示第 2 个 sheet。注意 sheet 顺序与打开 Excel 时看到的标签页顺序一致，但**与用户重排前的物理顺序可能不同**。建议优先用 sheet 名（更稳）。

## 公式与动态计算

`pd.read_excel(..., engine="openpyxl")` **不会计算公式**——openpyxl 读取的是 cell 缓存值（Excel 上次保存时计算的结果）。这意味着：

- 公式 cell 在 CSV 中反映"Excel 上次保存时"的值，不是当前最新计算值
- 如果用户在 Excel 中改了某个公式的依赖项但没保存，CSV 不会反映这次修改
- 引用了"外部文件"或"远程数据"的公式，openpyxl 也不会触发重新计算

如需最新计算值，请先在 Excel 中：

1. 全选工作表（`Ctrl+A` 或 `Cmd+A`）
2. 强制重算所有公式（`Ctrl+Alt+F9` / macOS `Cmd+Alt+⌘=`）
3. 保存（`Ctrl+S` / `Cmd+S`）
4. 再跑本 skill

## 案例与对应 evals

| 边界 | evals case | 期望文件 |
| --- | --- | --- |
| 单 sheet + 指定 sheet（带后缀） | evals #4 | `single_sheet_Sheet1.csv` |
| 多 sheet + 不指定（多文件） | evals #2 | `multi_sheet_员工信息.csv`, `multi_sheet_销售数据.csv` |
| 多 sheet + 指定 | evals #4 的多 sheet 变体 | `multi_sheet_员工信息.csv` |
| 不存在 sheet 名 | evals #6 | 失败 + stderr 提示 |
