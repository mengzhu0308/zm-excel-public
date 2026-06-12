# 脚本拆分与设计说明

`scripts/del_multi_rows.py` 是一个单文件 CLI 入口；本文件记录脚本内部的关键设计取舍，供维护者参考，避免后人改写时回归问题。

## 文件职责

| 文件 | 职责 | 是否可被 import |
| ---- | ---- | --------------- |
| `scripts/del_multi_rows.py` | CLI 入口与匹配/写盘全流程（约 400 行） | 否（仅作为 CLI 入口；同仓库级别一致） |

## 关键设计点

### 为什么抽 `CSV_KEY` 常量

CSV 单表在 `read_file` 中以 `data_dict[CSV_KEY] = df` 形式塞进字典；下游 `delete_rows` / `derive_output_path` / `write_file` / 主循环都需要用同一哨兵键名判断"这是 CSV 单表"。

把字符串 `"__csv__"` 抽为模块常量 `CSV_KEY`，避免 5 处字面量散落导致改一处忘改另一处（如把哨兵改名为 `"__csv_table__"` 漏改 1 处会让 CSV 路径被误判为 sheet 名）。

### 为什么 `derive_output_path` 拒绝 `-o` 与 `-f` 同源

`derive_output_path` 是"自动命名"路径；`--output` 显式路径若与 `-f` 指向同一文件，会直接覆盖源文件，与"源文件只读"核心原则冲突。

实现上先用 `os.path.abspath` 解析两侧真实路径再比较，规避相对路径 / 软链 / `..` 的歧义。命中即抛 `ValueError`，由 `main` 报"省略 -o 让脚本自动命名"提示并退出。

错误消息文案抽为模块常量 `_ERR_OUTPUT_SAME_AS_INPUT`，与 `main` 顶层的同源防护共用，避免改一处忘改另一处。

### 为什么 NaN 还原走 `.where(notna(), np.nan)`

`df.astype(str)` 会把 NaN 变成字符串 `"nan"`，导致关键词 `"nan"`（或任何恰好等于 `astype(str)` 后字符串值的子串，如 `"a"` / `"an"` / `"n"`）误删含 NaN 的行。

解决思路是：先用 `df.astype(str)` 字符串化所有单元格，再用 `df.notna()` 掩码把 NaN 还原为 `np.nan`。这样 `str.contains(..., na=False)` 与 `==` 比较时 NaN 会被自动屏蔽，不会被误匹配。

这个"先转字符串再还原 NaN"的组合比"逐单元格 `pd.isna` 判断再字符串化"快约一个数量级，因为向量化操作只走一次 `astype` 一次 `where`，没有 Python 层 row-wise 循环。

## 输出命名约定

| 场景 | 输出文件名 |
| ---- | ---------- |
| 单 sheet xlsx + 不指定 `--sheet` | `<原文件名>_删除多行.<ext>` |
| 多 sheet xlsx（不指定 `--sheet`） | 每个 sheet 一个文件：`<原文件名>_删除多行_<sheet名>.<ext>` |
| 显式 `--sheet`（单或多 sheet 都生效） | `<原文件名>_删除多行_<sheet名>.<ext>` |
| CSV（无 sheet 概念） | `<原文件名>_删除多行.csv` |
| `--format xlsx`（CSV 输入） | `<原文件名>_删除多行.xlsx`，内部 sheet 名固定为 `Sheet1` |

含特殊字符（`< > : " / \ | ? *` 与 ASCII 控制字符 `\x00-\x1f`）的 sheet 名在拼接到文件名前会被 `_sanitize_sheet_name` 替换为下划线，避免生成无法创建的文件路径。

## 边界护栏

脚本在主循环开始前对用户输入做三类显式护栏，与"源文件只读"原则对称：

- **空 / 空白关键词过滤**：`-k ""` / `-k '   '` / `-k '　'` / `-k '\xa0'` 等会被 `if k and not k.isspace()` 过滤并报错退出，避免误删整表
- **同源防护**：`--output` 与 `-f` 指向同一文件时直接报错退出，不进入任何行处理
- **系统目录白名单**：`--output` 落在 `/etc` / `/usr` / `/var` / `/bin` / `/sbin` / `/boot` / `/sys` / `/proc` / `/dev` / `/root` 等系统敏感目录时直接报错退出，避免 `os.makedirs` 越权创建
