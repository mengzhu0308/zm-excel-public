# 实现规范

本文件沉淀 `zm-xlsx2csv` 脚本实现的硬性约束，供脚本开发者参考；调用者无需阅读此文件。

## 读取约束

- 使用 `pandas.read_excel(..., engine=<按扩展名选>, header=0, dtype=str)` 读取
- 按扩展名显式选 engine：
  - `.xlsx` / `.xlsm` → `openpyxl`
  - `.xls` → `xlrd<2.0`（新版 xlrd 已不再支持 .xls）
- `dtype=str` 强制全字符串，避免前导零 / 工号被推断为 `int`
- `header=0` 显式声明首行为列名

## 写入约束

- 输出 CSV 使用 `utf-8-sig` 编码（确保 Excel 打开中文不乱码）
- 不输出 DataFrame 索引（`index=False`）
- 写入采用"先全读、再全写"两阶段：读阶段任何失败不写任何 CSV
- 写入使用 `tempfile` + `os.replace` 原子重命名，避免写一半被杀留半截

## 命名规则

```text
if sheet_name is None and len(sheet_names) == 1:
    out_name = "{base_stem}.csv"
else:
    out_name = "{base_stem}_{safe_sheet_name}.csv"
```

sheet 名中的非法文件名字符（`\ / : * ? " < > |`）自动替换为下划线。

## 批量容错与退出码

- 批量场景下，单文件失败应跳过并继续
- 默认退出码：
  - `0`：有失败但有成功（部分失败）
  - `1`：全部失败（零成功）
- `--strict` 标志让"任何失败 → 退出码 1"
- 失败列表在 stderr 汇总

## 路径与输出

- 目录扫描默认仅顶层；`--recursive` 开启递归
- `-r` / `--recursive` **仅在输入是目录时生效**；单文件输入下静默忽略
- 单文件读取支持 `--timeout SECONDS` 超时（Unix 用 SIGALRM；非 Unix 平台无超时）
- `--timeout` 必须为正整数（0、负数、非整数会被 argparse 拒绝）
- 输出目录创建失败时显式 `SystemExit(1)`，错误信息含"输出目录创建失败"
- 默认跳过已存在 CSV；`--overwrite` 显式覆盖
- `--unique` 开启时输出文件名冲突自动加 `_1`/`_2` 后缀（最多到 `_9999`）

## 参数校验边界

| 参数 | 边界 | 行为 |
| --- | --- | --- |
| `--timeout` | ≤ 0 / 非整数 | argparse 报错退出（exit 2） |
| `--timeout` | Unix 平台 > 0 | 用 `SIGALRM` 触发真实超时 |
| `--timeout` | Windows / 任何非 Unix | 静默忽略（`HAS_SIGALRM=False`），等价"无超时" |
| `--recursive` | 输入是单文件 | 静默忽略；不影响现有行为 |
| `-o` / `--output` | 路径不可写 | `SystemExit(1)`，stderr 含"输出目录创建失败" |
| `--overwrite` + 同名 CSV | 默认 | 跳过并 warning（`--unique` 需配 `--overwrite` 才生成后缀） |
| `--unique` 单用 + 同名 CSV | 单跑 `--unique` | 静默跳过（与默认无 `--unique` 行为一致；不生成后缀） |
| `--unique` + `--overwrite` + 同名 CSV | 输出存在 | 自动加 `_1`/`_2` 后缀，最多 `_9999` |

## 退出码契约（单一真相来源）

`SKILL.md` / `README.md` / `agents/openai.yaml` / `evals/evals.json` 中关于退出码的描述都引用本节，**不**各自复述。

| 场景 | success | failed | skipped | 默认退出码 | `--strict` 退出码 |
| --- | --- | --- | --- | --- | --- |
| 全部成功 | >0 | 0 | 0 | 0 | 0 |
| 部分成功 + 部分失败 | >0 | >0 | 0 | 0 | 1 |
| 部分成功 + 部分"已存在跳过" | >0 | 0 | >0 | 0 | 0 |
| 全部"已存在跳过"（幂等重跑） | 0 | 0 | >0 | 0 | 0 |
| 全部失败 | 0 | >0 | 0 | 1 | 1 |
| 无文件 / 无 sheet 匹配 / 文件打不开 | 0 | 0\* | 0 | 1 | 1 |

**关键约定**：

- "已存在跳过"**不**计入业务失败；不计入 `failed_files`；不影响退出码
- 退出码判断**只**看两个条件：
  - 默认模式：`success_count == 0` 且 `failed_files` 非空 → 1；否则 0
  - `--strict` 模式：`failed_files` 非空 → 1；否则 0
- 跳过文件也走 stdout 明细，标记为 `SKIP <path>`，方便统一收集流水
- 兼容旧版"全部失败 → 1"语义；新版**新增**"幂等重跑 → 0"语义，与 `--strict` 显式失败语义保持正交
- \* "无文件" 场景下 `failed_files` 实际为 0——`scripts/excel2csv.py:434-436` 在 `if not files: sys.exit(1)` 走单独分支，不经过 `failed_files` 累加。本表"failed 列"在"无文件"行只用于表达"该场景下退出码为 1"
