---
name: zm-excel-dedup-merge
description: >-
  将两个 CSV/XLSX 表格按关键列去重合并。触发：合并同名记录（期刊目录、产品清单、人员名册等）；
  不触发：垂直拼接多表（用垂直拼接类 skill）、单表查询（用查询类 skill）、格式调整（用格式整理类 skill）。
license: MIT
metadata:
  skill_mode: hybrid
compatibility:
  runtime:
    - name: agent-skills
      call_command: "conda run -n agent-skills python \"$SKILL_DIR/scripts/dedup_merge.py\" [args]"
---

# zm-excel-dedup-merge

将两个表格按关键列去重合并，同名记录合并为一行，保留两表全部列。支持 CSV/XLSX 混合格式。

> **命令模板**：本 skill 的真实可执行命令以本文档 frontmatter `compatibility.runtime.call_command` 为单一真相来源。正文示例只是该模板的展开形式，便于复制。

## 核心原则

- **关键列匹配，非行追加**：与垂直拼接类 skill（按行追加）不同，核心是"同一实体 → 合并为一行"
- **自动检测关键列**：两表有同名列时自动识别；无同名列时找最相似列名配对
- **精确匹配优先，模糊兜底**：默认先标准化精确匹配，未匹配项再做模糊替换
- **列名冲突自动加后缀**：两表非关键列同名时，文件2列自动加 `_2` 后缀
- **输入格式决定输出格式**：CSV+CSV→CSV，XLSX+XLSX→XLSX，混合→XLSX
- **源文件只读**：操作在内存中进行

## 匹配策略

**精确匹配**：大写化、去首尾空格、压缩连续空格、去末尾句点。**严格大小写敏感**（保护 ISBN/ISSN 前导零等数字 ID）。

**模糊匹配**（默认开启，`--no-fuzzy` 关闭）：对未覆盖记录做归一化后再匹配：

- `academic` 预设（默认）：先转大写，再做 `AND`→`&`、连字符变空格、去 `THE ` 前缀、`JOURNALS OF`→`JOURNAL OF`；原文保留大小写用于输出
- `cjk` 预设：仅 `AND`→`&`、连字符变空格，不做大写化，适合中文/产品名册
- `none` 预设：不做任何模糊归一（与 `--no-fuzzy` 等价；保留开关便于与 `--fuzzy-preset` 串用）

> 关键值原文始终保留在合并表中（用 `_key_left` / `_key_right` 推断后回填），归一化副本只用于匹配决策，不写回输出。

**关键列自动检测倾向**：高唯一性 + 语义关键词（id / 编号 / 名称 / name / title / key / journal / publisher / author / product / sku / issn / isbn / doi / category / subject / tag / member 等）；命中 url / 网址 / 地址 / address / comment / 备注 / desc / description 等负面词会被降权。两表无共有列时退化为 `difflib` 序列相似度配对，对短列名（`id` / `name`）可能给出非直觉结果，建议显式指定 `--key1` / `--key2`。

## 使用方式

> 本节示例与 `call_command` 同步；`$SKILL_DIR` 指向当前 skill 根目录。源码态可手动设为 `skills/zm-excel-dedup-merge`，安装态由运行时或安装路径确定。

### 脚本调用

```bash
# 安装态：环境变量 $SKILL_DIR 已被 project-install 注入到本机运行态；源码态可手动指定：
SKILL_DIR="${SKILL_DIR:-$PWD}"  # 安装态无需覆盖；源码态 cd 到 skills/zm-excel-dedup-merge 后执行
conda run -n agent-skills python "$SKILL_DIR/scripts/dedup_merge.py" \
  -1 tableA.csv -2 tableB.csv \
  [--key1 "列名"] [--key2 "列名"] \
  [--no-fuzzy] [--dry-run] [-v]
```

| 参数 | 说明 |
|------|------|
| `-1`, `--file1` | **必需** 第一个输入文件 |
| `-2`, `--file2` | **必需** 第二个输入文件 |
| `--key1` | 文件1关键列名（默认自动检测） |
| `--key2` | 文件2关键列名（默认与 `--key1` 相同或自动检测） |
| `--key-uniqueness-min` | 自动检测关键列时要求的最低唯一值比例（0-1，默认 `0.5`）；当短列名（`id` / `name`）占多数或重复值过多时下调到 `0.1`-`0.3` 可放宽阈值 |
| `-o`, `--output` | 输出路径（默认 `dedup_merged.<ext>`，位于当前工作目录） |
| `--force` | 允许覆盖已存在主表（默认拒绝覆盖；不影响 `-o` 与输入重名护栏） |
| `--no-fuzzy` | 关闭模糊匹配 |
| `--fuzzy-preset` | 模糊匹配预设（`academic` / `cjk` / `none`），默认 `academic`；详见下文 |
| `--sort-by-key` | 合并后按关键列排序（默认关闭，保留两表原行序） |
| `--dry-run` | 预览匹配统计，不写入主表 |
| `--match-log` | 显式指定匹配明细 CSV 路径（仅在 `dry-run` 下生效；缺省时不写） |
| `--version` | 打印脚本版本号（读取自 `VERSION.yaml.skill_info.version`）并退出 |
| `--check-conda-env-consistency` | 自检：核对 `SKILL.md` frontmatter `compatibility.runtime[*].name` 与 `agents/openai.yaml` `system_requirements.conda_env` 一致性；返回 0 表示一致，1 表示不一致；不必提供 `-1` / `-2` 即可单独运行 |
| `-v`, `--verbose` | 详细日志 |

## 前置依赖

- Python 3.8+
- `pandas`（读取 CSV/Excel）
- `openpyxl`（写入/读取 `.xlsx`）
- 真实运行命令以 `compatibility.runtime.call_command` 为准；默认走 `conda run -n agent-skills`；未装 `conda` 时按下方"conda 环境"段备选

安装示例：

```bash
# 方式 1：conda（与 call_command 一致）
conda create -n agent-skills python=3.9
conda run -n agent-skills pip install pandas openpyxl

# 方式 2：pip（仅当你已把脚本改用本地 python3 / 系统环境）
python3 -m pip install pandas openpyxl
```

## 模糊匹配预设

| 预设 | 规则 | 适用场景 |
|------|------|----------|
| `academic`（默认） | `AND`→`&`、连字符变空格、去 `THE ` 前缀、`JOURNALS OF`→`JOURNAL OF` | 英文学术期刊名册 |
| `cjk` | 仅 `AND`→`&`、连字符变空格 | 中文学术表、产品清单、人员名册 |
| `none` | 不做任何模糊归一 | 关键值已严格一致、要求 0 误匹配 |

## 输出统计字段

| 字段 | 含义 |
|------|------|
| `total` | 合并后总行数 |
| `both` | 两表都匹配上的行数（含模糊匹配） |
| `fuzzy` | 仅靠模糊匹配才匹配上的行数 |
| `left_only` | 仅文件1有的行数（关键列空值不计入此字段，详见下方脚注） |
| `right_only` | 仅文件2有的行数（关键列空值不计入此字段，详见下方脚注） |
| `null_key_1` | 文件1中关键列为空值的行数（不参与匹配，保留为 `left_only`） |
| `null_key_2` | 文件2中关键列为空值的行数（不参与匹配，保留为 `right_only`） |

> **脚注**：关键列空值（NaN / None / 空串 / 纯空白）不参与精确或模糊匹配，单独计入 `null_key_1` / `null_key_2`；它们在合并表中仍会作为 `left_only` / `right_only` 行出现，但不会和对方空值互相合并为一行。

## 注意事项

- 关键列空值（NaN、None、空串、纯空白）不参与匹配，会单独记入 `null_key_1` / `null_key_2`，同时在合并表中作为 `left_only` / `right_only` 行保留
- 两表关键列列名不同时，建议显式指定 `--key1` / `--key2`
- 大表（>10w 行）建议先用 `--dry-run` 预览；模糊匹配阶段会做两次 DataFrame 合并，性能随行数平方增长
- **fuzzy 实测参考**（`--fuzzy-preset academic`，可复现基准见 `evals/bench_fuzzy.py` 与 `assets/bench_fuzzy_1k.csv`）：1k×1k ≈ 0.39s、5k×5k ≈ 0.47s、10k×10k ≈ 0.53s；>10w 行建议先分桶或关闭模糊（`--no-fuzzy`）。复现命令：`python3 evals/bench_fuzzy.py --sizes 1k 5k 10k`；不同 Linux / Python / pandas 版本的数字会有偏差，应按"参考量级"对待。
- 两表非关键列同名时，文件1保持原名称，文件2追加 `_2` 后缀（已存在 `_1` / `_2` 时按 `_3`、`_4`... 顺延）
- 默认不排序；如需按关键列排序，传 `--sort-by-key`
- 两表无共有列时，列名自动检测会退化为 `difflib` 序列相似度配对，对短列名（如 `id` / `name`）可能给出非直觉结果，建议显式指定 `--key1` / `--key2`
- `.xls` / `.xlsm` 双表合并时，输出后缀沿用输入；`.xlsm` 的宏会丢失，请用 VBA / Excel 工具合并
- `--dry-run` 默认不写匹配明细；如需落盘匹配明细 CSV（含 `match_type`：exact / fuzzy / left_only / right_only / null_key_1 / null_key_2），显式传 `--match-log PATH`
- 默认输出到当前工作目录下的 `dedup_merged.<ext>`；用 `-o` 显式指定覆盖

## 已知限制

- **单文件去重未支持**：本 skill 聚焦两表去重合并；如需对单表按关键列去重（保留首条/末条），请用 `pandas.drop_duplicates` 或 `Excel` 内置去重
- **关键值字符串 `__NULL_KEY_SENTINEL__` 是脚本内部历史哨兵**：当前实现已不再用哨兵字符串做空值标记，但仍应避免在关键列里使用该字面量以免引起混淆
- **编码回退链是"试到能解为止"**：对于 `gb18030` 与 `shift_jis` 字节可能"成功"但产生 mojibake 的混合乱码样本，建议先显式用 `pandas.read_csv(..., encoding=...)` 验证
- **academic 模糊匹配严格只做归一化**：`AND`→`&`、连字符变空格、去 `THE ` 前缀、`JOURNALS OF`→`JOURNAL OF`；不做语义向量匹配；典型场景是"英文学术期刊名册"，对生僻简称、人名等不适用
