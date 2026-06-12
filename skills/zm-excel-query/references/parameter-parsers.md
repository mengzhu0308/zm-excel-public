# 参数解析器边界

本文件汇总 `--where` 之外各参数解析器的边界与陷阱，与 [where-expression.md](where-expression.md) 互补。

## `--agg`（`parse_agg`）

格式：`<func>:<col>[,<func>:<col>...]`，例如 `sum:salary,count:name,mean:age`。

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| `sum:salary,count`（缺冒号） | 抛 `ValueError: agg 条目缺少 ':' 分隔符: 'count'` | 检查每个条目都含 `:` |
| `:salary` / `sum:`（空 func 或 col） | 抛 `ValueError: agg 条目 func 或 col 为空` | func 与 col 都必须非空 |
| `sum:salary,sum:salary`（重复 col） | 抛 `ValueError: agg 重复列: 'salary'（已映射到 'sum'，又指定 'sum'）`（A-1 P1-2 修复） | 避免重复条目 |
| 引用不存在的列 | 抛 `ValueError: --agg 引用了不存在的列: [...]；可用列: [...]` | 先 `--preview` 看列名 |

## `--sort`（`parse_sort`）

格式：`<col>[ asc|desc][,<col>[ asc|desc]...]`，例如 `salary desc,age asc`。

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| `salary desc extra`（多余 token） | 静默忽略 `extra`（与 pandas 一致） | 保持每段 1-2 token |
| 大小写 `DESC` | 接受，与 `desc` 等价 | 习惯大写也行 |
| 引用不存在的列 | 抛 `ValueError: --sort 引用了不存在的列: [...]` | 先 `--preview` 看列名 |

## `--header`（传给 `pd.read_excel` / `pd.read_csv`）

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| `0`（默认） | 第 1 行作表头 | 标准场景 |
| `N >= 行数` | 抛 pandas `EmptyDataError` | 用 `--preview` 确认表头位置 |
| 负数 | argparse 不拦截（`type=int`），pandas 抛 `IndexError` | 仅传 `0` 或正整数 |

## `--tag`（0.3.0 新增白名单）

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| 包含 `..` / `/` / `\` / `:` / `*` 等 | 抛 `ValueError: tag 仅允许字母、数字、下划线、连字符`（由 `_validate_tag` 拦截） | 仅用 `EastHigh_v2` 这类 ASCII 标识 |
| 包含空格 | 抛 ValueError | 用下划线或连字符代替 |
| 空字符串 | 抛 ValueError | 始终提供非空 tag |

## `--limit` / `--offset`

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| 负数 | 抛 `ValueError: --limit 不能为负数 / --offset 不能为负数`（A-1 强化；之前 pandas `head(-1)` 反向行为） | 仅传 `0` 或正整数 |
| `0` | 合法：`head(0)` 返回 0 行；`offset=0` 等于不偏移 | 极值场景 |

## `--preview`（`type=int`）

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| 负数（如 `--preview -1`） | 抛 `ValueError: preview 行数不能为负数: -1`（A-2 P1-1 修复；之前 pandas `head(-1)` 反向返回除最后一行外全部） | 仅传 `0` 或正整数 |
| `0`（如 `--preview 0`） | 合法：走 `preview_data(df, n=0)`，返回 0 行样本 + 字段元信息（A-2 P1-2 修复；之前 `if preview:` 把 `0` 当 False，跳过 preview 走 `_query_single` 写文件） | 想"只查元信息不取样" |
| 省略值 | argparse 默认 `10` | 标准场景 |
| 多 sheet 模式 + `--preview 0` | 每个 sheet 都返回 5 行样本（被多 sheet 分支固定为 5 行避免 JSON 膨胀） | 多 sheet 想看全量请用 `0` 之外的预览或省略 `--preview` 走 CLI |

## 交互模式 `save <路径>`

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| 含 `..` 路径段 | 拒绝并提示（A-1 P0-1 修复；防止越界写父目录） | 用绝对路径或相对当前目录的子路径 |
| 正常路径 | 写入用户指定路径 | — |
