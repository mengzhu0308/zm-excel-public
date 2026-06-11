# where 条件表达式语法参考

`--where` 参数支持"比较 + 逻辑 + 列表 + 字符串匹配"四类表达式，运算符左右必须有空格（`>=`/`<=` 例外）；多条件用 `and` / `or` 组合，括号可嵌套。

## 运算符优先级

从高到低：

1. 括号 `(...)`
2. 比较运算符 `==` `!=` `>` `<` `>=` `<=`
3. 集合 / 字符串运算符 `in` `contains` `startswith` `endswith`
4. 逻辑非（暂未支持显式 `not`，可用括号和反向比较替代）
5. 逻辑与 `and`
6. 逻辑或 `or`

> 同级从左到右结合。

## 完整运算符表

| 运算符 | 含义 | 适用类型 | 示例 |
|--------|------|----------|------|
| `==` | 等于 | 任意 | `status == 'active'` |
| `!=` | 不等于 | 任意 | `status != 'deleted'` |
| `>` `<` `>=` `<=` | 比较 | 数值 / 日期 | `age > 18` |
| `and` `or` | 逻辑组合 | 布尔 mask | `age > 18 and dept == 'Eng'` |
| `in` | 包含于列表 | 任意 | `dept in ['Sales', 'HR']` |
| `contains` | 字符串包含 | 文本 | `name contains '张'` |
| `startswith` | 字符串前缀 | 文本 | `email startswith 'admin'` |
| `endswith` | 字符串后缀 | 文本 | `email endswith '.com'` |

## 边界与陷阱

### 1. `in` 列表必须能被 `ast.literal_eval` 解析

`in` 右操作数会被 `_parse_value` 用 `ast.literal_eval` 解析，因此：

- ✅ 数字：`age in [25, 30, 40]`
- ✅ 字符串：`dept in ['Sales', 'HR']`
- ✅ 混合：`val in [1, 'a', 2.5]`
- ❌ 不加引号的标识符：`dept in [Sales, HR]` 会触发 `ast.literal_eval` 失败

### 2. 字符串内可包含比较运算符

`contains` / `startswith` / `endswith` 的搜索值可以包含 `>=`/`<=`/`==` 等符号，不会被错切。例：

```bash
--where "name contains '>=10'"   # 搜索包含字符串 ">=10" 的行
```

### 3. 字符串值用单引号或双引号都行

`--where` 表达式本身用 shell 单引号包围时，列内字符串用双引号更方便：

```bash
--where 'name contains "张" and age > 18'
```

或在 shell 双引号下用单引号：

```bash
--where "name contains '张' and age > 18"
```

### 4. 3+ 条件的 and / or 必须用括号分组以保证可读性

`A and B or C and D` 会按 `and` 优先于 `or` 的规则被解析为 `(A and B) or (C and D)`，但强烈建议显式加括号以避免歧义。

### 5. 与 `pandas.DataFrame.query()` 的差异

本解析器是独立实现的精简版本，与 `pandas.DataFrame.query()` 在以下方面有差异：

- 不支持 `@` 引用 Python 变量；
- 不支持 `not` / `~` 显式取反（用括号 + 反向比较替代）；
- 多词运算符 `in` / `contains` / `startswith` / `endswith` 是本 skill 扩展，不在 pandas.query 中；
- 表达式会被 `ast.literal_eval` 解析右值，避免了 pandas.query 对引号、空格的诸多限制。

## 失败模式与提示

- 列名不存在：抛出 `KeyError: <列名>`，在 CLI 下报为 `读取/查询错误`；
- 类型不匹配（如对文本列做 `> 100`）：返回空 mask，结果可能为空；
- 表达式无法解析：抛 `无法解析的条件表达式: <expr>`。

## 其他解析器边界

`--agg`、`--sort`、`--header` 等参数也有独立解析器；下表汇总它们的边界与陷阱。

### `--agg`（`parse_agg`）

格式：`<func>:<col>[,<func>:<col>...]`，例如 `sum:salary,count:name,mean:age`。

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| `sum:salary,count`（缺冒号） | 抛 `ValueError: agg 条目缺少 ':' 分隔符: 'count'` | 检查每个条目都含 `:` |
| `:salary` / `sum:`（空 func 或 col） | 抛 `ValueError: agg 条目 func 或 col 为空` | func 与 col 都必须非空 |
| `sum:salary,sum:salary`（重复） | 后者覆盖前者，不报错 | 避免重复条目 |
| 引用不存在的列 | 抛 `ValueError: --agg 引用了不存在的列: [...]；可用列: [...]` | 先 `--preview` 看列名 |

### `--sort`（`parse_sort`）

格式：`<col>[ asc|desc][,<col>[ asc|desc]...]`，例如 `salary desc,age asc`。

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| `salary desc extra`（多余 token） | 静默忽略 `extra`（与 pandas 一致） | 保持每段 1-2 token |
| 大小写 `DESC` | 接受，与 `desc` 等价 | 习惯大写也行 |
| 引用不存在的列 | 抛 `ValueError: --sort 引用了不存在的列: [...]` | 先 `--preview` 看列名 |

### `--header`（传给 `pd.read_excel` / `pd.read_csv`）

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| `0`（默认） | 第 1 行作表头 | 标准场景 |
| `N >= 行数` | 抛 pandas `EmptyDataError` | 用 `--preview` 确认表头位置 |
| 负数 | argparse 不拦截（`type=int`），pandas 抛 `IndexError` | 仅传 `0` 或正整数 |

### `--tag`（A-3 新增白名单）

| 边界 | 行为 | 建议写法 |
|------|------|----------|
| 包含 `..` / `/` / `\` / `:` / `*` 等 | 抛 `ValueError: tag 仅允许字母、数字、下划线、连字符`（由 `_validate_tag` 拦截） | 仅用 `EastHigh_v2` 这类 ASCII 标识 |
| 包含空格 | 抛 ValueError | 用下划线或连字符代替 |
| 空字符串 | 抛 ValueError | 始终提供非空 tag |

### `_parse_value` 静默回退

`where` 右值由 `_parse_value` 解析：`ast.literal_eval` 成功则用 Python 字面量；失败则回退为字符串。

- ✅ `age == 18` → 18（int），与数值列比较正常
- ✅ `name == 'Alice'` → `'Alice'`（str），与文本列比较正常
- ⚠️ `name == Alice`（无引号）→ 回退为字符串 `'Alice'`，文本列比较可能正常也可能异常
- ⚠️ `age == abc`（无引号）→ 回退为字符串 `'abc'`，与数值列比较必然失败
- ⚠️ `col == None` → `None`，与 NaN 行的比较行为依赖 pandas 版本

建议：所有字符串字面量都加引号，避免静默回退掩盖输入错误。
