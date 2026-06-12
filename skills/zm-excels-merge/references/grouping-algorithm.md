# 分组算法说明

本文档说明 `merge_excels.py` 在默认合并模式与 `--preview` 模式下使用的字段兼容性分组算法。
阅读对象是希望了解内部行为或在修改阈值/算法时保持兼容性的开发者。

## 目标

在同一后缀（`.xlsx` / `.xls` / `.xlsm` / `.csv`）组内、同一 sheet 名下，按"列名集合"的相似度把多个文件划分成若干组，每组单独写到一个输出 sheet。这样：

- 列名完全一致的文件不会因为多了一列就被迫拆到不同 sheet；
- 差异显著的文件不会被强行合并，导致大量空值列污染数据；
- 用户可以通过 `--similarity-threshold` 调节严格度。

## 关键定义

- **列名集合**：对每个文件的 sheet，把表头读成 `set[str]`。读取阶段已经统一为 `dtype=str`，避免 `1` 与 `"1"` 的细微差异影响集合相等性。
- **Jaccard 相似度**：两个集合 `A`、`B` 的 Jaccard = `|A ∩ B| / |A ∪ B|`，范围 `[0, 1]`。
  - 当两集合都为空时按 1.0 处理（"两个文件都没有列名"视为完全一致）。
  - 当且仅当 union 为 0（即两集合都为空）时返回 0.0，避免除零。
- **阈值（threshold）**：CLI 参数 `--similarity-threshold` 决定两个集合是否"兼容"。CLI 层约束 `0 ≤ v ≤ 1`（见 `_probability` argparse type）。

## 分组算法

入口：`group_files_by_columns(sheet_files_data, threshold=0.8)`。

输入按 sheet 分桶后，组内调用本函数。算法是**贪心单链扩展**：

1. 维护一个待分组队列 `remaining`，初始为全部 `(filepath, columns_set)`。
2. 循环：
   1. 从队首取一个作为本组的 `seed`，放入新组。
   2. 遍历 `remaining` 中每个未分组条目 `item`：
      - 如果 `item.columns_set` 与组内**所有成员**的相似度都 `>= threshold`，则把 `item` 加入当前组。
      - 否则把 `item` 放回下一轮的 `remaining`。
   3. 把当前组追加到 `groups`，回到步骤 1 直到 `remaining` 为空。
3. 返回 `[[(filepath, cols), ...], ...]`。

> **关键不变量**：组内任意两成员的相似度都 `>= threshold`（不只是 seed）。
> 这是"全对全"判定，不是"与 seed 兼容即可"——避免大组在扩展过程中被稀释。

### 示例

阈值 `0.8`，三个文件：

| 文件 | 列名集合 | 共同列 | 并集 | Jaccard |
| --- | --- | --- | --- | --- |
| a.xlsx | {姓名, 年龄, 城市} | — | — | — |
| b.xlsx | {姓名, 年龄, 城市, 备注} | 3 / 4 | 4 | 0.75 |
| c.xlsx | {姓名, 部门} | 1 / 4 | 4 | 0.25 |

分组过程：

1. seed = a，组 = [{a}]。
   - b 与 a 相似度 0.75 < 0.8 → 不入组
   - c 与 a 相似度 0.25 < 0.8 → 不入组
2. seed = b，组 = [{b}]。
   - c 与 b 相似度 0.25 < 0.8 → 不入组
3. seed = c，组 = [{c}]。

结果：`[{a}, {b}, {c}]` —— 三个文件各成一组，输出三个独立 sheet。

## 阈值选择

- **默认 0.8**：保守，要求差异不超过 20%，避免语义不同的字段被强行合并。
- **调低到 0.5 ~ 0.6**：放宽，适合"列名拼写差异、扩展列"为主的场景。
- **调到 1.0**：等价于"列名集合完全相等才合一组"。
- **调到 0**：所有文件都互相兼容，会被贪心算法全部并入第一组（等价于单组合并）。这是算法自然结果，不是独立开关；想明确表达"不分组的单组合并"语义时，推荐使用 `--no-auto-group`（按同名 sheet 分组）。

## 与 `--no-auto-group` 的关系

`--no-auto-group` 不调用 `group_files_by_columns`；改为按同名 sheet 强制合并，遇到列名不一致时按 `pd.concat` 的列并集补空值。
它与阈值调成 0 行为相似，但意图不同：前者是"放弃分组，按 sheet 名硬合"，后者是"分组阈值最宽松"。

## 边界情况

- **单文件**：直接返回 `[(file, cols)]`，不分组。
- **同名列名重复**：列名去重在 `merge_excels` 主流程的列对齐段处理，本函数不感知。
- **空 sheet**（无表头）：列集合为空，与其他空列集合按 Jaccard=1.0 兼容；与非空集合按"交集为 0 / 并集为非空"得 0.0，不兼容。
- **极大数据**：分组只依赖列集合（轻量），不需要读出全部数据，复杂度近似 O(n² · m)（n 文件数，m 平均列数）。

## 变更注意

- 改阈值默认（如 0.8 → 0.7）属于用户可见行为变化，必须同步更新 `SKILL.md` 字段兼容性分析段、`README.md` 设计理念与 FAQ。
- 改算法（如引入"全集覆盖度"作为第二约束）属于行为变更，应在 `CHANGELOG.md` 的 `Changed` 段记录，并提供迁移说明。
- 改 `_non_negative_int` / `_probability` argparse type 越界行为属于用户契约变更；放宽或收紧前请先评估对现有用户的影响。

## 关联代码

- `scripts/merge_excels.py:jaccard_similarity`
- `scripts/merge_excels.py:group_files_by_columns`
- `scripts/merge_excels.py:analyze_compatibility`
- `scripts/merge_excels.py:_probability`（CLI 层阈值范围校验）
