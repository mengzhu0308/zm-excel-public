# 运行时环境与 `${SKILL_DIR}` 来源

本文件沉淀 `zm-xlsx2csv` 的 `agent-skills` runtime 执行约定与 `${SKILL_DIR}` 变量来源，便于脚本命令可复制粘贴。

## `${SKILL_DIR}` 是什么

`${SKILL_DIR}` 是 skill 运行时（Codex CLI、Claude Code 等）注入到调用上下文的 skill 根目录绝对路径。它**不是** bash / zsh / PowerShell 等通用 shell 的默认变量；仅在调用方明确把 skill 文档 / 脚本转发到 shell 执行时由 runtime 注入。

典型取值示例：

- 源码态：`/home/zm/zm-skills/zm-excel/skills/zm-xlsx2csv`
- 安装态：`<runtime-config-dir>/skills/zm-xlsx2csv`（如 `~/.codex/skills/zm-xlsx2csv` 或 `~/.claude/skills/zm-xlsx2csv`）

## runtime 模板

`SKILL.md` frontmatter `compatibility.runtime` 以 `agent-skills` 为单一推荐入口：

| runtime | 调用命令 |
| --- | --- |
| `agent-skills` | `conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" data.xlsx` |

### agent-skills

- 适用：项目推荐环境 `agent-skills`（已装 `pandas` / `openpyxl` / `xlrd<2.0`）
- 优势：依赖与 skill 测试环境一致，避免本地 Python 缺包
- 限制：需要先 `conda env create -f environment.yml` 或类似方式安装

## 纯 shell 环境（runtime 未注入 `${SKILL_DIR}`）

直接用 skill 根目录的绝对路径替换 `${SKILL_DIR}`。例如在仓库根目录运行：

```bash
conda run -n agent-skills python ./skills/zm-xlsx2csv/scripts/excel2csv.py data.xlsx
```

安装态下，把 `./skills/zm-xlsx2csv/scripts/excel2csv.py` 替换为安装路径，例如 `~/.codex/skills/zm-xlsx2csv/scripts/excel2csv.py`。

## Windows PowerShell

`${SKILL_DIR}` 在 PowerShell 中是合法变量名（PowerShell 变量以 `$` 开头），无需特殊转义。仍建议把 `${SKILL_DIR}` 替换为绝对路径或手动设置：

```powershell
$env:SKILL_DIR = "$HOME\.claude\skills\zm-xlsx2csv"
conda run -n agent-skills python "$env:SKILL_DIR\scripts\excel2csv.py" data.xlsx
```

## 平台兼容性

| 平台 | 默认 shell | `${SKILL_DIR}` 是否默认存在 | 复制粘贴 SKILL.md 命令会怎样 |
| --- | --- | --- | --- |
| macOS | zsh | 否 | 报 `python: can't open file '${SKILL_DIR}/scripts/excel2csv.py'` |
| Linux | bash | 否 | 同上 |
| Windows | PowerShell | 否 | 报路径不存在 |
| Windows | cmd.exe | 否 | `${SKILL_DIR}` 被当作字面量处理，不报错但不解析 |

**结论**：纯 shell 环境（无论 macOS / Linux / Windows）下，都需要先用本节"纯 shell 环境"方法把 `${SKILL_DIR}` 替换为绝对路径。runtime（Codex CLI / Claude Code 等）注入 `${SKILL_DIR}` 才让模板可直接复制。

## 故障排查

| 症状 | 可能原因 | 处理 |
| --- | --- | --- |
| `python: can't open file '${SKILL_DIR}/scripts/excel2csv.py'` | shell 未注入 `${SKILL_DIR}` | 用 skill 根绝对路径替换 |
| `ModuleNotFoundError: No module named 'pandas'` | `agent-skills` 环境未装依赖 | 补装 `pandas openpyxl xlrd<2.0` 到 `agent-skills` 环境 |
| `XLRDError: Excel xlsx file; not supported` | .xls 用了 `xlrd>=2.0` | `pip install "xlrd<2.0"` 后重试 |
| 输出 CSV 打开中文乱码 | 没用 utf-8-sig 打开 | 用 Excel 双击打开或 `iconv` 检查 BOM |
