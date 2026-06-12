# 运行时环境与 `${SKILL_DIR}` 来源

本文件沉淀 `zm-xlsx2csv` 的 3 套 runtime 执行约定与 `${SKILL_DIR}` 变量来源，便于脚本命令可复制粘贴。

## `${SKILL_DIR}` 是什么

`${SKILL_DIR}` 是 skill 运行时（Codex CLI、Claude Code 等）注入到调用上下文的 skill 根目录绝对路径。它**不是** bash / zsh / PowerShell 等通用 shell 的默认变量；仅在调用方明确把 skill 文档 / 脚本转发到 shell 执行时由 runtime 注入。

典型取值示例：

- 源码态：`/home/zm/zm-skills/zm-excel/skills/zm-xlsx2csv`
- 安装态：`<runtime-config-dir>/skills/zm-xlsx2csv`（如 `~/.codex/skills/zm-xlsx2csv` 或 `~/.claude/skills/zm-xlsx2csv`）

## 3 套 runtime 模板

`SKILL.md` frontmatter `compatibility.runtime` 给出 3 套入口，**优先使用 `agent-skills`**，如不可用降级到 `system-python` 或 `uv`：

| runtime | 调用命令 | 适用 |
| --- | --- | --- |
| `agent-skills` | `conda run -n agent-skills python "$SKILL_DIR/scripts/excel2csv.py" data.xlsx` | 项目推荐环境（已装 `pandas` / `openpyxl` / `xlrd<2.0`） |
| `system-python` | `python3 "$SKILL_DIR/scripts/excel2csv.py" data.xlsx` | 依赖已装到系统 Python（`pip install pandas openpyxl 'xlrd<2.0'`） |
| `uv` | `uv run --with pandas --with openpyxl --with 'xlrd<2.0' python "$SKILL_DIR/scripts/excel2csv.py" data.xlsx` | 用 `uv` 自动按 `--with` 拉依赖，无系统污染 |

### agent-skills

- 适用：项目推荐环境 `agent-skills`（已装 `pandas` / `openpyxl` / `xlrd<2.0`）
- 优势：依赖与 skill 测试环境一致，避免本地 Python 缺包
- 限制：需要先 `conda env create -f environment.yml` 或类似方式安装

### system-python

- 适用：用户系统 Python 已装好依赖（如 `pip install pandas openpyxl 'xlrd<2.0'`，或 system 包管理器已装）
- 优势：跨平台一致，无需 conda
- 限制：依赖需用户自行管理

### uv

- 适用：`uv` 工具链已装（`pip install uv` 或 `brew install uv`）
- 优势：自动按 `--with` 装依赖到临时 venv，不污染系统
- 限制：首次调用会下载依赖，需联网

## 纯 shell 环境（runtime 未注入 `${SKILL_DIR}`）

直接用 skill 根目录的绝对路径替换 `${SKILL_DIR}`。例如在仓库根目录运行：

```bash
conda run -n agent-skills python ./skills/zm-xlsx2csv/scripts/excel2csv.py data.xlsx
```

安装态下，把 `./skills/zm-xlsx2csv/scripts/excel2csv.py` 替换为安装路径，例如 `~/.codex/skills/zm-xlsx2csv/scripts/excel2csv.py`。

## Windows PowerShell

PowerShell 字符串内 `${SKILL_DIR}` 与 `"$SKILL_DIR"` 等价（都走 session 变量插值），与 SKILL.md 正文命令示例保持一致；先把 session 变量设好即可直接复用：

```powershell
$SKILL_DIR = "$HOME\.claude\skills\zm-xlsx2csv"
conda run -n agent-skills python "$SKILL_DIR\scripts\excel2csv.py" data.xlsx
```

注意：PowerShell 中的 `$env:SKILL_DIR`（process 环境变量）和 `$SKILL_DIR`（session 变量）是两个不同作用域。如果用 `$env:SKILL_DIR`，需把 SKILL.md 命令示例中的 `${SKILL_DIR}` 替换为 `$env:SKILL_DIR`。

## Windows cmd.exe

`${SKILL_DIR}` 在 cmd.exe 中**不是**合法变量，会被当作字面量处理。`%SKILL_DIR%` 才是 cmd.exe 的变量语法（注意 `%` 而非 `${}`），且需先 `set` 定义：

```cmd
set SKILL_DIR=%USERPROFILE%\.claude\skills\zm-xlsx2csv
python "%SKILL_DIR%\scripts\excel2csv.py" data.xlsx
```

或直接用绝对路径：

```cmd
python "%USERPROFILE%\.claude\skills\zm-xlsx2csv\scripts\excel2csv.py" data.xlsx
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
| `ModuleNotFoundError: No module named 'pandas'` | 当前 runtime 未装依赖 | 切到 `agent-skills`，或降级 `uv` 自动装，或 `pip install pandas openpyxl 'xlrd<2.0'` |
| `XLRDError: Excel xlsx file; not supported` | `.xls` 用了 `xlrd>=2.0` | `pip install "xlrd<2.0"` 后重试 |
| 输出 CSV 打开中文乱码 | 没用 utf-8-sig 打开 | 用 Excel 双击打开或 `iconv` 检查 BOM |
| `output directory X is outside source directory Y` | `-o` 写到了源文件父目录之外 | 确认是预期行为（soft warning，不阻断） |
| `%SKILL_DIR%` 报"系统找不到指定的路径"（cmd.exe） | 没先 `set SKILL_DIR=...` | 按 Windows cmd.exe 段落先 set 再用 |
