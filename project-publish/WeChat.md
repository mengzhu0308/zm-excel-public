# WeChat

## v0.4.3
- 更新时间：2026-06-12
- Release URL：https://github.com/mengzhu0308/zm-excel-public/releases/tag/v0.4.3

zm-excel-public v0.4.3 已完成发布。本版本把过去一轮 A-1 / A-2 静态审查里识别的 P0/P1 风险点系统化落地为 33 个文件的代码与文档改动，覆盖全部 9 个 skill；重点是「明确护栏、补齐 i18n / 错误语义、收敛描述」三件事，不引入破坏性 CLI 行为。主要更新：dedup-merge 新增 NUL 字节护栏与公式注入防护、formalization 修复损坏/加密 xlsx 的友好错误、del-multi-rows 修复纯数字 --sheet 解析顺序与 -o / --format 一致性、全部 skill 的 SKILL.md 补充「不触发条件」与边界护栏段。Release URL：https://github.com/mengzhu0308/zm-excel-public/releases/tag/v0.4.3。
