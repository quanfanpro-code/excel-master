# excel-master 项目版本沿革

本文档帮助区分不同位置的 excel-master 版本，避免混淆。

## 版本脉络

| 位置 | 角色 | 状态 |
|------|------|------|
| `$HERMES_HOME/skills/data-science/excel-master/` | **正式版** — Hermes Agent skill，持续迭代中 | ✅ 当前版本 |
| `D:\projects\excel-master\` | GitHub 独立仓库（benyichan/excel-master），与 skill 内容对齐 | ✅ 同步版本 |
| `D:\Trae_project\excel_master\` | 早期试验原型，纯 openpyxl 三文件（core.py+guard.py+\_\_init\_\_.py） | ⛔ 存档（不反映当前能力） |
| `D:\Trae_project\excel_master_new\` | 重构尝试，仅 .git 目录 | ⛔ 已废弃 |

## 判断方法

遇到 `excel_master` 或 `excel-master` 目录时，按以下优先级判断当前版本：

1. **Hermes skill**（`$HERMES_HOME/skills/data-science/excel-master/`）— 始终优先
2. **D:\projects\excel-master** — 独立 GitHub 仓库，与 skill 内容一致
3. **D:\Trae_project\excel_master** — **早期试验版**，不要引用或依赖其中的代码

## 核心差异

早期试验版（Trae_project）与正式版的主要区别：

- **架构**：早期是单体模块（core.py + guard.py），正式版是 Hermes skill 结构（SKILL.md + scripts/make_excel.py + references/）
- **入口**：早期无 CLI，正式版提供 `make_excel.py` 命令行入口和 `beautify` 功能
- **列宽**：早期版可能依赖 xlwings 自动调整，正式版改用字符估算法（纯 openpyxl 一次保存）
- **验证**：正式版有完整的 `implementation-checklist.md` 格式验证清单
- **格式覆盖度**：正式版完整实现了 9 大摩根系原则
