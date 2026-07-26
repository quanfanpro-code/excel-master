# excel-master 版本定位

## 当前唯一运行源

本 `excel-master` 目录是自包含的正式版本。运行时只读取本目录内的：

- `SKILL.md`
- `scripts/`
- `references/`
- `tests/`

不得从 Hermes、旧试验项目、其他 skill 或 GitHub 工作目录导入运行代码。

## 上游关系

GitHub 的 `benyichan/excel-master` 只作为上游参考。2026-07-26 的本地版采用内容级融合：

- 保留本地扩展关键词、截图打码和安全校验；
- 吸收上游12套主题、冻结行和公式对象处理思路；
- 修复双方已验证的问题；
- 不通过普通 Git 合并替换本地文件。

## 识别正式版本

正式目录必须同时满足：

1. 根目录存在合法的 `SKILL.md`；
2. `scripts/make_excel.py` 可独立导入；
3. `python -X utf8 -B -m unittest discover -s tests -v` 全部通过；
4. 不存在第二个嵌套的同名 `excel-master/SKILL.md`；
5. 文档不包含对外部项目目录的运行依赖。

历史审计文件只用于追溯，不代表当前功能状态；当前行为以代码、测试和 `implementation-checklist.md` 为准。
