# excel-master 移除截图功能实施计划

> **For agentic workers:** 未授权使用子 agent；读取 `executing-plans.md` 后由当前 agent 串行执行。所有已有文件先集中备份并校验，删除只进入 Windows 回收站。

**Goal:** 完整移除真实 Excel 截图和敏感列打码分支，并证明 Excel 生成、美化和复杂区块功能不受影响。

**Architecture:** 删除独立可选分支，不修改三个核心入口的业务逻辑。活动说明、依赖、测试和技能卡同步收敛；历史审计保留。当前工作树有未提交改动，因此不建工作树、不提交，用集中备份和 SHA256 作为恢复棘轮。

**Tech Stack:** Markdown、Python 3、pytest/unittest、pandas、openpyxl、Windows 回收站。

## Global Constraints

- 只处理 `D:\BaiduSyncdisk\workbuddy skills\excel-master`。
- 不修改 `make_excel.py`、`format_excel_sections.py` 的业务逻辑或公开接口。
- 不卸载系统包，不启动或结束 Excel 进程。
- 不提交、不推送、不创建 PR，不使用 stash/reset/revert。
- 已有文件先备份到 `C:\Users\27651\BackUp\` 并校验 SHA256。
- 删除只使用 Windows 回收站。
- `SKILL.md` 保持无 BOM UTF-8 和 CRLF；其他中文文本保持 UTF-8 BOM 和 CRLF。

---

### Task 1: 建立删除契约和集中备份

**Files:**
- Read: `SKILL.md`
- Read: `README.md`
- Read: `tests\test_self_contained_tools.py`
- Backup: 全部待修改和待删除文件

**Interfaces:**
- Consumes: 需求中的删除清单和验收标准。
- Produces: 预期失败的删除契约、哈希一致的集中备份。

- [x] **Step 1: 运行 RED 检查**

```powershell
python -X utf8 -B -c "from pathlib import Path; r=Path(r'D:\BaiduSyncdisk\workbuddy skills\excel-master'); assert all(not (r/p).exists() for p in ['scripts/excel_camera_screenshot.py','scripts/blur-excel-column.py','references/excel-screenshot-camera.md'])"
```

Expected: 断言失败，因为待删除文件仍存在。

- [x] **Step 2: 确认删除和修改清单**

逐个解析绝对路径，确认目标均在 `excel-master` 内；缓存文件只按精确文件名纳入，不删除目录。

- [x] **Step 3: 建立集中备份**

按原相对路径复制所有待修改和待删除文件，逐文件比较 SHA256。Expected: 不一致数量为 0。

---

### Task 2: 清理活动说明和测试

**Files:**
- Modify: `SKILL.md`
- Modify: `README.md`
- Modify: `references\implementation-checklist.md`
- Modify: `tests\test_self_contained_tools.py`
- Modify: `test-prompts.json`
- Modify: `skill-card.html`

**Interfaces:**
- Consumes: 方案 B 的活动边界。
- Produces: 只描述生成、美化和复杂区块能力的当前文档与测试。

- [x] **Step 1: 清理技能入口和依赖**

从 `SKILL.md` 删除截图、打码、Excel COM、pywin32、Pillow、剪贴板相关入口、步骤、失败处理、命令、参考资源和黑名单；保留备份、公式、值、条件格式和复杂区块安全门。

- [x] **Step 2: 清理用户说明和检查表**

从 `README.md` 与实施检查表删除当前截图能力、可选依赖和命令；修正原地覆盖必须保持自动备份的表述。

- [x] **Step 3: 清理专用测试和提示**

删除截图与打码测试、导入和第三个测试提示；保留复杂区块的两个真实行为测试。

- [x] **Step 4: 更新技能卡**

把副标题改为“Excel 生成、美化与复杂区块 · 12 套色系主题”，不改变卡片布局。

---

### Task 3: 回收活动文件和缓存

**Files:**
- Recycle: `scripts\excel_camera_screenshot.py`
- Recycle: `scripts\blur-excel-column.py`
- Recycle: `references\excel-screenshot-camera.md`
- Recycle: 对应的精确 `__pycache__` 文件

**Interfaces:**
- Consumes: 已通过哈希校验的备份。
- Produces: 活动技能目录中不存在的截图分支。

- [x] **Step 1: 再次核对目标**

确认每个目标绝对路径仍在技能目录，且备份文件存在、哈希等于当前目标。

- [x] **Step 2: 进入 Windows 回收站**

使用 `Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile` 的 `SendToRecycleBin` 模式逐文件删除，不删除任何目录。

- [x] **Step 3: 运行 GREEN 检查**

重新运行 Task 1 Step 1。Expected: PASS。

---

### Task 4: 生成技能卡并回归保留功能

**Files:**
- Modify: `skill-card.png`
- Update: `results.tsv`
- Create: `docs\reviews\2026-07-26-excel-master-remove-screenshot-review.md`

**Interfaces:**
- Consumes: 更新后的技能卡 HTML 和三个保留入口。
- Produces: 当前技能卡、真实回归证据和复核报告。

- [x] **Step 1: 重新生成技能卡图片**

用本机 Edge 渲染 `skill-card.html` 为 1280×720 PNG，并目视确认副标题正确、没有截断。

- [x] **Step 2: 运行完整测试**

```powershell
python -X utf8 -m pytest -q
```

Expected: 12 passed，0 failed。

- [x] **Step 3: 运行三个真实样例**

执行 CSV 生成、现有工作簿美化和显式区块格式化；比较值、公式、条件格式和关键格式。Expected: 三个样例均通过。

- [x] **Step 4: 运行静态与安全检查**

运行 `quick_validate.py`、runtime 红灯、资源路径、自包含、编码、换行、活动引用和 `git diff --check`；比较验证前后 Excel 进程编号。Expected: 全部通过且 Excel 进程集合不变。

- [x] **Step 5: 记录与复核**

在 `results.tsv` 追加功能删除记录，写复核报告并逐项对照本计划。

## 自审

- 需求的八项验收标准均映射到 Task 1–4。
- 没有占位符、待定项、新依赖或新抽象。
- 修改和删除均由当前 agent 串行完成。
- 方案只删除独立分支，保留核心实现与历史审计。
