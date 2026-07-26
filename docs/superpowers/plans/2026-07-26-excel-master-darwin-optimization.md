# excel-master Darwin 优化实施计划

> **For agentic workers:** 用户已授权独立评审 agent；每个 agent 只能写自己的验证目录。主 agent 串行修改 `SKILL.md`，禁止多个 agent 写同一文件。

**Goal:** 在不改变核心 Excel 功能、公开接口和依赖的前提下，提高 `excel-master` 的工作流、失败处理、检查点和风险黑名单质量；真实任务若暴露自包含辅助脚本的共同根因，追加最小修复，并用独立盲评证明改进。

**Architecture:** `SKILL.md` 是主要可编辑运行资产。真实测试若稳定复现辅助脚本缺陷，只修改直接相关脚本并增加一个最小回归检查。测试提示、评分日志和流程文档记录验证证据。当前工作树有未提交改动，因此用集中备份和 SHA256 代替 Git 提交棘轮。

**Tech Stack:** Markdown、Python 3、unittest、openpyxl、pandas、Microsoft Excel COM、Darwin 九维 rubric。

## Global Constraints

- 说明结构只修改 `D:\BaiduSyncdisk\workbuddy skills\excel-master\SKILL.md`；辅助脚本仅允许修复真实测试稳定复现的共同根因。
- 不修改主题、核心格式引擎、公开接口或依赖。
- 不提交、不推送、不创建 PR，不使用 stash/reset/revert。
- `SKILL.md` 保持无 BOM UTF-8 和 CRLF；其他新增中文文本使用 UTF-8 BOM 和 CRLF。
- 新版大小不得超过旧版 150%，并优先缩短。
- 分数必须严格提高，且至少 2 个独立盲评 judge 一致偏好新版，否则恢复备份。

---

### Task 1: 建立基线

**Files:**
- Read: `SKILL.md`
- Read: `README.md`
- Create: `test-prompts.json`
- Create: `results.tsv`

**Interfaces:**
- Consumes: Darwin 九维 rubric、3 个测试提示和固定输入样例。
- Produces: 基线九维评分、runtime 扫描结果、无技能/带技能执行结果。

- [ ] **Step 1: 运行静态基线扫描**

检查 frontmatter、编号工作流、显性检查点、三段式失败表、软化措辞、黑名单、资源路径和 runtime 红灯。

- [ ] **Step 2: 执行无技能基线**

独立 agent 不得读取目标技能，输出到集中备份目录的 `baseline-run`。

- [ ] **Step 3: 执行带技能基线**

独立 agent 必须读取当前技能，输出到 `with-skill-run`。

- [ ] **Step 4: 验证真实输出**

检查 xlsx 可打开、订单号文本、金额/百分比格式、公式和值、条件格式、复杂区块、截图和打码。

- [ ] **Step 5: 记录基线**

把九维分数、总分、`runtime_warn=0` 和 `eval_mode=full_test` 写入 `results.tsv`。

---

### Task 2: Round 1 优化工作流与失败处理簇

**Files:**
- Modify: `SKILL.md`

**Interfaces:**
- Consumes: Task 1 最低分维度和现有脚本接口。
- Produces: 线性执行工作流、显性检查点、三段式失败处理。

- [ ] **Step 1: 运行失败检查**

```powershell
python -X utf8 -B -c "from pathlib import Path; t=Path('SKILL.md').read_text(encoding='utf-8'); assert '## 执行工作流' in t and '🔴 CHECKPOINT' in t and '仍失败时' in t"
```

Expected: 断言失败。

- [ ] **Step 2: 最小实现**

重组入口、依赖和执行说明，增加六步线性流程、两个显性检查点及“触发条件/第一处理/仍失败时”表。复用现有命令和规则，不增加功能。

- [ ] **Step 3: 运行定向检查**

重新运行 Step 1；Expected: PASS。

- [ ] **Step 4: 运行完整验证**

运行技能元数据校验、16 项测试、runtime 扫描和资源路径检查。

- [ ] **Step 5: 重新评分**

只把本轮归因维度记为“失败模式编码”；同时观察工作流和检查点的相关提升。

---

### Task 3: Round 2 优化反例黑名单

**Files:**
- Modify: `SKILL.md`

**Interfaces:**
- Consumes: Round 1 最优版本。
- Produces: 独立红线与反例黑名单，不增加重复说明。

- [ ] **Step 1: 运行失败检查**

```powershell
python -X utf8 -B -c "from pathlib import Path; t=Path('SKILL.md').read_text(encoding='utf-8'); assert '## 红线与反例黑名单' in t and '禁止结束全部 Excel 进程' in t"
```

Expected: 断言失败。

- [ ] **Step 2: 最小实现**

把现有 Pitfalls 中的安全规则压缩为黑名单表，至少覆盖猜测复杂布局、改变值或公式、静默重写条件格式、关闭全部 Excel、备份失败仍继续、未验证即交付。

- [ ] **Step 3: 运行定向检查**

重新运行 Step 1；Expected: PASS。

- [ ] **Step 4: 运行完整验证并重新评分**

若总分没有严格提高，恢复 Round 1；若提高则保留。

---

### Task 4: 修复真实截图的 COM 释放

**Files:**
- Modify: `scripts/excel_camera_screenshot.py`
- Modify: `scripts/blur-excel-column.py`
- Modify: `tests/test_self_contained_tools.py`

**Interfaces:**
- Consumes: 两轮真实执行都稳定复现的隐藏 Excel 进程残留。
- Produces: 显式 COM 代理释放、独立实例退出验证和既有 Excel 共存证据。

- [ ] **Step 1: 写失败测试**

用假 COM 对象记录关闭顺序，并用多工作表与已有 PNG 样本覆盖工作表一致性和覆盖备份；修复前必须失败。

- [ ] **Step 2: 最小实现**

真实截图增加确定性 COM 清理；打码定位复用已安装的 openpyxl 读取列宽，直接移除第二个 COM 实例。截图和打码使用相同工作表及显式范围，已有 PNG 覆盖前集中备份。不引入新依赖，不结束全部 Excel 进程。

- [ ] **Step 3: 运行单元与真实回归**

先运行定向测试，再运行完整测试；随后重复真实截图，确认本脚本 Excel PID 自动退出，并确认受控既有 Excel PID 保持存活。

---

### Task 5: 独立盲评与交付

**Files:**
- Create: `docs/reviews/2026-07-26-excel-master-darwin-review.md`
- Update: `results.tsv`

**Interfaces:**
- Consumes: 随机标签的改前/改后 `SKILL.md`、真实执行结果。
- Produces: 至少 2 个独立 judge 的偏好、最终评分、保留或恢复结论。

- [ ] **Step 1: 准备盲评版本**

复制改前和改后内容为随机标签 A/B，不向 judge 透露身份。

- [ ] **Step 2: 启动至少两个独立 judge**

每个 judge 按九维 rubric 独立评分，并判断三个测试提示下哪版更可靠。

- [ ] **Step 3: 应用棘轮**

仅当总分严格提高且 judge 一致偏好新版时保留；否则用集中备份恢复。

- [ ] **Step 4: 最终验证**

重新运行 16 项测试、技能校验、编码检查、runtime 扫描、资源检查及真实样例。

- [ ] **Step 5: 写复核报告**

记录基线、各轮变化、盲评、验证证据、残余风险和未执行的 Git 动作。

## 自审

- 需求与设计中的每项验收标准均映射到 Task 1–4。
- 无待定项、占位符或新增依赖。
- 所有文件写入由主 agent 串行完成；子 agent 只写互斥验证目录。
