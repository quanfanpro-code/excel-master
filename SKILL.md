---
name: excel-master
description: "自包含的 Excel 生成、美化与复杂区块格式化技能。使用 pandas + openpyxl 生成或美化 xlsx，保留公式和值，支持12套主题、冻结表头、类型推断和显式条件格式适配。当用户说出个Excel、美化表格、导出报表、生成xlsx、格式化表格、整理Excel或摩根士丹利格式时使用。"
---

# excel-master

## 什么时候用

用户要求"出个Excel"、"导出报表"、"生成表格"、"保存为xlsx"、"美化/格式化这个Excel"时加载本 skill。用户未提供文件时，通过 Windows 文件选择窗口获取路径，不要求用户手工输入绝对路径。普通参数只在确实缺失时问一次；下文安全检查点不受此限制。用户已说明的信息不得重复询问。

## 前置依赖

```bash
pip install pandas "openpyxl>=3.0"
```

生成、美化和复杂区块格式化不需要 xlwings，也不调用 Excel 进程；列宽使用确定性估算法。

本 skill 不依赖其他 skill、Hermes 目录或外部项目源码。

## 执行工作流

### 1. 明确输入与输出

- 输入：用户选择的 CSV/xlsx、目标工作表、主题和明确的数据区块。
- 动作：确认文件存在，输出路径位于用户指定范围；CSV 中订单号、编码、账号等标识列在读取时显式使用字符串类型。
- 输出：任务类型、输入路径、输出路径和不能改变的内容清单。

从 Python 调用时，把本目录的 `scripts` 加入模块搜索路径后再导入：

```python
import sys
sys.path.insert(0, r'技能目录\scripts')
from make_excel import make_excel, beautify
```

### 2. 选择唯一入口

| 用户目标 | 入口 | 关键约束 |
|---|---|---|
| CSV/DataFrame 生成报表 | `scripts/make_excel.py` / `make_excel()` | 标识列读取前声明为字符串 |
| 美化普通 xlsx | `beautify()` | 只改格式，保留值和公式 |
| 多数据块或复杂表头 | `scripts/format_excel_sections.py` | 必须显式提供矩形区块，不自动猜测 |

### 🔴 CHECKPOINT 1 · 🛑 STOP：修改前保护

修改已有文件，或准备覆盖已存在的输出文件前，必须把对应文件备份到 `C:\Users\<用户名>\BackUp\excel-master_时间戳\` 并校验副本内容一致。备份校验失败时停止，不能继续美化或覆盖。

### 3. 识别不能自动决定的内容

- 多数据块：明确每个区块的表头行、结束行、起始列和结束列。
- 条件格式：默认保持；只有用户明确要求主题适配时才传 `adapt_conditional_formatting=True`。
- 冻结窗格：一个工作表只能有一处；`freeze_rows` 作用于整张工作表，不能分别冻结第二数据块。
- 来源不明的 xlsx：先确认公式仍真实存在；缓存值不能恢复为公式。

### 🔴 CHECKPOINT 2 · 🛑 STOP：业务含义确认

复杂区块、条件格式适配或来源不明公式没有明确依据时停止。取得明确范围后再执行；不得靠猜测继续。

### 4. 执行

优先使用下文现成命令或函数。不要复制主题、类型推断和边框逻辑到临时脚本；复杂布局脚本已经复用核心规则。

### 5. 验证

至少检查：

1. 输出文件可重新打开。
2. 值和公式与输入逐项一致；生成模式核对行列数。
3. 金额、百分比、日期和标识列格式正确。
4. 条件格式规则数量和类型未被意外改变。
5. 冻结窗格、表头、边框、网格线和复杂区块符合参数。

完整交付检查使用 `references/implementation-checklist.md`。

### 🔴 CHECKPOINT 3 · 🛑 STOP：交付门

任何必需输出的验证失败都不得交付。先恢复原文件或保留未修改输入，报告具体差异。

### 6. 交付

报告输出文件绝对路径、备份位置、实际验证结果和仍存在的限制。

## 失败处理

| 触发条件 | 第一处理 | 仍失败时 |
|---|---|---|
| 输入文件不存在或打不开 | 重新通过 Windows 文件选择窗口取得文件 | 停止，不创建空输出 |
| 备份复制或内容校验失败 | 保持原文件不动，在统一 `C:\Users\<用户名>\BackUp\` 根目录下换新时间戳目录重试一次 | 停止并报告备份失败 |
| CSV 标识列前导零可能丢失 | 在 `read_csv` 时对订单号、编码、账号列指定字符串类型 | 已丢失时回到原始来源，禁止猜测补零 |
| 表头或多数据块范围误判 | 改用显式区块坐标 | 范围仍不明确时停止并要求业务确认 |
| 输出目录不可写 | 改用用户选择的可写目录 | 停止，保留输入和备份 |
| 公式、值或条件格式验证不一致 | 不交付输出，重新从原文件执行 | 恢复备份并报告差异 |

## 快速使用

```python
from make_excel import make_excel, beautify
import pandas as pd

# 从零生成（DataFrame → 摩根系 Excel）
df = pd.read_csv('data.csv', dtype={'订单号': 'string'})
make_excel(df, '报表.xlsx')

# 多 sheet
make_excel([('汇总', df_summary), ('明细', df_detail)], '报表.xlsx')

# 切换色系主题（默认: default 水蓝）
make_excel(df, '报表-深海蓝.xlsx', theme='deep-navy')
make_excel(df, '报表-珊瑚橙.xlsx', theme='coral')

# 自定义数字格式（覆盖默认的 #,##0.00 等）
make_excel(df, '报表.xlsx', fmt_override={'money': '#,##0.0', 'pct': '0.0%'})

# 冻结前3行；传0表示取消冻结
make_excel(df, '报表.xlsx', freeze_rows=3)

# beautify — 只改格式不改数据（保留公式）
beautify('已有报表.xlsx')                    # 原地美化，自动备份
beautify('已有报表.xlsx', '美化后.xlsx')      # 输出文件不存在时另存

# beautify + 手动指定列类型
beautify('订单表.xlsx', col_types={'订单号': 'text', '金额': 'money'})

# beautify + 切换主题
beautify('已有报表.xlsx', '报表-墨玉绿.xlsx', theme='jade')

# beautify + 自定义数字格式
beautify('已有报表.xlsx', '美化后.xlsx', fmt_override={'money': '#,##0'})

# 显式让已有色阶跟随主题；默认不改变条件格式语义
beautify('已有报表.xlsx', theme='coral', adapt_conditional_formatting=True)
```

## 命令行

```bash
python scripts/make_excel.py 数据.csv 输出.xlsx
python scripts/make_excel.py --beautify 已有报表.xlsx 美化后.xlsx
python scripts/make_excel.py --beautify 已有报表.xlsx 美化后.xlsx --theme coral
python scripts/make_excel.py --beautify 已有报表.xlsx 美化后.xlsx --freeze-rows 3
python scripts/make_excel.py 数据.csv 深海蓝报表.xlsx --theme deep-navy

# 明确指定复杂表格区块：表头行:结束行:起始列号:结束列号
python scripts/format_excel_sections.py 复杂表.xlsx 输出.xlsx --section 3:10:2:12
```

## 色系主题

通过 `theme` 参数切换配色，共 12 套：

| 主题名 | 中文 | 表头 | 合计行 | 特点 |
|--------|------|------|--------|------|
| `default` | 水蓝 | #4472C4 白字 | #D9E2F3 | 经典摩根系默认 |
| `deep-navy` | 深海蓝 | #1F4E79 白字 | #D6E4F0 | 更深沉专业 |
| `jade` | 墨玉绿 | #375623 白字 | #E2EFDA | 清爽自然 |
| `slate` | 陨石灰蓝 | #404040 白字 | #D9D9D9 | 现代极简 |
| `burgundy` | 勃艮第红 | #843C0C 白字 | #FCE4D6 | 暖色调高级感 |
| `coral` | 珊瑚橙 | #D84B4B 白字 | #FDE8E8 | 活泼明亮 |
| `sakura` | 樱花粉 | #D94F70 白字 | #FDE8EE | 温柔活力 |
| `warm-sun` | 暖阳橙 | #E8843A 白字 | #FEF0E6 | 阳光醒目 |
| `lavender` | 薰衣草紫 | #8B7EC8 白字 | #EDEAF8 | 知性温婉 |
| `matcha` | 抹茶绿 | #7BA23F 白字 | #F0F5E6 | 清新自然 |
| `peach` | 蜜桃 | #E8897C 白字 | #FEF0ED | 柔和暖色 |
| `misty` | 雾蓝紫 | #6B7FB5 白字 | #EDF0F8 | 现代冷色 |

用法：

```python
make_excel(df, '输出.xlsx', theme='deep-navy')
beautify('输入.xlsx', '输出.xlsx', theme='coral')
```

## 参考文件

- `references/type-inference-rules.md` — 列类型推断关键词规则和优先级
- `references/implementation-checklist.md` — 交付前验证清单
- `references/conditional-formatting.md` — 条件格式规则和安全适配方法
- `references/complex-layout-formatting.md` — 多级表头、多数据块的显式区块处理
- `scripts/format_excel_sections.py` — 复杂布局显式区块格式化
- `references/project-history.md`、`references/audit-2026-06-11.md` — 历史背景，只供追溯，不是运行依赖

## 摩根系标准 — 9 大原则（来自《为什么精英都是Excel控》）

### 已自动化的原则

| # | 原则 | 实现 |
|---|------|------|
| 1 | 行高 18 | 所有数据行 height=18 |
| 2 | 英文字体 Arial，字号统一 11 | 所有单元格 Font(name='Arial', size=11)。文本列默认黑，数字列叠加颜色 |
| 3 | 数字千分撇区隔 | money → #,##0.00，number → #,##0，pct → 0.00% |
| 7 | 框线上下粗中间虚线无竖线 | TOP_BORDER / MID_BORDER / BOTTOM_BORDER，left/right=None |
| 8 | 文字左对齐，数字右对齐 + 表头全部右对齐 | LEFT_ALIGN / RIGHT_ALIGN，表头统一右对齐 |
| 9 | 不从 A1 开始 | `make_excel` 从 B1 写表头、A列留空；`beautify` 保留原始数据位置，不搬动单元格 |
| — | 隐藏网格线 | ws.sheet_view.showGridLines = False |
| — | 冻结表头 | `freeze_rows` 指定；默认冻结检测到的表头，0表示取消 |
| — | 右侧空白列（宽 3） | 数据区域最后一列右侧加一栏，宽度 3 |
| — | 数字颜色：值=主题色，公式=黑 | 默认主题保留经典蓝；其他主题使用匹配主题的数据色 |
| — | 水蓝表头(#4472C4)白字粗体 | HDR_FILL + HDR_FONT + HDR_ALIGN |

### 需要人工控制的原则（不自动化，但要知道）

这些依赖业务上下文和数据层级关系，自动化会做出错误判断。做表时提醒用户注意。

| # | 原则 | 为什么不自动做 |
|---|------|---------------|
| 4 | 项目下细项要缩排（空白列列宽1） | 需要知道数据层级关系（哪些行是父级、哪些是子级），无法自动推断 |
| 5 | 单位自成一栏（元/个/%单独一列） | 单位信息通常混在列名或值中，抽离为独立列会改变数据结构 |
| 6 | 同一层级栏宽统一 | 需要知道哪几列属于同一层级。当前用字符估算法单独调宽 |
| — | 空单元格填 N/A / N/M | 内容层面的决策，不知道哪些是"不需要填"哪些是"漏填" |
| — | 隐藏行/列用群组而非隐藏 | 群组需要知道哪些行/列属于同一逻辑组，无法自动判断 |
| — | 绿色字体（跨工作表引用） | 需要检测公式中的跨表引用，格式是 `=Sheet2!A1`——可以加但当前未实现 |
| — | 仅3种颜色、淡色背景 | 设计原则，不是代码能控制的 |
| — | 边输入边格式 | 工作习惯，不是输出结果 |

## 设计原则

1. **生成和美化核心零 xlwings** — 不启动 Excel 外部进程。xlwings 保存时会丢掉 openpyxl 的边框/颜色/数字格式，之前为了 autofit 做了三步修补循环（写→保存→xlwings打开→读列宽→写回→再保存），这是架构级别的脆弱性。改为字符估算法后，一次保存完事。
2. **一次保存** — wb.save() 只调一次，没有三步修补循环。格式始终完整。
3. **确定性的核心格式引擎** — 相同数据与参数使用相同的格式规则；原地修改会先写入统一备份目录。
4. **核心单文件** — 生成和美化逻辑集中在 `make_excel.py`；复杂布局只用一个小型辅助脚本，不复制主题和类型规则。
5. **自包含** — 所有运行脚本和参考都在本目录；除声明的 Python 库外，不依赖其他技能或项目路径。

## 红线与反例黑名单

| 禁止动作 | 风险 | 正确做法 |
|---|---|---|
| 未备份就覆盖已有文件，或传 `backup=False` 绕过保护 | 原文件损坏后无法恢复 | 先写入统一 BackUp 目录并校验副本 |
| 用普通 `read_csv` 读取订单号、编码、账号等标识列 | 前导零永久丢失 | 读取前把标识列声明为字符串 |
| 猜测多数据块或表头范围 | 格式落到错误区域 | 使用明确矩形区块 |
| 改写 `cell.value`、把公式转成缓存值或用 `data_only=True` 后另存 | 业务数据和公式丢失 | `beautify` 只改样式，使用 `data_only=False` 验证 |
| 按 0～1 数值范围自动判成百分比 | 0.84 万元等小额金额被误判 | 百分比只按明确列名关键词判断 |
| 用 xlwings 自动列宽后再次保存 | openpyxl 边框、颜色或数字格式可能丢失 | 使用内置字符宽度估算，一次保存 |
| 静默清空、重建或适配条件格式 | 数据条、图标集和业务阈值被改变 | 默认保持；仅显式适配 colorScale 最大色 |
| 只看表面样式就声称完成 | 公式、值和条件格式错误可能被遗漏 | 按实施清单做程序化核对 |

## 关键格式规则

- 类型优先级：`pct > date > text > money`；含糊列名下的 0～1 浮点数按 money。
- 所有单元格使用 Arial 11；文本黑色，手工数字使用主题数据色，公式使用黑色。
- 公式检测同时使用安全显示文本和 `cell.data_type == 'f'`。
- `beautify()` 默认保持条件格式；显式适配时只调整 colorScale 最大色。
- 一个工作表只有一处冻结窗格；复杂数据块不能分别冻结。
- 右侧空白列宽度为 3；若原表已用满所有列，交付前确认是否接受新增空白列。
- 无效主题直接抛出 `ValueError`，不得静默回退。
- 详细类型规则、条件格式和交付验证分别见 `references/type-inference-rules.md`、`references/conditional-formatting.md`、`references/implementation-checklist.md`。
