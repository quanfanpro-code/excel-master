# 实施前扫码清单

在调用 `make_excel` 或 `beautify` 之前，确认以下事项。

## 前置确认

- [ ] 数据来源：DataFrame / CSV / 已有 xlsx
- [ ] 如果输入是已有 xlsx（可能含公式）→ **必须用 beautify**，不是 make_excel
- [ ] 输出路径：最终产物放 D:\\hermes_files\\YYYY-MM-DD\\，临时测试放 Desktop
- [ ] 列类型推断是否需要手动覆盖（col_types 参数）
- [ ] beautify 原地覆盖是否需要关闭自动备份（backup=False）

## 验证清单（交付前逐项检查）

用以下代码遍历所有单元格核验。这是用户（本义）要求的级别——他逐格检查 A1 是否空白、每个单元格边框/颜色/数字格式是否正确。不能只目测说"看起来可以"。

```python
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re

wb = load_workbook(path)
ws = wb.active

# 0. 合计行检测（如果末尾行含"合计/小计/总计"）
last_row_val = ws.cell(ws.max_row, 1).value or ws.cell(ws.max_row, 2).value
has_summary = last_row_val and any(kw in str(last_row_val) for kw in ['合计','小计','总计','sum','total'])

# 1. A1 必须空白无样式（make_excel 模式）
# beautify 模式会保留 A 列原有数据，所以 A1 可能有值。但应无样式
a1 = ws.cell(1, 1)
if a1.value is None:
    assert a1.font.color is None or a1.fill.start_color.rgb == '00000000', '空A1不应有样式'

# 2. 表头：水蓝底白字粗体右对齐
for col in range(2, ws.max_column + 1):
    c = ws.cell(1, col)
    assert c.fill.start_color.rgb == '004472C4', f'R1C{col} fill应为#4472C4'
    assert c.font.color.rgb == '00FFFFFF', f'R1C{col} font应为白色'
    assert c.font.bold == True, f'R1C{col} 应加粗'
    assert c.alignment.horizontal == 'right', f'R1C{col} 应右对齐'

# 3. 边框无竖线
for r in range(1, ws.max_row + 1):
    for c in range(2, ws.max_column + 1):
        cell = ws.cell(r, c)
        assert cell.border.left.style is None, f'R{r}C{c} 有左边框'
        assert cell.border.right.style is None, f'R{r}C{c} 有右边框'
        # 上端=medium，下端=medium，中间=dashed
        if r == 1:
            assert cell.border.top.style == 'medium', f'R{r}C{c} top应为medium'
        elif r == ws.max_row and not has_summary:
            assert cell.border.bottom.style == 'medium', f'R{r}C{c} bottom应为medium'
        elif r > 1 and r < ws.max_row:
            assert cell.border.top.style == 'dashed', f'R{r}C{c} 中间行top应为dashed'
            assert cell.border.bottom.style == 'dashed', f'R{r}C{c} 中间行bottom应为dashed'

# 4. 行高 18
for r in range(1, ws.max_row + 1):
    assert ws.row_dimensions[r].height == 18, f'R{r} 行高应为18'

# 5. 数据行字体统一 Arial 11
for r in range(2, ws.max_row + 1):
    for c in range(2, ws.max_column + 1):
        cell = ws.cell(r, c)
        assert cell.font.name == 'Arial', f'R{r}C{c} font应为Arial'
        assert cell.font.size == 11, f'R{r}C{c} size应为11'

# 6. 数字列颜色：make_excel 全蓝；beautify 值=蓝 公式=黑
def _infer_num_col(hdr):
    n = str(hdr or '').lower()
    return bool(re.search(r'(金额|价格|收入|成本|费用|毛利|净利|合计|总额|单价|预算|率$|占比)', n))

for c in range(2, ws.max_column + 1):
    if not _infer_num_col(ws.cell(1, c).value):
        continue  # 跳过文本列
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(r, c)
        val = cell.value
        fc = cell.font.color.rgb if cell.font.color else None
        if isinstance(val, str) and val.startswith('='):
            assert fc == '00000000', f'R{r}C{c} 公式应黑色(000000), got {fc}'
        elif isinstance(val, (int, float)):
            assert fc == '000000FF', f'R{r}C{c} 数值应蓝色(0000FF), got {fc}'

# 7. 网格线隐藏
assert ws.sheet_view.showGridLines == False

print('✅ 格式验证全部通过')
```

## 常见问题快速排查

| 症状 | 原因 | 修复 |
|------|------|------|
| 公式变成了值 | 用了 make_excel 而不是 beautify | 改用 beautify |
| 列格式不对（如数字成文本） | 类型推断没命中或误判 | 加 col_types 参数手动指定 |
| 边框有竖线 | 边框 left/right 没设 None | 检查 Border() 的 left/right 参数 |
| 表格不是从B列开始 | 数据包含A列或索引列 | beautify 保留原位置；make_excel 用 reset_index() |
| 表格右侧多了一列空白 | 右侧空白列宽3 | 手动删除 |
| A1 有值 | beautify 保留 A 列数据 | 正常现象，不处理 |
| 数字颜色不对 | font 设了 DATA_FONT 但被后续覆盖 | 检查 _apply_styles 中 font 赋值顺序 |
| 边框验证失败 | beautify 的合计行 bottom=medium 位置可能偏移 | 检查 data_end 是否准确对应合计行 |
