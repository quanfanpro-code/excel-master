# 复杂表头和多数据块格式化

## 适用边界

`beautify()` 适合一个连续数据区域。遇到以下结构时，不要让程序自动猜：

- 大标题、单位说明、多级表头连续出现；
- 同一工作表纵向或横向排列多个独立数据块；
- 合并单元格表达父级和子级；
- 不同区块的起始列、结束列不同。

这类工作簿使用本 skill 自带的 `scripts/format_excel_sections.py`，明确给出每个矩形区块。

## 区块定义

每个区块包含四个整数：

```python
{
    "header_row": 3,
    "data_end_row": 10,
    "start_col": 2,
    "end_col": 12,
}
```

- `header_row`：实际列头所在行；
- `data_end_row`：该区块最后一行；
- `start_col`、`end_col`：Excel 列号，A=1、B=2。

区块可以处于相同的行但不同列，也可以上下排列；矩形范围不得重叠。

## 命令行

```bash
python scripts/format_excel_sections.py 输入.xlsx 输出.xlsx \
  --section 3:10:2:12 \
  --section 12:20:2:8 \
  --theme coral \
  --freeze-rows 3
```

`--section` 的格式为：

```text
表头行:结束行:起始列号:结束列号
```

可以重复传入多个区块。

## Python 调用

```python
from format_excel_sections import format_excel_sections

format_excel_sections(
    "输入.xlsx",
    [
        {"header_row": 3, "data_end_row": 10, "start_col": 2, "end_col": 12},
        {"header_row": 12, "data_end_row": 20, "start_col": 2, "end_col": 8},
    ],
    "输出.xlsx",
    theme="deep-navy",
    freeze_rows=3,
)
```

脚本复用 `make_excel.py` 中的主题、类型推断、字体和边框，不维护第二套格式常量。

## 合并单元格处理

脚本不会拆分或填充合并单元格，也不会更改其值。开始前先人工确认：

1. 父级标题是否位于合并区域左上角；
2. `header_row` 是否指向真正的列头，而不是大标题或单位说明；
3. 每个数据块是否有明确结束行；
4. 区块之间是否存在重叠。

如果业务要求把多级表头展开为扁平列名，应先另存副本并人工确认父子关系；格式化脚本只负责样式，不改数据结构。

## 验证

每个区块逐项检查：

- 表头为主题色底、白字、右对齐；
- 首行上边框和末行下边框为粗线；
- 中间行为虚线，无竖线；
- 数字格式和字体颜色符合类型；
- 公式和值未改变；
- 区块之外的单元格未被格式化；
- 冻结窗格符合 `freeze_rows`。
