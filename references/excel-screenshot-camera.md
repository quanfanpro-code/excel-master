# Excel 照相机截图（CopyPicture + Clipboard）

当用户要求"把这个Excel截图发给我"时，**必须使用真实 Excel 文件截图**，不要用 HTML 重新渲染。

HTML 渲染会丢失条件格式（绿色色阶、数据条、图标集）和精准的 Excel 边框样式，用户一眼能看出差异。

## 前置条件

- `pywin32` 已安装（Hermes venv 自带）
- 本机安装了 Microsoft Excel（Windows 环境）
- `PIL/Pillow` 已安装

## 标准流程

```python
import win32com.client as win32
import time
from PIL import ImageGrab

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(xlsx_path)
ws = wb.ActiveSheet

# 1. Copy range as picture (Appearance=xlScreen=1, Format=xlBitmap=2)
ws.Range("B1:K14").CopyPicture(Appearance=1, Format=2)
time.sleep(0.5)

# 2. Paste into a temporary worksheet
tmp_ws = wb.Worksheets.Add()
tmp_ws.Paste()
time.sleep(0.5)

# 3. Find the picture shape (Type=13 is Picture)
pic = None
for shp in tmp_ws.Shapes:
    if shp.Type == 13:  # msoPicture
        pic = shp
        break

# 4. Copy the shape to clipboard
pic.Copy()
time.sleep(0.3)

# 5. Grab from clipboard and save
img = ImageGrab.grabclipboard()
img.save(output_path)

wb.Close(SaveChanges=False)
excel.Quit()
```

## 关键要点

| 步骤 | 说明 |
|------|------|
| `CopyPicture(1, 2)` | 第1参数=xlScreen（屏幕显示效果），第2参数=xlBitmap（位图格式）。不能用默认参数 |
| `Shape.Type == 13` | 13 是 `msoPicture` 类型常量。不要用 11（那是 msoLinkedPicture）|
| `pic.Copy()` | 必须从 shape 再次 Copy，因为 `CopyPicture` 放到 paste sheet 后需要二次复制才能进系统剪贴板 |
| `time.sleep(0.3-0.5)` | 每个 COM 操作之间留间隔，win32com 需要时间同步 |

## 常见坑点

### Range 地址必须写对
如果 Workbook 用的是 B2 起始（摩根标准），截图范围要从 B 列开始。直接写 `Range("B1:M14")`，不要从 A1 开始。

### 无头模式不影响截图
`excel.Visible = False` 下 CopyPicture 仍然正常工作，不需要让 Excel 窗口可见。

### 条件格式保留
这是用真实 Excel 截图的核心优势——ColorScaleRule（绿色色阶）、DataBar、IconSet 都完整保留。HTML+puppeteer 做不到这一点。

### 截图含前后留白
本义要求截图边界比数据区域前后各多一列/一行。如果数据区域是 B1:M14，则截图范围应为 A1:N15：

```python
# 数据在 B1:M14，截图含前后留白
ws.Range("A1:N15").CopyPicture(Appearance=1, Format=2)
# 后续步骤相同
```

原理：左侧多一列 A（空列）让表格边界清晰可见，右侧多一列 N（右侧空白列）让合计列不紧贴图片边缘，上下各多一行留出视觉呼吸空间。

### 截图中打码敏感列
数据安全要求：用户名、店铺名等可识别业务信息的列，在截图发到飞书/TG 前必须模糊处理。

流程分两步：
1. 用 win32com 从 Excel 获取敏感列（如列 C）的精确像素位置
2. 用 PIL 对对应区域做 GaussianBlur

```python
# Step 1: 获取列精确位置
excel = win32.gencache.EnsureDispatch("Excel.Application")
wb = excel.Workbooks.Open(xlsx_path)
ws = wb.ActiveSheet

cap_range = ws.Range("A1:N15")         # 截图范围
col_c = ws.Columns("C")                # 用户名列
x_start_ratio = (col_c.Left - cap_range.Left) / cap_range.Width
x_end_ratio = (col_c.Left + col_c.Width - cap_range.Left) / cap_range.Width
# x_start_ratio ≈ 0.08, x_end_ratio ≈ 0.26 （按列宽比例计算）

wb.Close(SaveChanges=False)
excel.Quit()

# Step 2: 对截图做 blur
from PIL import Image, ImageFilter
img = Image.open(screenshot_path)
w, h = img.size
x1, x2 = int(w * x_start_ratio), int(w * x_end_ratio)
region = img.crop((x1, 0, x2, h))
blurred = region.filter(ImageFilter.GaussianBlur(radius=12))
result = img.copy()
result.paste(blurred, (x1, 0))
result.save(output_path)
```

关键点：
- 用 `Column.Left` / `Column.Width` 和 `cap_range.Left` / `cap_range.Width` 计算比例，不依赖字符宽度估算
- GaussianBlur radius=12 足够覆盖用户名文字
- 只能 blur 整列宽度，因为你没有单个单元格的像素坐标

### 备用方案
如果 Excel 不可用（无许可证、无 pywin32），回退到 openpyxl 读取数据 + 自行构建 HTML 表格 + puppeteer-core 截图。但需告知用户条件格式会丢失。
