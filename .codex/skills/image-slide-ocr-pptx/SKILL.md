# image-slide-ocr-pptx

将“每页都是图片”的 PPTX、单张幻灯片截图，或一组幻灯片图片，重建为保留原始视觉效果且文字可编辑的 PPTX。

## 适用场景

使用本 skill，当用户提出以下需求时：

- 将图片版 PPTX 转换成可编辑 PPTX
- 将幻灯片截图转换成可编辑 PPTX
- OCR 识别每页图片上的文字，并在 PPTX 中放置可编辑文本框
- 保持原图、图表、背景、版式和视觉精美度，同时增加可编辑文字层
- 批量处理多页图片并合并成一个 PPTX
- 替换、追加、合并已经生成的 PPTX 页面

典型用户表述包括：

- “把这个图片转换成可编辑的 pptx，保持所有图像”
- “文字识别出来，确保可以编辑”
- “每页都是图片的 PPTX，逐页 OCR 并重建”
- “保留原图效果，同时文字可编辑”
- “将这些生成的 PPTX 合并成一个文件”

## 目标输出

输出一个 `.pptx` 文件，要求：

1. 每页保留原始图片作为底图，完整覆盖幻灯片画布。
2. 识别出的文字作为独立可编辑文本框叠加在对应位置。
3. 文本框位置、字号、颜色、粗细、对齐方式尽量接近原图。
4. 不要用空文本框冒充 OCR 结果。
5. 页面比例应与原始图片一致，常见为 16:9 宽屏。
6. 多页输入时，按原始顺序合并为一个 PPTX。
7. 最终返回可下载的 PPTX 文件路径。

## 推荐工具与库

优先使用 Python 实现：

- `python-pptx`：创建和编辑 PPTX
- `Pillow`：读取图片尺寸、裁剪、预处理
- `pytesseract` 或系统 OCR：英文/数字 OCR
- 可用时优先使用视觉模型能力读取中文、复杂版式和图表文字
- `opencv-python`：可选，用于图像增强、二值化、检测文本区域

注意：OCR 库通常对中文和复杂版式不稳定。对于中文技术汇报、管理层 PPT、复杂图表页面，应优先使用模型视觉理解结果，再用程序生成文本框。

## 工作流

### 1. 读取输入

支持以下输入：

- 单张图片：`.png`、`.jpg`、`.jpeg`、`.webp`
- 多张图片
- 图片版 `.pptx`
- 已生成的多个单页 `.pptx`

若输入是 PPTX：

1. 将每页导出或渲染为图片。
2. 保留每页原始图像。
3. 对每页图像执行 OCR 与版面重建。

若输入是图片：

1. 读取图片原始宽高。
2. 根据图片比例设置幻灯片尺寸。
3. 将图片作为整页底图插入。

### 2. OCR 与版面解析

对每页识别：

- 标题
- 小标题
- 正文段落
- 图注
- 表格文字
- 坐标轴标签
- 页码、角标、备注
- Logo 附近的固定标识文字

每个文字块至少记录：

```json
{
  "text": "识别出的文字",
  "x": 0.1,
  "y": 0.2,
  "w": 0.3,
  "h": 0.05,
  "font_size": 18,
  "bold": false,
  "color": "#000000",
  "align": "left"
}
```

坐标建议使用归一化比例，最后映射到 PPT 坐标。

### 3. 创建 PPTX

1. 创建空白 PPTX。
2. 根据输入图片比例设置 `slide_width` 和 `slide_height`。
3. 每页使用 blank layout。
4. 插入原图作为底图：
   - `left = 0`
   - `top = 0`
   - `width = slide_width`
   - `height = slide_height`
5. 对 OCR 得到的每个文字块，创建 textbox。
6. 设置文本框为透明填充、无边框。
7. 设置字体、字号、颜色、粗体、对齐方式。
8. 必要时设置文本框文本自动缩放或微调高度，避免截断。

### 4. 文字层质量要求

- 文字必须可选中、可编辑。
- 不允许只放图片而没有文字层，除非该页确实无文字。
- 不允许生成大量空文本框。
- 不允许把整页文字塞进一个巨大文本框。
- 同一视觉块应尽量对应一个文本框。
- 标题、段落、表格、注释应拆分为合理文本框。
- 中文、英文、数字、单位、公式符号要尽量保留。
- 对不确定字符，用最合理猜测，不要留空。

### 5. 坐标映射

若 OCR 输出像素坐标：

```python
ppt_x = px_x / image_width * slide_width
ppt_y = px_y / image_height * slide_height
ppt_w = px_w / image_width * slide_width
ppt_h = px_h / image_height * slide_height
```

推荐 16:9 PPT 尺寸：

```python
from pptx.util import Inches
prs.slide_width = Inches(13.333333)
prs.slide_height = Inches(7.5)
```

若输入图片不是 16:9，应按原图比例设置页面，或在保持完整画面前提下居中填充。

### 6. 合并多个 PPTX

当用户要求合并多个单页 PPTX：

1. 按用户提供顺序排序。
2. 逐页复制到底图和文字层。
3. 确保页面尺寸一致。
4. 输出单个合并后的 PPTX。

如果页面尺寸不同，优先使用第一个文件的尺寸，并对后续页面等比例适配。

### 7. 质检

生成后必须检查：

- PPTX 能正常打开。
- 页数与输入页数一致。
- 每页底图完整覆盖，没有黑边、裁切、变形。
- 每页至少包含图片底图。
- 有文字的页面应包含可编辑文本框。
- 文本框内容非空。
- 文本框数量与页面复杂度大致匹配。
- 文件名清晰，例如：`image_slide_ocr_editable.pptx`。

可以用 `python-pptx` 重新打开输出文件并统计：

```python
from pptx import Presentation
prs = Presentation("output.pptx")
for i, slide in enumerate(prs.slides, start=1):
    text_shapes = [s for s in slide.shapes if hasattr(s, "text") and s.text.strip()]
    pic_shapes = [s for s in slide.shapes if s.shape_type == 13]
    print(i, len(pic_shapes), len(text_shapes))
```

## 实现模板

```python
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image


def hex_to_rgb(hex_color: str):
    hex_color = hex_color.strip().lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def add_image_slide(prs, image_path, text_blocks):
    img = Image.open(image_path)
    img_w, img_h = img.size

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    slide.shapes.add_picture(str(image_path), 0, 0, width=slide_w, height=slide_h)

    for block in text_blocks:
        text = block.get("text", "").strip()
        if not text:
            continue

        x = block["x"] * slide_w
        y = block["y"] * slide_h
        w = block["w"] * slide_w
        h = block["h"] * slide_h

        textbox = slide.shapes.add_textbox(x, y, w, h)
        textbox.fill.background()
        textbox.line.fill.background()

        tf = textbox.text_frame
        tf.clear()
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text

        align = block.get("align", "left")
        p.alignment = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
        }.get(align, PP_ALIGN.LEFT)

        run = p.runs[0]
        run.font.size = Pt(block.get("font_size", 14))
        run.font.bold = bool(block.get("bold", False))
        run.font.name = block.get("font", "Microsoft YaHei")

        color = block.get("color", "#000000")
        r, g, b = hex_to_rgb(color)
        run.font.color.rgb = RGBColor(r, g, b)


def build_pptx(image_paths, ocr_pages, output_path):
    first = Image.open(image_paths[0])
    img_w, img_h = first.size

    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(13.333333 * img_h / img_w)

    for image_path, text_blocks in zip(image_paths, ocr_pages):
        add_image_slide(prs, image_path, text_blocks)

    prs.save(output_path)
```

## 重要注意事项

- 对图片版 PPTX，不要删除原始图片底图；它是视觉保真度的基础。
- OCR 文字层是为了可编辑，不要求完全替代底图。
- 对复杂中文页面，优先保证“可编辑文字尽可能完整”，其次再做像素级位置微调。
- 若用户指出文字框为空、文字未识别、漏字，应重新执行 OCR 或手工补全文字块。
- 若页面较多，应先完成全部页面，再统一合并和质检。
- 不要只输出脚本；必须尽最大努力生成用户要求的 PPTX 文件。

## 输出回复格式

完成后简要说明：

- 已处理页数
- 输出文件链接
- 是否保留原图底图
- 是否添加可编辑文字层

示例：

```text
已完成：共 25 页，保留每页原始图片作为底图，并叠加可编辑 OCR 文字层。
下载：image_slide_ocr_editable.pptx
```
