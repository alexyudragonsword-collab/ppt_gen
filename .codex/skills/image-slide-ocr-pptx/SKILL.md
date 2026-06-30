# image-slide-ocr-pptx

将“每页都是图片”的 PPTX、单张幻灯片截图，或一组幻灯片图片，重建为保留原始视觉效果且文字可编辑的 PPTX。

## 什么时候使用

当用户提出以下需求时，使用本 skill：

- 将图片版 PPTX 转换成可编辑 PPTX
- 将幻灯片截图转换成可编辑 PPTX
- OCR 识别每页图片上的文字，并在 PPTX 中放置可编辑文本框
- 保持原图、图表、背景、版式和视觉精美度，同时增加可编辑文字层
- 批量处理多页图片并合并成一个 PPTX
- 替换、追加、合并已经生成的 PPTX 页面

典型表述：

- “把这个图片转换成可编辑的 pptx，保持所有图像”
- “文字识别出来，确保可以编辑”
- “每页都是图片的 PPTX，逐页 OCR 并重建”
- “保留原图效果，同时文字可编辑”
- “将这些生成的 PPTX 合并成一个文件”

## 已提供的可执行脚本

本 skill 包含可直接运行的 Python 脚本：

```text
.codex/skills/image-slide-ocr-pptx/scripts/image_slide_ocr_pptx.py
```

依赖文件：

```text
.codex/skills/image-slide-ocr-pptx/requirements.txt
```

安装 Python 依赖：

```bash
pip install -r .codex/skills/image-slide-ocr-pptx/requirements.txt
```

如果使用脚本内置 OCR，还需要系统安装 Tesseract OCR 以及对应语言包，例如中文简体 `chi_sim` 和英文 `eng`。如果环境没有 Tesseract，优先使用视觉模型或其它 OCR 工具生成 `--ocr-json`，再调用脚本生成 PPTX。

## 脚本能力

`image_slide_ocr_pptx.py` 支持：

- 输入单张图片、多张图片或图片目录
- 输入图片版 `.pptx`，自动提取每页主图作为底图
- 可选使用 LibreOffice + pdftoppm 将非纯图片 PPTX 渲染成图片
- 使用 `pytesseract` 自动 OCR
- 接收外部 OCR JSON，适合中文复杂页面或视觉模型识别结果
- 每页插入原图作为全页底图
- 将 OCR 结果叠加为可编辑文本框
- 自动设置页面比例、文本框位置、字号、颜色和对齐方式
- 输出 `.pptx`
- 生成后重新打开 PPTX，统计每页图片数和文本框数用于质检

## 常用命令

### 1. 图片目录转可编辑 PPTX

```bash
python .codex/skills/image-slide-ocr-pptx/scripts/image_slide_ocr_pptx.py \
  ./slides_images \
  --output ./image_slide_ocr_editable.pptx \
  --lang chi_sim+eng
```

### 2. 多张图片转可编辑 PPTX

```bash
python .codex/skills/image-slide-ocr-pptx/scripts/image_slide_ocr_pptx.py \
  ./page01.png ./page02.png ./page03.png \
  --output ./image_slide_ocr_editable.pptx \
  --lang chi_sim+eng
```

### 3. 图片版 PPTX 转可编辑 PPTX

```bash
python .codex/skills/image-slide-ocr-pptx/scripts/image_slide_ocr_pptx.py \
  ./image_only_slides.pptx \
  --output ./image_slide_ocr_editable.pptx \
  --lang chi_sim+eng
```

### 4. 使用外部 OCR JSON 生成 PPTX

复杂中文、技术汇报、图表页建议先用视觉模型识别页面文字和坐标，保存为 JSON，再调用脚本：

```bash
python .codex/skills/image-slide-ocr-pptx/scripts/image_slide_ocr_pptx.py \
  ./slides_images \
  --ocr-json ./ocr_blocks.json \
  --output ./image_slide_ocr_editable.pptx
```

### 5. 只保留图片底图，不做 OCR

```bash
python .codex/skills/image-slide-ocr-pptx/scripts/image_slide_ocr_pptx.py \
  ./slides_images \
  --no-ocr \
  --output ./image_only_preserved.pptx
```

## OCR JSON 格式

脚本支持以下 JSON 格式。最外层是页面列表，每页是文字块列表：

```json
[
  [
    {
      "text": "识别出的文字",
      "x": 0.1,
      "y": 0.2,
      "w": 0.3,
      "h": 0.05,
      "font_size": 18,
      "bold": false,
      "color": "#000000",
      "align": "left",
      "font": "Microsoft YaHei"
    }
  ]
]
```

坐标 `x/y/w/h` 使用 0 到 1 的归一化比例，分别表示文字框相对于整页图片的左、上、宽、高。

## 目标输出

输出一个 `.pptx` 文件，要求：

1. 每页保留原始图片作为底图，完整覆盖幻灯片画布。
2. 识别出的文字作为独立可编辑文本框叠加在对应位置。
3. 文本框位置、字号、颜色、粗细、对齐方式尽量接近原图。
4. 不要用空文本框冒充 OCR 结果。
5. 页面比例应与原始图片一致，常见为 16:9 宽屏。
6. 多页输入时，按原始顺序合并为一个 PPTX。
7. 最终返回可下载的 PPTX 文件路径。

## 推荐工作流

### 1. 读取输入

支持以下输入：

- 单张图片：`.png`、`.jpg`、`.jpeg`、`.webp`
- 多张图片
- 图片目录
- 图片版 `.pptx`

若输入是 PPTX：

1. 优先用脚本提取每页主图。
2. 如果不是纯图片页，可尝试 LibreOffice + pdftoppm 渲染。
3. 如果环境无法渲染，先导出 PPTX 为图片，再运行脚本。

若输入是图片：

1. 读取图片原始宽高。
2. 根据图片比例设置幻灯片尺寸。
3. 将图片作为整页底图插入。

### 2. OCR 与版面解析

对每页尽量识别：

- 标题
- 小标题
- 正文段落
- 图注
- 表格文字
- 坐标轴标签
- 页码、角标、备注
- Logo 附近的固定标识文字

对于中文复杂页面，脚本内置 Tesseract OCR 可能不稳定。更推荐由视觉模型生成 `ocr_blocks.json`，再调用脚本生成 PPTX。

### 3. 创建 PPTX

1. 创建空白 PPTX。
2. 根据输入图片比例设置 `slide_width` 和 `slide_height`。
3. 每页使用 blank layout。
4. 插入原图作为底图，覆盖整页。
5. 对 OCR 得到的每个文字块，创建 textbox。
6. 设置文本框为透明填充、无边框。
7. 设置字体、字号、颜色、粗体、对齐方式。
8. 必要时微调文本框高度，避免截断。

### 4. 质量要求

- 文字必须可选中、可编辑。
- 不允许只放图片而没有文字层，除非该页确实无文字或用户指定 `--no-ocr`。
- 不允许生成大量空文本框。
- 不允许把整页文字塞进一个巨大文本框。
- 同一视觉块应尽量对应一个文本框。
- 标题、段落、表格、注释应拆分为合理文本框。
- 中文、英文、数字、单位、公式符号要尽量保留。
- 对不确定字符，用最合理猜测，不要留空。

### 5. 坐标映射

若 OCR 输出像素坐标，转换为归一化坐标：

```python
x = px_x / image_width
y = px_y / image_height
w = px_w / image_width
h = px_h / image_height
```

脚本内部会把归一化坐标映射到 PPT 坐标：

```python
ppt_x = x * slide_width
ppt_y = y * slide_height
ppt_w = w * slide_width
ppt_h = h * slide_height
```

## 质检

生成后必须检查：

- PPTX 能正常打开。
- 页数与输入页数一致。
- 每页底图完整覆盖，没有黑边、裁切、变形。
- 每页至少包含图片底图。
- 有文字的页面应包含可编辑文本框。
- 文本框内容非空。
- 文本框数量与页面复杂度大致匹配。

脚本生成后会输出类似统计：

```text
Created: /path/to/image_slide_ocr_editable.pptx
Pages: 3
Page 1: pictures=1 textboxes=18
Page 2: pictures=1 textboxes=24
Page 3: pictures=1 textboxes=12
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
