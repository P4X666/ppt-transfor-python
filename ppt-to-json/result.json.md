# result.json 文件结构文档

## 概述

`result.json` 是 PPTX 文件转换后的结构化输出文件，包含幻灯片内容、样式信息、图片引用等完整数据。该文件用于下游系统（如 RAG、内容管理系统）消费 PPT 内容。

---

## 整体结构

```json
{
  "meta": {...},           // 元数据信息
  "headings": [...],       // 标题列表
  "slides": [...],         // 幻灯片数组（核心内容）
  "texts_flat": [...],     // 扁平化文本列表（快速检索）
  "tables": [...],         // 表格列表
  "images": [...],         // 全局图片列表
  "slide_image_map": {...} // 幻灯片-图片映射关系
}
```

---

## 字段详细说明

### 1. meta（元数据）

| 字段名 | 类型 | 说明 | 示例 |
|--------|------|------|------|
| `source_file` | string | 源 PPTX 文件的绝对路径 | `"I:\\wefor\\ai\\knowledge\\ppt\\UA模版-封面目录结尾.pptx"` |
| `filename` | string | 源文件名 | `"UA模版-封面目录结尾.pptx"` |
| `conversion_status` | string | 转换状态：`SUCCESS` / `FAILED` | `"SUCCESS"` |
| `slide_count` | integer | PPT 总幻灯片数 | `3` |

---

### 2. headings（标题列表）

类型：`array<object>`

存储文档中的章节标题，用于构建目录结构。当前模板无显式标题标记时为空数组。

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `text` | string | 标题文本 |
| `level` | integer | 标题层级（1=一级标题，2=二级标题...） |
| `slide_index` | integer | 所属幻灯片索引（1-based） |

---

### 3. slides（幻灯片数组）⭐ 核心

类型：`array<object>`

每张幻灯片的完整数据，包含元素、图片和背景色。

#### 3.1 幻灯片基础字段

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `slide_index` | integer | 幻灯片索引（1-based） |
| `background_color` | string | 幻灯片背景色（十六进制） |
| `elements` | array | 幻灯片中的元素列表 |
| `images` | array | 关联的图片列表 |

#### 3.2 elements（元素列表）

元素类型包括：`paragraph`（段落文本）、`table`（表格）、`image`（图片）

##### 3.2.1 paragraph（段落元素）

```json
{
  "type": "paragraph",
  "text": "INSERT TITLE",
  "label": "text",
  "paragraphs": [...],
  "shape_style": {...}
}
```

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `type` | string | 固定值 `"paragraph"` |
| `text` | string | 段落完整文本 |
| `label` | string | 标签类型：`"text"` / `"title"` |
| `paragraphs` | array | 段落对象数组 |
| `shape_style` | object | 形状样式信息 |

##### 3.2.2 paragraphs[].runs（文本段）

```json
{
  "text": "INSERT TITLE",
  "style": {
    "font_name": "Neue Plak Compressed Bold",
    "font_size_pt": 61.0,
    "font_color": "#FFFFFF",
    "bold": true,
    "italic": false
  }
}
```

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `text` | string | 文本内容 |
| `style.font_name` | string | 字体名称 |
| `style.font_size_pt` | number | 字号（磅） |
| `style.font_color` | string | 字体颜色（十六进制） |
| `style.bold` | boolean | 是否粗体 |
| `style.italic` | boolean | 是否斜体 |
| `style.underline` | boolean | 是否下划线 |

##### 3.2.3 paragraphs[].style（段落样式）

```json
{
  "alignment": "ctr",
  "space_before_pt": 12.0,
  "space_after_pt": 6.0
}
```

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `alignment` | string | 对齐方式：`"ctr"`（居中）/ `"l"`（左对齐）/ `"r"`（右对齐） |
| `space_before_pt` | number | 段前间距（磅） |
| `space_after_pt` | number | 段后间距（磅） |

##### 3.2.4 shape_style（形状样式）

```json
{
  "background_color": "#F0F0F0",
  "border_color": "#CCCCCC",
  "border_width_pt": 1.0,
  "left": 3868980,
  "top": 3962130,
  "width": 4454040,
  "height": 1233425
}
```

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `background_color` | string | 背景填充色 |
| `border_color` | string | 边框颜色 |
| `border_width_pt` | number | 边框宽度（磅） |
| `left` | integer | 左边界位置（EMU 单位） |
| `top` | integer | 上边界位置（EMU 单位） |
| `width` | integer | 宽度（EMU 单位） |
| `height` | integer | 高度（EMU 单位） |

> **EMU 单位说明**：PPT 内部使用 English Metric Units，1 EMU = 1/914400 英寸 ≈ 0.028 毫米

##### 3.2.5 table（表格元素）

```json
{
  "type": "table",
  "html": "<table>...</table>",
  "table_type": "simple_table",
  "shape_style": {...},
  "cell_styles": [...]
}
```

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `type` | string | 固定值 `"table"` |
| `html` | string | 表格的 HTML 表示 |
| `table_type` | string | 表格类型：`"simple_table"` |
| `shape_style` | object | 表格形状样式 |
| `cell_styles` | array | 单元格样式二维数组 |

##### 3.2.6 cell_styles（单元格样式）

```json
{
  "background_color": "#000000",
  "font_style": {
    "font_size_pt": 18.0,
    "font_color": "#FFFFFF"
  }
}
```

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `background_color` | string | 单元格背景色 |
| `font_style` | object | 单元格字体样式 |

#### 3.3 images（幻灯片关联图片）

```json
{
  "image_index": 0,
  "filename": "img_b5fa2df3d875c2c1.png",
  "extracted_path": "I:\\wefor\\ai\\knowledge\\ppt\\images\\img_b5fa2df3d875c2c1.png",
  "source": "master:",
  "shape_name": "Image",
  "size_bytes": 697868,
  "width_px": 8371,
  "height_px": 4707
}
```

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `image_index` | integer | 在全局 `images` 数组中的索引 |
| `filename` | string | 提取后的文件名（hash 命名） |
| `extracted_path` | string | 图片绝对路径 |
| `source` | string | 来源：`"master:名称"` / `"layout:名称"` / `"slide:序号"` |
| `shape_name` | string | PPT 中的形状名称 |
| `size_bytes` | integer | 文件大小（字节） |
| `width_px` | integer | 宽度（像素） |
| `height_px` | integer | 高度（像素） |

---

### 4. texts_flat（扁平化文本列表）

类型：`array<object>`

用于快速文本检索的扁平结构，包含所有文本及其位置信息。

```json
{
  "text": "INSERT TITLE",
  "type": "paragraph",
  "level": 2,
  "label": "text",
  "slide_index": 1
}
```

| 字段名 | 类型 | 说明 |
|--------|------|------|
| `text` | string | 文本内容 |
| `type` | string | 元素类型 |
| `level` | integer | 层级 |
| `label` | string | 标签 |
| `slide_index` | integer | 所属幻灯片 |

---

### 5. tables（表格列表）

类型：`array<object>`

所有表格的汇总，结构同 `slides[].elements[].table`。

---

### 6. images（全局图片列表）

类型：`array<object>`

所有提取图片的去重列表，结构同 `slides[].images[]`。

---

### 7. slide_image_map（幻灯片-图片映射）

类型：`object<string, array<integer>>`

建立幻灯片索引与图片索引的映射关系。

```json
{
  "1": [0, 1],    // 幻灯片1 包含 images[0] 和 images[1]
  "2": [0],       // 幻灯片2 包含 images[0]
  "3": [0, 3]     // 幻灯片3 包含 images[0] 和 images[3]
}
```

---

## 数据流转示意图

```
PPTX 文件
    │
    ▼
┌─────────────────────────────────────────────────────────────────┐
│                    转换流程                                    │
├─────────────────────────────────────────────────────────────────┤
│  1. 解析幻灯片结构  →  slides[].slide_index                    │
│  2. 提取文本内容    →  slides[].elements[].text               │
│  3. 提取样式信息    →  slides[].elements[].paragraphs[].style │
│  4. 提取图片资源    →  images[] + slide_image_map              │
│  5. 提取表格        →  slides[].elements[].table              │
│  6. 提取背景色      →  slides[].background_color              │
└─────────────────────────────────────────────────────────────────┘
    │
    ▼
result.json
```

---

## 使用场景示例

### 场景1：提取所有文本内容

```python
import json

with open("result.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# 方式1：从 texts_flat 快速获取
all_texts = [item["text"] for item in data["texts_flat"]]

# 方式2：从 slides 完整获取
all_texts = []
for slide in data["slides"]:
    for elem in slide["elements"]:
        if elem["type"] == "paragraph":
            all_texts.append(elem["text"])
```

### 场景2：获取特定幻灯片的背景色

```python
slide_idx = 1
slide = next(s for s in data["slides"] if s["slide_index"] == slide_idx)
bg_color = slide.get("background_color", "#FFFFFF")  # 默认白色
```

### 场景3：获取幻灯片关联的图片路径

```python
slide_idx = 1
slide = next(s for s in data["slides"] if s["slide_index"] == slide_idx)
image_paths = [img["extracted_path"] for img in slide["images"]]
```

---

## 版本说明

| 版本 | 日期 | 变更说明 |
|------|------|----------|
| v1.0 | 2026-05 | 初始版本 |
| v1.1 | 2026-05 | 新增 `background_color` 字段 |
| v1.2 | 2026-05 | 新增完整样式信息（字体、段落、形状） |

---

## 字段兼容性矩阵

| 字段 | docling.py | ppt2json_mineru.py | 说明 |
|------|-----------|-------------------|------|
| `meta` | ✅ | ✅ | 一致 |
| `slides[].background_color` | ✅ | ✅ | 一致 |
| `slides[].elements[].paragraphs[].style` | ✅ | ✅ | 一致 |
| `slides[].elements[].shape_style` | ✅ | ✅ | 一致 |
| `slide_image_map` | ✅ | ✅ | 一致 |
| `raw_content_list` | ❌ | ✅ | MinerU 特有 |
| `raw_middle_json` | ❌ | ✅ | MinerU 特有 |