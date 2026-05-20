# PPT转换与生成项目 - 技术文档

## 目录

1. [项目概述](#1-项目概述)
2. [JSON结构规范](#2-json结构规范)
3. [模块说明](#3-模块说明)
4. [使用指南](#4-使用指南)
5. [常见问题与注意事项](#5-常见问题与注意事项)

---

## 1. 项目概述

### 1.1 项目简介

本项目提供PPT与JSON之间的双向转换能力，基于 `python-pptx` 库实现：

- **PPT转JSON**: 解析 `.pptx` 文件，提取所有视觉元素为结构化JSON数据
- **JSON转PPT**: 读取JSON数据，精确还原为 `.pptx` 文件

### 1.2 项目结构

```
ppt-transfor/
├── main.py                    # 主程序入口
├── config.py                  # 配置文件
├── core/
│   ├── __init__.py
│   ├── ppt_to_json.py         # PPT转JSON核心模块
│   └── json_to_ppt.py         # JSON转PPT核心模块
├── utils/
│   ├── __init__.py
│   ├── color_utils.py         # 颜色处理工具
│   └── image_utils.py         # 图片处理工具
├── output/                    # 默认输出目录
├── assets/                    # 资源目录
│   └── images/               # 图片资源
└── TECHNICAL_DOCUMENTATION.md # 本文档
```

### 1.3 依赖环境

- Python 3.8+
- python-pptx >= 1.0.0
- Pillow (可选，用于获取图片尺寸)

---

## 2. JSON结构规范

### 2.1 顶层结构

```json
{
  "metadata": { ... },      // PPT元数据
  "slides": [ ... ]         // 幻灯片数组
}
```

### 2.2 字段详细说明

#### 2.2.1 metadata (元数据)

| 字段名 | 数据类型 | 说明 | 示例 |
|--------|----------|------|------|
| `file_path` | string | 原始PPT文件路径 | `"ppt/template.pptx"` |
| `slide_width` | integer | 幻灯片宽度 (EMU单位) | `12192000` |
| `slide_height` | integer | 幻灯片高度 (EMU单位) | `6858000` |
| `slide_width_inches` | float | 幻灯片宽度 (英寸) | `13.333` |
| `slide_height_inches` | float | 幻灯片高度 (英寸) | `7.5` |
| `slide_count` | integer | 幻灯片总数 | `3` |

> **EMU单位说明**: 1 英寸 = 914400 EMU (English Metric Units)

#### 2.2.2 slides (幻灯片数组)

每张幻灯片包含以下字段：

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `slide_index` | integer | 幻灯片序号 (从1开始) |
| `slide_id` | integer | 幻灯片唯一标识 |
| `layout` | object | 幻灯片布局信息 |
| `background` | object | 背景样式信息 |
| `shapes` | array | 形状元素数组 |

#### 2.2.3 layout (布局信息)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `name` | string | 布局名称 |
| `placeholder_count` | integer | 占位符数量 |
| `placeholders` | array | 占位符详情数组 |

**placeholder 对象结构**:

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `idx` | integer | 占位符索引 |
| `type` | string | 占位符类型枚举 |
| `name` | string | 占位符名称 |
| `left` | integer | 左侧位置 (EMU) |
| `top` | integer | 顶部位置 (EMU) |
| `width` | integer | 宽度 (EMU) |
| `height` | integer | 高度 (EMU) |

#### 2.2.4 background (背景信息)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `type` | string | 背景填充类型 |
| `fill` | object | 填充详情 |

**fill 对象结构**:

| 字段名 | 数据类型 | 说明 | 可能取值 |
|--------|----------|------|----------|
| `type` | string | 填充类型 | `"SOLID"`, `"GRADIENT"`, `"BACKGROUND"`, `null` |
| `color` | string | 填充颜色 (HEX) | `"#FF0000"`, `null` |
| `theme_color` | string | 主题颜色 | `"ACCENT_1"`, `null` |
| `brightness` | float | 亮度调整值 | `-0.5` ~ `0.5`, `null` |
| `gradient_stops` | array | 渐变停止点 (渐变时) | 见下方说明 |

**gradient_stop 对象结构**:

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `position` | float | 停止点位置 (0.0 ~ 1.0) |
| `color` | string | 停止点颜色 (HEX) |

#### 2.2.5 shapes (形状元素)

每个形状元素的基础结构：

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `shape_id` | integer | 形状唯一标识 |
| `name` | string | 形状名称 |
| `shape_type` | string | 形状类型枚举 |
| `left` | integer | 左侧位置 (EMU) |
| `top` | integer | 顶部位置 (EMU) |
| `width` | integer | 宽度 (EMU) |
| `height` | integer | 高度 (EMU) |
| `rotation` | float | 旋转角度 |
| `position` | object | 英寸单位的位置信息 |
| `element_type` | string | 元素分类类型 |

**position 对象结构**:

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `left_inches` | float | 左侧位置 (英寸) |
| `top_inches` | float | 顶部位置 (英寸) |
| `width_inches` | float | 宽度 (英寸) |
| `height_inches` | float | 高度 (英寸) |

**element_type 取值说明**:

| 取值 | 说明 | 额外字段 |
|------|------|----------|
| `"picture"` | 图片 | `image_info`, `crop` |
| `"text_box"` | 文本框 | `text_frame`, `fill`, `line` |
| `"auto_shape"` | 自动形状 | `auto_shape_type`, `text_frame`, `fill`, `line`, `adjustments` |
| `"placeholder"` | 占位符 | `placeholder_format`, `text_frame`, `fill` |
| `"group"` | 组合形状 | `shapes` (子形状数组) |
| `"table"` | 表格 | `table` |
| `"chart"` | 图表 | `chart_type`, `has_title` |
| `"line"` | 线条 | `line`, `begin`, `end` |
| `"freeform"` | 自由形状 | `fill`, `line` |
| `"generic"` | 通用形状 | `text_frame`, `fill`, `line` |

#### 2.2.6 text_frame (文本框详情)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `text` | string | 完整文本内容 |
| `paragraphs` | array | 段落数组 |
| `margin_left` | integer | 左边距 (EMU) |
| `margin_top` | integer | 上边距 (EMU) |
| `margin_right` | integer | 右边距 (EMU) |
| `margin_bottom` | integer | 下边距 (EMU) |
| `word_wrap` | boolean | 是否自动换行 |
| `auto_size` | string | 自动调整大小模式 |
| `vertical_anchor` | string | 垂直对齐方式 |

**auto_size 可能取值**:
- `"NONE (0)"` - 不自动调整
- `"SHAPE_TO_FIT_TEXT (1)"` - 形状适应文本
- `"TEXT_TO_FIT_SHAPE (2)"` - 文本适应形状

**vertical_anchor 可能取值**:
- `"TOP (1)"` - 顶部对齐
- `"MIDDLE (3)"` - 居中对齐
- `"BOTTOM (4)"` - 底部对齐

#### 2.2.7 paragraph (段落)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `text` | string | 段落文本 |
| `level` | integer | 缩进层级 (0-8) |
| `alignment` | string | 水平对齐方式 |
| `line_spacing` | float | 行距 |
| `space_before` | integer | 段前间距 (EMU) |
| `space_after` | integer | 段后间距 (EMU) |
| `runs` | array | 文本运行数组 |

**alignment 可能取值**:
- `"LEFT (1)"` - 左对齐
- `"CENTER (2)"` - 居中对齐
- `"RIGHT (3)"` - 右对齐
- `"JUSTIFY (4)"` - 两端对齐
- `"DISTRIBUTE (5)"` - 分散对齐

#### 2.2.8 run (文本运行)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `text` | string | 文本内容 |
| `font` | object | 字体样式 |

**font 对象结构**:

| 字段名 | 数据类型 | 说明 | 可能取值 |
|--------|----------|------|----------|
| `name` | string | 字体名称 | `"微软雅黑"`, `"Arial"`, `null` |
| `size` | integer | 字体大小 (EMU) | `1270000` (约10pt), `null` |
| `size_pt` | float | 字体大小 (磅) | `10.0`, `null` |
| `bold` | boolean | 是否粗体 | `true`, `false`, `null` |
| `italic` | boolean | 是否斜体 | `true`, `false`, `null` |
| `underline` | boolean | 是否下划线 | `true`, `false`, `null` |
| `strike` | boolean | 是否删除线 | `true`, `false`, `null` |
| `color` | object | 字体颜色 | 见 color 对象 |
| `highlight_color` | string | 高亮颜色 | `"YELLOW"`, `null` |

**color 对象结构**:

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `type` | string | 颜色类型 |
| `rgb` | string | RGB颜色 (HEX) |
| `theme_color` | string | 主题颜色 |

#### 2.2.9 image_info (图片信息)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `filename` | string | 图片文件名 |
| `original_ext` | string | 原始格式扩展名 |
| `width` | integer | 图片宽度 (像素) |
| `height` | integer | 图片高度 (像素) |
| `hash` | string | MD5哈希值 (用于去重) |
| `base64` | string | Base64编码数据 |

#### 2.2.10 line (线条样式)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `type` | string | 线条类型 |
| `color` | object | 线条颜色 (同 fill 结构) |
| `width` | integer | 线条宽度 (EMU) |
| `width_pt` | float | 线条宽度 (磅) |
| `dash_style` | string | 虚线样式 |

**dash_style 可能取值**:
- `"SOLID (1)"` - 实线
- `"DASH (2)"` - 虚线
- `"DOT (3)"` - 点线
- `"DASH_DOT (4)"` - 点划线
- `"DASH_DOT_DOT (5)"` - 双点划线

#### 2.2.11 table (表格)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `rows` | integer | 行数 |
| `columns` | integer | 列数 |
| `cells` | array | 单元格数组 |

**cell 对象结构**:

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `row` | integer | 行索引 (从0开始) |
| `column` | integer | 列索引 (从0开始) |
| `text` | string | 单元格文本 |
| `text_frame` | object | 文本框详情 |
| `fill` | object | 单元格填充 |

---

## 3. 模块说明

### 3.1 core/ppt_to_json.py

**PPTToJsonConverter 类**

核心转换器，负责将PPT解析为JSON。

**主要方法**:

| 方法名 | 说明 |
|--------|------|
| `convert()` | 执行完整的PPT到JSON转换 |
| `_extract_metadata(prs)` | 提取PPT元数据 |
| `_extract_slide(slide, idx)` | 提取单张幻灯片信息 |
| `_extract_shape(shape)` | 提取形状元素信息 |
| `_extract_text_frame(shape)` | 提取文本框信息 |
| `_extract_picture(shape)` | 提取图片信息 |
| `_extract_table(shape)` | 提取表格信息 |
| `save_json(data, path)` | 保存JSON到文件 |

**便捷函数**:

```python
def convert_ppt_to_json(pptx_path: str, output_dir: str, json_filename: str = "output.json") -> str
```

### 3.2 core/json_to_ppt.py

**JsonToPPTConverter 类**

核心转换器，负责将JSON还原为PPT。

**主要方法**:

| 方法名 | 说明 |
|--------|------|
| `load_json()` | 加载JSON数据 |
| `convert(output_path)` | 执行JSON到PPT转换 |
| `_create_slide(prs, data)` | 创建单张幻灯片 |
| `_add_shape(slide, data)` | 添加形状到幻灯片 |
| `_apply_text_frame(tf, data)` | 应用文本框内容 |
| `_apply_paragraph(p, data)` | 应用段落格式 |
| `_apply_run(run, data)` | 应用文本运行格式 |
| `_apply_fill(fill, data)` | 应用填充样式 |
| `_apply_line(line, data)` | 应用线条样式 |

**便捷函数**:

```python
def convert_json_to_ppt(json_path: str, output_pptx_path: str, assets_dir: str = None) -> str
```

### 3.3 utils/color_utils.py

**颜色处理工具函数**:

| 函数名 | 说明 |
|--------|------|
| `rgb_color_to_hex(color)` | RGBColor转HEX字符串 |
| `hex_to_rgb_color(hex_str)` | HEX字符串转RGBColor |
| `theme_color_to_string(tc)` | 主题颜色枚举转字符串 |
| `get_fill_color(fill)` | 从填充对象提取颜色信息 |

### 3.4 utils/image_utils.py

**图片处理工具函数**:

| 函数名 | 说明 |
|--------|------|
| `extract_image(part, dir, idx)` | 从PPT提取并保存图片 |
| `load_image_from_base64(b64, path)` | 从Base64加载图片 |
| `get_image_path_from_json(info, dir)` | 从JSON获取图片路径 |

---

## 4. 使用指南

### 4.1 命令行使用

#### 完整转换 (PPT -> JSON -> PPT)

```bash
python main.py convert --input "template.pptx" --output "./output"
```

#### 仅PPT转JSON

```bash
python main.py ppt-to-json --input "template.pptx" --output "./output"
```

#### 仅JSON转PPT

```bash
python main.py json-to-ppt --input "./output/output.json" --output "reconstructed.pptx"
```

### 4.2 编程使用

#### PPT转JSON

```python
from core.ppt_to_json import convert_ppt_to_json

json_path = convert_ppt_to_json(
    pptx_path="template.pptx",
    output_dir="./output",
    json_filename="output.json"
)
```

#### JSON转PPT

```python
from core.json_to_ppt import convert_json_to_ppt

ppt_path = convert_json_to_ppt(
    json_path="./output/output.json",
    output_pptx_path="reconstructed.pptx",
    assets_dir="./output/assets/images"
)
```

#### 高级自定义

```python
from core.ppt_to_json import PPTToJsonConverter
from core.json_to_ppt import JsonToPPTConverter

# 自定义PPT转JSON
converter = PPTToJsonConverter("input.pptx", "./output")
data = converter.convert()
# 可在此修改data...
converter.save_json(data, "./output/custom.json")

# 自定义JSON转PPT
converter = JsonToPPTConverter("custom.json", "./assets/images")
converter.convert("output.pptx")
```

---

## 5. 常见问题与注意事项

### 5.1 已知限制

1. **组合形状**: python-pptx不支持直接创建组合形状，组合中的子形状会平铺添加
2. **图表**: 图表仅保留类型信息，数据需要单独处理
3. **动画效果**: 动画和切换效果不在转换范围内
4. **母版样式**: 母版级别的样式可能无法完全还原

### 5.2 字体兼容性

- 还原PPT时，如果目标环境缺少原PPT使用的字体，会自动使用替代字体
- 建议在相同环境中进行转换和还原，确保字体一致性

### 5.3 图片处理

- 图片以Base64编码嵌入JSON，同时保存为独立文件
- 大图片会导致JSON文件体积增大
- 可通过修改 `config.py` 中的 `IMAGE_CONFIG` 调整图片处理策略

### 5.4 颜色精度

- 颜色以HEX格式存储，支持RGB颜色空间
- 主题颜色会记录类型，但还原时可能因主题不同而有差异
- 渐变填充支持多停止点，但复杂渐变可能不完全一致

### 5.5 位置精度

- 所有位置使用EMU单位 (914400 EMU = 1英寸)
- 还原时位置精度取决于python-pptx的实现

### 5.6 错误处理

- 转换过程使用日志记录所有警告和错误
- 单个元素处理失败不会中断整体转换
- 建议启用DEBUG级别日志进行问题排查

### 5.7 性能优化建议

1. 对于大型PPT，考虑分批处理
2. 图片去重功能默认启用，可减少存储空间
3. 不需要Base64编码时，可在配置中关闭以减小JSON体积

---

## 附录: 形状类型枚举

常见 `shape_type` 取值：

| 枚举值 | 说明 |
|--------|------|
| `AUTO_SHAPE (1)` | 自动形状 |
| `TEXT_BOX (17)` | 文本框 |
| `PICTURE (13)` | 图片 |
| `PLACEHOLDER (14)` | 占位符 |
| `GROUP (6)` | 组合形状 |
| `TABLE (19)` | 表格 |
| `CHART (3)` | 图表 |
| `LINE (20)` | 线条 |
| `FREEFORM (5)` | 自由形状 |
