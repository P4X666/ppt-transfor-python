# PPT转换与生成项目 - 技术文档

## 目录

1. [项目概述](#1-项目概述)
2. [JSON结构规范](#2-json结构规范)
3. [模块说明](#3-模块说明)
4. [插件系统](#4-插件系统)
5. [辅助工具](#5-辅助工具)
6. [使用指南](#6-使用指南)
7. [常见问题与注意事项](#7-常见问题与注意事项)

---

## 1. 项目概述

### 1.1 项目简介

本项目提供PPT与JSON之间的双向转换能力，基于 `python-pptx` 库实现：

- **PPT转JSON**: 解析 `.pptx` 文件，提取所有视觉元素为结构化JSON数据
- **JSON转PPT**: 读取JSON数据，精确还原为 `.pptx` 文件
- **图标/图片提取**: 从指定页面提取图标和图片资源，保持原始格式和质量

### 1.2 项目结构

```
ppt-transfor/
├── main.py                    # 主程序入口（CLI）
├── config.py                  # 配置文件
├── extract_page_icons.py      # 辅助工具：按页面提取底部区域图标
├── extract_slide_images.py    # 辅助工具：三级层级扫描提取图片
├── core/
│   ├── __init__.py
│   ├── ppt_to_json.py         # PPT转JSON核心模块
│   └── json_to_ppt.py         # JSON转PPT核心模块
├── plugins/
│   ├── __init__.py
│   ├── base.py                # 插件基类定义
│   └── partition_detector.py  # 分区检测插件
├── utils/
│   ├── __init__.py
│   ├── color_utils.py         # 颜色处理工具
│   └── image_utils.py         # 图片处理工具
├── output/                    # 默认输出目录
├── input/                     # 输入文件目录
└── TECHNICAL_DOCUMENTATION.md # 本文档
```

### 1.3 依赖环境

| 依赖 | 版本要求 | 说明 |
|------|----------|------|
| Python | >= 3.8 | 推荐 3.12+，注意 3.12 起 hasattr 行为变更 |
| python-pptx | >= 1.0.0 | PPT文件解析与生成 |
| Pillow | 可选 | 用于获取图片尺寸和格式信息 |
| lxml | 自动随python-pptx安装 | XML解析（主题颜色提取等） |

> **Python 3.12+ 兼容性提示**: 从 Python 3.12 开始，`hasattr()` 不再抑制 `ValueError`（仅抑制 `AttributeError`）。`python-pptx` 库中某些属性（如 `auto_shape_type`、`adjustments`）在遇到未映射的枚举值时会抛出 `ValueError`，本项目已通过 `try/except` 妥善处理此问题。

---

## 2. JSON结构规范

### 2.1 顶层结构

```json
{
  "metadata": { ... },        // PPT元数据
  "theme_colors": { ... },    // 主题颜色方案
  "slides": [ ... ]           // 幻灯片数组
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

#### 2.2.2 theme_colors (主题颜色方案)

从PPT的Slide Master主题中提取的颜色方案，键名为主题颜色标识，值为HEX颜色字符串。

```json
{
  "dk1": "#14141E",
  "lt1": "#B3B3BF",
  "dk2": "#A7A7A7",
  "lt2": "#535353",
  "accent1": "#DC1117",
  "accent2": "#14141E",
  "accent3": "#282832",
  "accent4": "#505059",
  "accent5": "#8C8C96",
  "accent6": "#B3B3BF",
  "hlink": "#0000FF",
  "folHlink": "#FF00FF"
}
```

| 键名 | 说明 |
|------|------|
| `dk1` / `lt1` | 深色1 / 浅色1（主要文字/背景色） |
| `dk2` / `lt2` | 深色2 / 浅色2 |
| `accent1` ~ `accent6` | 强调色1~6 |
| `hlink` | 超链接颜色 |
| `folHlink` | 已访问超链接颜色 |

#### 2.2.3 slides (幻灯片数组)

每张幻灯片包含以下字段：

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `slide_index` | integer | 幻灯片序号 (从1开始) |
| `slide_id` | integer | 幻灯片唯一标识 |
| `layout` | object | 幻灯片布局信息 |
| `background` | object | 背景样式信息 |
| `shapes` | array | 形状元素数组 |

#### 2.2.4 layout (布局信息)

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

#### 2.2.5 background (背景信息)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `type` | string | 背景填充类型 |
| `fill` | object | 填充详情 |
| `master_background` | object | 母版背景信息（当fill类型为BACKGROUND时） |

**fill 对象结构**:

| 字段名 | 数据类型 | 说明 | 可能取值 |
|--------|----------|------|----------|
| `type` | string | 填充类型 | `"SOLID"`, `"GRADIENT"`, `"BACKGROUND"`, `null` |
| `color` | string | 填充颜色 (HEX) | `"#FF0000"`, `null` |
| `theme_color` | string | 主题颜色 | `"ACCENT_1"`, `null` |
| `theme_color_hex` | string | 主题颜色对应的HEX值 | `"#DC1117"`, `null` |
| `brightness` | float | 亮度调整值 | `-0.5` ~ `0.5`, `null` |
| `gradient_stops` | array | 渐变停止点 (渐变时) | 见下方说明 |

**gradient_stop 对象结构**:

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `position` | float | 停止点位置 (0.0 ~ 1.0) |
| `color` | string | 停止点颜色 (HEX) |

#### 2.2.6 shapes (形状元素)

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
| `source` | string | 形状来源层级（仅layout/master级别） |

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

**source 取值说明**:

| 取值 | 说明 |
|------|------|
| `"layout"` | 来自Slide Layout的形状 |
| `"master"` | 来自Slide Master的形状 |
| 无此字段 | 来自Slide本身的形状 |

#### 2.2.7 text_frame (文本框详情)

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
| `lst_style` | object | 列表样式（包含各级别默认段落/字体属性） |

**auto_size 可能取值**:
- `"NONE (0)"` - 不自动调整
- `"SHAPE_TO_FIT_TEXT (1)"` - 形状适应文本
- `"TEXT_TO_FIT_SHAPE (2)"` - 文本适应形状

**vertical_anchor 可能取值**:
- `"TOP (1)"` - 顶部对齐
- `"MIDDLE (3)"` - 居中对齐
- `"BOTTOM (4)"` - 底部对齐

**lst_style 对象结构**:

列表样式按级别（level1~level9）存储默认段落和字体属性，每级可能包含：

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `alignment` | string | 对齐方式（如 `"ctr"`） |
| `def_tab_sz` | string | 默认制表位大小 |
| `line_spacing_pct` | string | 行距百分比 |
| `default_run_props` | object | 默认文本运行属性 |

**default_run_props 对象结构**:

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `size` | integer | 字体大小（百分之一磅） |
| `cap` | string | 大写样式 |
| `spc` | string | 字符间距 |
| `color` | string | 字体颜色 (HEX) |
| `color_scheme` | string | 主题颜色方案名 |
| `font_name` | string | 字体名称 |
| `bold` | boolean | 是否粗体（字体名包含"Bold"时推断） |

#### 2.2.8 paragraph (段落)

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

#### 2.2.9 run (文本运行)

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
| `type` | string | 颜色类型（如 `"RGB (1)"`, `"SCHEME (2)"`） |
| `rgb` | string | RGB颜色 (HEX)，如 `"#FFFFFF"` |
| `theme_color` | string | 主题颜色枚举名，如 `"ACCENT_4 (8)"` |
| `theme_color_hex` | string | 主题颜色对应的HEX值，如 `"#505059"` |

#### 2.2.10 image_info (图片信息)

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `filename` | string | 图片文件名 |
| `original_ext` | string | 原始格式扩展名 |
| `width` | integer | 图片宽度 (像素) |
| `height` | integer | 图片高度 (像素) |
| `hash` | string | MD5哈希值 (用于去重) |
| `base64` | string | Base64编码数据 |

#### 2.2.11 line (线条样式)

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

#### 2.2.12 table (表格)

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
| `borders` | object | 单元格边框信息 |

**borders 对象结构**:

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `left` | object | 左边框 |
| `top` | object | 上边框 |
| `right` | object | 右边框 |
| `bottom` | object | 下边框 |

每个边框对象包含：

| 字段名 | 数据类型 | 说明 |
|--------|----------|------|
| `type` | string | 边框类型 |
| `width` | integer | 边框宽度 (EMU) |
| `color` | string | 边框颜色 (HEX) |

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
| `_extract_slide(slide, idx)` | 提取单张幻灯片信息（含layout/master层级） |
| `_extract_shape(shape)` | 提取形状元素信息 |
| `_extract_picture(shape)` | 提取图片信息 |
| `_extract_text_frame(shape)` | 提取文本框信息 |
| `_extract_auto_shape(shape)` | 提取自动形状信息 |
| `_extract_table(shape)` | 提取表格信息（含单元格边框） |
| `_extract_line(shape)` | 提取线条信息 |
| `_extract_freeform(shape)` | 提取自由形状信息 |
| `_extract_group(shape)` | 提取组合形状信息（递归） |
| `_extract_lst_style(text_frame)` | 提取文本框列表样式 |
| `save_json(data, path)` | 保存JSON到文件 |

**关键特性**:
- 自动提取主题颜色方案（`theme_colors`）
- 三级层级扫描：Slide → Layout → Master
- 图片去重（基于MD5哈希）
- 主题颜色自动解析为HEX值（`theme_color_hex`）
- 兼容 Python 3.12+ 的 `hasattr` 行为变更

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
| `_parse_shape_type(shape_type_str)` | 解析形状类型字符串为枚举 |

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
| `get_fill_color(fill)` | 从填充对象提取颜色信息（含theme_color） |

### 3.4 utils/image_utils.py

**图片处理工具函数**:

| 函数名 | 说明 |
|--------|------|
| `extract_image(part, dir, idx)` | 从PPT提取并保存图片（含Base64编码） |
| `load_image_from_base64(b64, path)` | 从Base64加载图片 |
| `get_image_path_from_json(info, dir)` | 从JSON获取图片路径 |

### 3.5 config.py

**项目配置文件**，集中管理所有可配置参数：

| 配置项 | 说明 |
|--------|------|
| `DEFAULT_OUTPUT_DIR` | 默认输出目录 |
| `DEFAULT_ASSETS_DIR` | 默认资源目录 |
| `IMAGE_CONFIG` | 图片提取配置（Base64、去重、格式） |
| `COLOR_CONFIG` | 颜色输出配置（格式、主题颜色、亮度） |
| `TEXT_CONFIG` | 文本提取配置（run级别、段落格式、空段落） |
| `SLIDE_DIMENSIONS` | PPT尺寸预设（4:3 / 16:9） |

---

## 4. 插件系统

### 4.1 概述

项目提供可扩展的插件接口，允许通过添加新插件来改变形状过滤行为。每个插件实现 `should_include_shape` 方法来决定某个形状是否应被包含在还原的幻灯片中。

### 4.2 插件基类

**SlideShapeFilterPlugin** (定义于 `plugins/base.py`)

| 属性/方法 | 说明 |
|-----------|------|
| `name` | 插件名称（用于日志和调试） |
| `priority` | 插件优先级，数值越小越先执行（默认100） |
| `should_include_shape(shape_data, slide_data, context)` | 判断形状是否应被包含 |

### 4.3 内置插件

#### PartitionDetectorPlugin (分区检测插件)

**文件**: `plugins/partition_detector.py`  
**优先级**: 10

**功能**: 通过分析幻灯片中内容形状的布局位置，自动识别Master级别的区域背景形状是否构成可见分区。

**检测逻辑**:
1. 对于 `source="master"` 的 auto_shape（Rectangle），计算其包围盒
2. 扫描幻灯片中非Master的内容形状，按以下规则判断：
   - 内容形状左边界需在Rectangle右侧阈值之外（右边界+5%宽度）
   - 仅检查 table 和 text_box 类型
   - 形状宽度 ≥ 阈值（max(Rectangle宽度×40%, 2000000 EMU)）
   - 内容形状垂直范围需与Rectangle有交集
3. 满足条件 → Rectangle为活跃分区 → 保留；否则排除

**适用场景**: 左侧浅色背景 + 右侧表格内容的分区页面

---

## 5. 辅助工具

### 5.1 extract_page_icons.py

从PPT指定页面提取底部区域图标资源，按页面编号分别保存并生成报告。

**使用方法**:

```bash
python extract_page_icons.py <pptx文件路径> <页码列表> <输出目录>
```

**示例**:

```bash
# 提取第7页和第8页的底部区域图标
python extract_page_icons.py "ppt/sample.pptx" "7,8" "output"

# 提取多个页面的图标
python extract_page_icons.py "ppt/sample.pptx" "1,2,3" "icons"
```

**输出结构**:

```
output/
├── page_07/                    # 第7页图标
│   ├── icon_0001_xxxxxxxx.png
│   └── icon_0002_xxxxxxxx.png
├── page_08/                    # 第8页图标
│   ├── icon_0001_xxxxxxxx.png
│   └── icon_0002_xxxxxxxx.png
├── extraction_report.json      # JSON格式报告
└── extraction_report.txt       # 文本格式报告
```

**关键特性**:
- 底部区域检测：默认提取页面高度60%以下的图片元素
- 支持全页背景图识别（覆盖整个页面的图片也会被提取）
- 按页面编号分别保存到子目录
- 自动去重（基于图片哈希）
- 生成JSON和文本双格式报告
- 保持原始格式和质量（无损提取）

### 5.2 extract_slide_images.py

从指定页码的幻灯片及其 Layout/Master 中提取全部图片，支持三级层级扫描。

**使用方法**:

直接修改脚本中的 `pptx_path`、`output_dir`、`slide_number` 变量后运行：

```bash
python extract_slide_images.py
```

**三级扫描层级**:

| 层级 | 说明 |
|------|------|
| Slide | 直接形状层，遍历所有shape（含GROUP递归） |
| Layout | 版式层，提取版面背景/装饰图片 |
| Master | 母版层，提取母版图片 |

**关键特性**:
- 支持SVG内嵌图片提取
- 二进制文件头格式检测（PNG/JPEG/GIF/BMP）
- 按blob哈希去重
- 输出按层级分组的汇总信息

---

## 6. 使用指南

### 6.1 命令行使用

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

#### 日志级别设置

```bash
python main.py --log-level DEBUG ppt-to-json --input "template.pptx" --output "./output"
```

### 6.2 编程使用

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

## 7. 常见问题与注意事项

### 7.1 已知限制

1. **组合形状**: python-pptx不支持直接创建组合形状，组合中的子形状会平铺添加
2. **图表**: 图表仅保留类型信息，数据需要单独处理
3. **动画效果**: 动画和切换效果不在转换范围内
4. **母版样式**: 母版级别的样式可能无法完全还原
5. **未映射的形状类型**: 某些 `auto_shape_type` 值（如 `"line"`）在 `MSO_AUTO_SHAPE_TYPE` 枚举中没有XML映射，会回退为 `"unknown"` 并通过XML直接获取 `prstGeom` 属性值

### 7.2 Python版本兼容性

- **Python 3.12+ 重要变更**: `hasattr()` 不再抑制 `ValueError`，仅抑制 `AttributeError`。这会影响 `python-pptx` 中以下属性的访问：
  - `shape.auto_shape_type`: 当 `prstGeom` 的 `prst` 属性值不在枚举映射中时抛出 `ValueError`
  - `shape.adjustments`: 内部访问 `prstGeom.prst` 时同样可能抛出 `ValueError`
  
  本项目已通过 `try/except` 替代 `hasattr` 来处理这些情况。

### 7.3 字体兼容性

- 还原PPT时，如果目标环境缺少原PPT使用的字体，会自动使用替代字体
- 建议在相同环境中进行转换和还原，确保字体一致性

### 7.4 图片处理

- 图片以Base64编码嵌入JSON，同时保存为独立文件
- 大图片会导致JSON文件体积增大
- 可通过修改 `config.py` 中的 `IMAGE_CONFIG` 调整图片处理策略
- 图片去重基于MD5哈希，相同内容的图片只保存一次

### 7.5 颜色精度

- 颜色以HEX格式存储，支持RGB颜色空间
- 主题颜色会记录类型，同时自动解析为HEX值（`theme_color_hex` 字段）
- 渐变填充支持多停止点，但复杂渐变可能不完全一致

### 7.6 位置精度

- 所有位置使用EMU单位 (914400 EMU = 1英寸)
- 还原时位置精度取决于python-pptx的实现

### 7.7 错误处理

- 转换过程使用日志记录所有警告和错误
- 单个元素处理失败不会中断整体转换
- 建议启用DEBUG级别日志进行问题排查：`--log-level DEBUG`

### 7.8 性能优化建议

1. 对于大型PPT，考虑分批处理
2. 图片去重功能默认启用，可减少存储空间
3. 不需要Base64编码时，可在配置中关闭以减小JSON体积

---

## 附录A: 形状类型枚举

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

## 附录B: 主题颜色名称映射

| 主题颜色标识 | 枚举名 | 说明 |
|-------------|--------|------|
| `dk1` | DARK_1 | 深色1 |
| `lt1` | LIGHT_1 | 浅色1 |
| `dk2` | DARK_2 | 深色2 |
| `lt2` | LIGHT_2 | 浅色2 |
| `accent1` | ACCENT_1 | 强调色1 |
| `accent2` | ACCENT_2 | 强调色2 |
| `accent3` | ACCENT_3 | 强调色3 |
| `accent4` | ACCENT_4 | 强调色4 |
| `accent5` | ACCENT_5 | 强调色5 |
| `accent6` | ACCENT_6 | 强调色6 |
| `hlink` | HYPERLINK | 超链接颜色 |
| `folHlink` | FOLLOWED_HYPERLINK | 已访问超链接颜色 |

## 附录C: 配置参考

### IMAGE_CONFIG 图片提取配置

| 配置项 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `save_base64` | bool | `True` | 是否在JSON中嵌入Base64编码的图片 |
| `save_file` | bool | `True` | 是否保存图片文件到磁盘 |
| `deduplicate` | bool | `True` | 是否对图片进行去重 |
| `supported_formats` | list | `["png","jpg","jpeg","gif","bmp","tiff"]` | 支持的图片格式 |

### COLOR_CONFIG 颜色配置

| 配置项 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `format` | string | `"hex"` | 颜色输出格式 |
| `include_theme_color` | bool | `True` | 是否包含主题颜色信息 |
| `include_brightness` | bool | `True` | 是否包含亮度调整信息 |

### TEXT_CONFIG 文本提取配置

| 配置项 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `extract_runs` | bool | `True` | 是否提取文本run级别的格式 |
| `extract_paragraph_format` | bool | `True` | 是否提取段落格式 |
| `preserve_empty_paragraphs` | bool | `False` | 是否保留空段落 |
