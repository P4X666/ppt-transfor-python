# -*- coding: utf-8 -*-
"""
使用 Docling 将 PowerPoint (PPTX) 文件转换为结构化 JSON。

功能：
- 解析幻灯片结构（逐页提取）
- 提取文本内容、标题层级
- 提取图片引用（文件名、尺寸、位置）
- 提取表格（如有）
- 输出为结构化 JSON，便于下游消费

运行示例：
    python docling.py --input "i:/wefor/ai/knowledge/ppt/UA模版-封面目录结尾.pptx" --output "result.json"
"""

from __future__ import annotations

import argparse
import io
import json
import sys
import hashlib
from pathlib import Path
from typing import Any

from docling_core.types.doc.base import BoundingBox, CoordOrigin
from docling_core.types.doc.document import (
    DoclingDocument,
    GroupItem,
    ProvenanceItem,
    RefItem,
    Size,
    TableData,
)
from docling_core.types.doc.labels import DocItemLabel


def _clean_text(text: str) -> str:
    """
    清理文本中的特殊字符
    
    移除或替换可能影响JSON解析和显示的特殊字符：
    - \u000b 垂直制表符
    - \r 回车符（保留\n作为换行）
    - 其他控制字符
    """
    if text is None:
        return ""
    
    cleaned = text.replace('\u000b', ' ')
    cleaned = cleaned.replace('\r', '')
    
    result = []
    for char in cleaned:
        if char == '\n' or char == '\t':
            result.append(char)
        elif ord(char) < 32:
            result.append(' ')
        else:
            result.append(char)
    
    return ''.join(result).strip()


# DrawingML 命名空间
_A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def _xml_color_to_hex(xml_element) -> Optional[str]:
    """从 DrawingML 颜色 XML 元素中提取十六进制颜色值"""
    tag = xml_element.tag.split('}')[-1]
    val = xml_element.get('val', '')
    if tag == 'srgbClr' and val:
        return f'#{val.upper()}'
    if tag == 'schemeClr':
        return f'theme({val})' if val else 'theme'
    if tag == 'sysClr' and val:
        return f'#{val.upper()}'
    return None


def _parse_solid_fill_color(parent_element) -> Optional[str]:
    """从元素中查找第一个 solidFill 并返回颜色"""
    for sf in parent_element.iter(f'{{{_A_NS}}}solidFill'):
        for child in sf:
            color = _xml_color_to_hex(child)
            if color:
                return color
    return None


def _parse_defRPr(p_elem) -> dict[str, Any]:
    """
    从段落 XML 元素及其祖先的 defRPr（默认 Run 属性）中提取字体样式。
    先查找段落自身 pPr 中的 defRPr，再向上查找 txBody → lstStyle → defRPr。
    """
    style: dict[str, Any] = {}

    # 1. 先查找段落自身 pPr 中的 defRPr
    pPr = p_elem.find(f'{{{_A_NS}}}pPr')
    if pPr is not None:
        for defRPr in pPr.iter(f'{{{_A_NS}}}defRPr'):
            _merge_defRPr_style(defRPr, style)

    # 2. 向上查找 txBody → lstStyle → lvl1pPr → defRPr
    parent = p_elem.getparent()
    for _ in range(4):
        if parent is None:
            break
        tag = parent.tag.split('}')[-1]
        if tag == 'txBody':
            # 查找 lstStyle 中的所有 defRPr
            for defRPr in parent.iter(f'{{{_A_NS}}}defRPr'):
                _merge_defRPr_style(defRPr, style)
            break
        parent = parent.getparent()

    return style


def _merge_defRPr_style(defRPr, style: dict[str, Any]):
    """将 defRPr XML 元素的属性合并到 style dict 中（不覆盖已有值）"""
    # 字号（hundredths of a point）
    sz = defRPr.get('sz')
    if sz and 'font_size_pt' not in style:
        try:
            style['font_size_pt'] = int(sz) / 100.0
        except ValueError:
            pass

    # 粗体/斜体
    b = defRPr.get('b')
    if b is not None and 'bold' not in style:
        style['bold'] = b == '1' or b.lower() == 'true'
    i = defRPr.get('i')
    if i is not None and 'italic' not in style:
        style['italic'] = i == '1' or i.lower() == 'true'

    # 字体颜色
    if 'font_color' not in style:
        font_color = _parse_solid_fill_color(defRPr)
        if font_color:
            style['font_color'] = font_color

    # 字体名称
    if 'font_name' not in style:
        for latin in defRPr.iter(f'{{{_A_NS}}}latin'):
            tf = latin.get('typeface')
            if tf:
                style['font_name'] = tf
                break
    if 'font_name' not in style:
        for ea in defRPr.iter(f'{{{_A_NS}}}ea'):
            tf = ea.get('typeface')
            if tf:
                style['font_name'] = tf
                break


def _get_color_hex(color_obj) -> Optional[str]:
    """获取颜色的十六进制表示"""
    if color_obj is None:
        return None
    try:
        if hasattr(color_obj, 'rgb') and color_obj.rgb is not None:
            rgb = color_obj.rgb
            return '#{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])
        elif hasattr(color_obj, 'theme_color'):
            return str(color_obj.theme_color)
    except Exception:
        pass
    return None


def _extract_font_style(font, p_elem=None) -> dict[str, Any]:
    """提取字体样式（支持从 XML defRPr 补充）"""
    style = {}
    if font is not None:
        if hasattr(font, 'name') and font.name:
            style['font_name'] = font.name
        
        if hasattr(font, 'size') and font.size:
            try:
                style['font_size_pt'] = font.size.pt
            except Exception:
                pass
        
        if hasattr(font, 'bold') and font.bold is not None:
            style['bold'] = bool(font.bold)
        
        if hasattr(font, 'italic') and font.italic is not None:
            style['italic'] = bool(font.italic)
        
        if hasattr(font, 'underline') and font.underline is not None:
            style['underline'] = bool(font.underline)
        
        if hasattr(font, 'color'):
            color_hex = _get_color_hex(font.color)
            if color_hex:
                style['font_color'] = color_hex
    
    # 当 python-pptx 没有返回具体值时，从 defRPr 补充
    if p_elem is not None:
        defaults = _parse_defRPr(p_elem)
        for key in ('font_name', 'font_size_pt', 'bold', 'italic', 'font_color'):
            if key not in style and key in defaults:
                style[key] = defaults[key]
    
    return style


def _extract_shape_style(shape) -> dict[str, Any]:
    """提取形状样式（支持从 XML 补充）"""
    style = {}
    
    # 填充颜色（背景色）—— 先试 python-pptx，再试 XML
    try:
        fill = shape.fill
        if fill.type != 0:  # msoFillNone
            fill_color = _get_color_hex(fill.fore_color)
            if fill_color:
                style['background_color'] = fill_color
    except Exception:
        pass
    
    # 如果 python-pptx 没取到，直接从 XML 中找 solidFill
    if 'background_color' not in style:
        try:
            spPr = shape._element.find(f'{{http://schemas.openxmlformats.org/presentationml/2006/main}}spPr')
            if spPr is not None:
                bg = _parse_solid_fill_color(spPr)
                if bg:
                    style['background_color'] = bg
        except Exception:
            pass
    
    # 线条样式
    try:
        line = shape.line
        if line.width is not None and line.width > 0:
            line_color = _get_color_hex(line.color)
            if line_color:
                style['border_color'] = line_color
            try:
                style['border_width_pt'] = line.width.pt
            except Exception:
                pass
    except Exception:
        pass
    
    # 边框宽度也可以从 XML 读取
    if 'border_width_pt' not in style:
        try:
            spPr = shape._element.find(f'{{http://schemas.openxmlformats.org/presentationml/2006/main}}spPr')
            if spPr is not None:
                ln = spPr.find(f'{{{_A_NS}}}ln')
                if ln is not None:
                    w = ln.get('w')
                    if w:
                        style['border_width_pt'] = int(w) / 12700.0
        except Exception:
            pass
    
    # 形状位置和大小
    try:
        style['left'] = shape.left
        style['top'] = shape.top
        style['width'] = shape.width
        style['height'] = shape.height
    except Exception:
        pass
    
    return style


def _extract_paragraph_style(paragraph, p_elem=None) -> dict[str, Any]:
    """提取段落样式（支持从 lstStyle/lvl1pPr 补充）"""
    style = {}
    
    if hasattr(paragraph, 'alignment') and paragraph.alignment is not None:
        style['alignment'] = str(paragraph.alignment).split('.')[-1].lower()
    
    if hasattr(paragraph, 'space_before') and paragraph.space_before is not None:
        try:
            style['space_before_pt'] = paragraph.space_before.pt
        except Exception:
            pass
    
    if hasattr(paragraph, 'space_after') and paragraph.space_after is not None:
        try:
            style['space_after_pt'] = paragraph.space_after.pt
        except Exception:
            pass
    
    # 从 lstStyle/lvl1pPr 补全对齐方式
    if 'alignment' not in style and p_elem is not None:
        # 向上查找 lstStyle → lvl1pPr
        parent = p_elem.getparent()
        for _ in range(3):
            if parent is None:
                break
            for lvl in parent.iter(f'{{{_A_NS}}}lvl1pPr'):
                algn = lvl.get('algn')
                if algn:
                    style['alignment'] = algn
                    break
            if 'alignment' in style:
                break
            parent = parent.getparent()
    
    return style


def _extract_slide_background_color(slide) -> Optional[str]:
    """
    提取幻灯片的背景色。
    优先级：幻灯片自身 > 布局 > 母版
    """
    # 1. 先尝试从幻灯片自身获取
    try:
        bg = slide.background
        if bg.fill.type != 0:  # msoFillNone
            fc = bg.fill.fore_color
            if hasattr(fc, 'rgb') and fc.rgb:
                rgb = fc.rgb
                return f'#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}'
            elif hasattr(fc, 'theme_color') and fc.theme_color:
                return f'theme({fc.theme_color})'
    except Exception:
        pass
    
    # 2. 从布局获取
    try:
        layout = slide.slide_layout
        bg = layout.background
        if bg.fill.type != 0:
            fc = bg.fill.fore_color
            if hasattr(fc, 'rgb') and fc.rgb:
                rgb = fc.rgb
                return f'#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}'
            elif hasattr(fc, 'theme_color') and fc.theme_color:
                return f'theme({fc.theme_color})'
    except Exception:
        pass
    
    # 3. 从母版获取
    try:
        master = slide.slide_layout.slide_master
        bg = master.background
        if bg.fill.type != 0:
            fc = bg.fill.fore_color
            if hasattr(fc, 'rgb') and fc.rgb:
                rgb = fc.rgb
                return f'#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}'
            elif hasattr(fc, 'theme_color') and fc.theme_color:
                return f'theme({fc.theme_color})'
    except Exception:
        pass
    
    return None


def _table_to_html(table) -> str:
    """将PPT表格转换为HTML格式"""
    rows = []
    for row in table.rows:
        row_data = []
        for cell in row.cells:
            text = _clean_text(cell.text)
            row_data.append(text)
        rows.append(row_data)
    
    html_parts = ["<table>"]
    
    for row_idx, row in enumerate(rows):
        if row_idx == 0:
            html_parts.append("  <tr>")
            for cell in row:
                html_parts.append(f"    <th>{cell}</th>")
            html_parts.append("  </tr>")
        else:
            html_parts.append("  <tr>")
            for cell in row:
                html_parts.append(f"    <td>{cell}</td>")
            html_parts.append("  </tr>")
    
    html_parts.append("</table>")
    
    return '\n'.join(html_parts)


def _extract_images_by_slide(
    pptx_path: str | Path,
    output_dir: str | Path,
) -> tuple[list[dict[str, Any]], dict[int, list[int]]]:
    """
    按幻灯片维度提取图片资源，建立幻灯片与图片的映射关系。

    遍历每张幻灯片自身、其布局、其母版中的图片形状，
    提取图片 blob 并保存到本地，同时记录每张幻灯片关联了哪些图片。

    返回：
        (images, slide_image_map)
        - images: 全局图片列表（去重）
        - slide_image_map: {slide_index(1-based): [image_index, ...]}
    """
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from PIL import Image

    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    prs = Presentation(str(pptx_path))
    images: list[dict[str, Any]] = []
    seen_hashes: dict[str, int] = {}

    def _process_shape_container(
        container, source: str
    ) -> list[int]:
        """遍历 shape 容器，提取图片，返回新增的图片索引列表。"""
        added_indices: list[int] = []
        for shape in container.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                added_indices.extend(_process_shape_container(shape, source))
                continue
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            try:
                image = shape.image
                image_bytes = image.blob
                if not image_bytes:
                    continue
                h = hashlib.sha256(image_bytes).hexdigest()[:16]
                if h in seen_hashes:
                    added_indices.append(seen_hashes[h])
                    continue

                ext = ".png"
                if image.ext:
                    ext = image.ext if image.ext.startswith(".") else f".{image.ext}"
                elif image.content_type:
                    ct = image.content_type
                    if "jpeg" in ct or "jpg" in ct:
                        ext = ".jpg"
                    elif "png" in ct:
                        ext = ".png"
                    elif "gif" in ct:
                        ext = ".gif"
                    elif "svg" in ct:
                        ext = ".svg"

                filename = f"img_{h}{ext}"
                out_path = output_dir / filename
                out_path.write_bytes(image_bytes)

                width_px = height_px = None
                try:
                    with Image.open(io.BytesIO(image_bytes)) as im:
                        width_px, height_px = im.size
                except Exception:
                    pass

                img_index = len(images)
                seen_hashes[h] = img_index
                images.append(
                    {
                        "image_index": img_index,
                        "filename": filename,
                        "extracted_path": str(out_path.resolve()),
                        "source": source,
                        "shape_name": shape.name,
                        "size_bytes": len(image_bytes),
                        "width_px": width_px,
                        "height_px": height_px,
                    }
                )
                added_indices.append(img_index)
            except Exception:
                pass
        return added_indices

    # 1. 提取所有母版图片
    master_image_indices: dict[int, list[int]] = {}
    for m_idx, master in enumerate(prs.slide_masters):
        master_image_indices[m_idx] = _process_shape_container(
            master, f"master:{master.name}"
        )

    # 2. 提取所有布局图片
    layout_image_indices: dict[int, list[int]] = {}
    for l_idx, layout in enumerate(prs.slide_layouts):
        layout_image_indices[l_idx] = _process_shape_container(
            layout, f"layout:{layout.name}"
        )

    # 3. 提取每张幻灯片自身的图片，并建立映射
    slide_image_map: dict[int, list[int]] = {}
    master_list = list(prs.slide_masters)
    layout_list = list(prs.slide_layouts)

    for slide_idx, slide in enumerate(prs.slides, start=1):
        associated: set[int] = set()

        # 幻灯片自身的图片
        own_indices = _process_shape_container(slide, f"slide:{slide_idx}")
        associated.update(own_indices)

        # 继承母版图片
        master = slide.slide_layout.slide_master
        m_idx = next((i for i, m in enumerate(master_list) if m is master), -1)
        if m_idx >= 0:
            associated.update(master_image_indices.get(m_idx, []))

        # 继承布局图片
        layout = slide.slide_layout
        l_idx = next((i for i, l in enumerate(layout_list) if l is layout), -1)
        if l_idx >= 0:
            associated.update(layout_image_indices.get(l_idx, []))

        if associated:
            slide_image_map[slide_idx] = sorted(associated)

    return images, slide_image_map


def _build_slide_tree(doc: DoclingDocument) -> list[dict[str, Any]]:
    """
    遍历 DoclingDocument 的 iterate_items，按页/幻灯片组织内容。
    Docling 对 PPTX 的解析结果中，iterate_items 会按阅读顺序输出元素，
    并附带层级 level（标题深度）。
    """
    slides: list[dict[str, Any]] = []
    current_slide: dict[str, Any] | None = None
    current_slide_no = 0

    for item, level in doc.iterate_items():
        # 通过 item 的 prov 判断所属幻灯片（Docling 对 PPTX 用 page_no 表示幻灯片序号）
        provs = getattr(item, "prov", []) or []
        page_no = provs[0].page_no if provs else current_slide_no

        # 如果进入新幻灯片，创建新节点
        if current_slide is None or page_no != current_slide_no:
            current_slide_no = page_no
            current_slide = {
                "slide_index": int(page_no),
                "elements": [],
            }
            slides.append(current_slide)

        element_type = type(item).__name__
        element: dict[str, Any] = {
            "type": "paragraph" if element_type in ("TextItem", "TitleItem") else 
                    "table" if element_type == "TableItem" else element_type.lower(),
            "level": level,
        }

        text = getattr(item, "text", None)
        if text is not None:
            element["text"] = _clean_text(text)
        
        if element_type == "TableItem" and provs:
            prov = provs[0]
            if hasattr(prov, "extra") and prov.extra and "html" in prov.extra:
                element["html"] = prov.extra["html"]
                element["table_type"] = "simple_table"
        
        label = getattr(item, "label", None)
        if label is not None:
            element["label"] = label.value if hasattr(label, "value") else str(label)

        current_slide["elements"].append(element)

    return slides


def _extract_text_flat(doc: DoclingDocument) -> list[dict[str, Any]]:
    """
    扁平化提取所有文本元素，保留标题层级信息，用于快速检索。
    """
    texts: list[dict[str, Any]] = []
    for item, level in doc.iterate_items():
        text = getattr(item, "text", None)
        if text is None:
            continue
        provs = getattr(item, "prov", []) or []
        page_no = provs[0].page_no if provs else None
        label = getattr(item, "label", None)
        label_name = label.value if hasattr(label, "value") else str(label)
        
        item_type = type(item).__name__
        elem_type = "paragraph" if item_type in ("TextItem", "TitleItem") else item_type.lower()
        
        texts.append(
            {
                "text": _clean_text(text),
                "type": elem_type,
                "level": level,
                "label": label_name,
                "slide_index": int(page_no) if page_no is not None else None,
            }
        )
    return texts


def _extract_tables(doc: DoclingDocument) -> list[dict[str, Any]]:
    """
    提取所有表格元素。
    """
    tables: list[dict[str, Any]] = []
    for item, level in doc.iterate_items():
        item_type = type(item).__name__
        if item_type != "TableItem":
            continue
        
        provs = getattr(item, "prov", []) or []
        page_no = provs[0].page_no if provs else None
        
        table_data = {
            "slide_index": int(page_no) if page_no is not None else None,
        }
        
        if provs:
            prov = provs[0]
            if hasattr(prov, "extra") and prov.extra and "html" in prov.extra:
                table_data["html"] = prov.extra["html"]
                table_data["table_type"] = "simple_table"
        
        tables.append(table_data)
    
    return tables


def _extract_tables_from_pptx(pptx_path: str | Path) -> list[dict[str, Any]]:
    """
    直接从 PPTX 文件中提取表格（不依赖 docling）
    """
    from pptx import Presentation
    
    prs = Presentation(str(pptx_path))
    tables: list[dict[str, Any]] = []
    
    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                html_content = _table_to_html(table)
                tables.append({
                    "slide_index": slide_idx,
                    "html": html_content,
                    "table_type": "simple_table"
                })
    
    return tables


def _extract_elements_from_pptx(pptx_path: str | Path) -> list[dict[str, Any]]:
    """
    直接从 PPTX 文件中提取所有元素（用于补充 docling 的不足）
    包含完整的样式信息：字体颜色、背景色、字体样式等
    """
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    
    prs = Presentation(str(pptx_path))
    slides_data: list[dict[str, Any]] = []
    
    for slide_idx, slide in enumerate(prs.slides, start=1):
        slide_data = {
            "slide_index": slide_idx,
            "elements": []
        }
        
        # 提取幻灯片背景色（从母版继承）
        bg_color = _extract_slide_background_color(slide)
        if bg_color:
            slide_data["background_color"] = bg_color
        
        for shape in slide.shapes:
            # 文本框
            if shape.has_text_frame:
                paragraphs_data = []
                
                for paragraph in shape.text_frame.paragraphs:
                    p_elem = paragraph._p
                    runs_data = []
                    
                    for run in paragraph.runs:
                        font_style = _extract_font_style(run.font, p_elem)
                        runs_data.append({
                            "text": _clean_text(run.text),
                            "style": font_style
                        })
                    
                    paragraph_style = _extract_paragraph_style(paragraph, p_elem)
                    paragraphs_data.append({
                        "runs": runs_data,
                        "style": paragraph_style
                    })
                
                shape_style = _extract_shape_style(shape)
                
                full_text = _clean_text(shape.text_frame.text)
                if full_text:
                    element = {
                        "type": "paragraph",
                        "text": full_text,
                        "label": "text",
                        "paragraphs": paragraphs_data,
                        "shape_style": shape_style
                    }
                    slide_data["elements"].append(element)
            
            # 表格
            elif shape.has_table:
                table = shape.table
                html_content = _table_to_html(table)
                
                # 提取表格样式
                shape_style = _extract_shape_style(shape)
                
                # 提取单元格样式
                cell_styles = []
                for row in table.rows:
                    row_styles = []
                    for cell in row.cells:
                        cell_style = {}
                        try:
                            # 从 XML 直接读取单元格背景色
                            tc = cell._tc
                            bg = _parse_solid_fill_color(tc)
                            if bg:
                                cell_style['background_color'] = bg
                            
                            # 单元格文本样式
                            if cell.text:
                                for paragraph in cell.text_frame.paragraphs:
                                    pe = paragraph._p
                                    fs = _parse_defRPr(pe)
                                    if fs:
                                        cell_style['font_style'] = fs
                                        break
                        except Exception:
                            pass
                        row_styles.append(cell_style)
                    cell_styles.append(row_styles)
                
                slide_data["elements"].append({
                    "type": "table",
                    "html": html_content,
                    "table_type": "simple_table",
                    "shape_style": shape_style,
                    "cell_styles": cell_styles
                })
            
            # 图片
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_style = _extract_shape_style(shape)
                # 可以在这里添加图片样式信息
        
        slides_data.append(slide_data)
    
    return slides_data


def _extract_headings(doc: DoclingDocument) -> list[dict[str, Any]]:
    """
    提取所有标题/章节头，构建目录结构。
    """
    headings: list[dict[str, Any]] = []
    for item, level in doc.iterate_items():
        text = getattr(item, "text", None)
        if text is None:
            continue
        label = getattr(item, "label", None)
        label_name = label.value if hasattr(label, "value") else str(label)
        if label_name in ("SECTION_HEADER", "TITLE", "HEADING"):
            provs = getattr(item, "prov", []) or []
            page_no = provs[0].page_no if provs else None
            headings.append(
                {
                    "text": text,
                    "level": level,
                    "slide_index": int(page_no) if page_no is not None else None,
                }
            )
    return headings


def _convert_pptx_with_docling_core(pptx_path: str | Path) -> DoclingDocument:
    """
    使用 docling_core 的能力将 PPTX 转为 DoclingDocument。
    由于当前环境缺少完整的 docling 包（空壳），我们借助 python-pptx 读取 PPTX 内容，
    然后手动构建 DoclingDocument。
    """
    from pptx import Presentation

    prs = Presentation(str(pptx_path))
    doc = DoclingDocument(name=Path(pptx_path).stem)

    # 遍历幻灯片
    for slide_idx, slide in enumerate(prs.slides, start=1):
        # 注册页面
        try:
            width = prs.slide_width
            height = prs.slide_height
        except Exception:
            width = height = 0
        doc.add_page(page_no=slide_idx, size=Size(width=width, height=height))

        # 为每页创建一个 GroupItem 作为容器
        slide_group = GroupItem(
            name=f"slide_{slide_idx}",
            self_ref=f"#/groups/{len(doc.groups)}",
        )
        doc.groups.append(slide_group)
        doc.body.children.append(RefItem(cref=slide_group.self_ref))

        for shape in slide.shapes:
            # 文本框 / 标题 / 段落
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if not text:
                    continue
                # 简单启发式判断标题：字号大或位于顶部
                is_title = False
                try:
                    if shape.text_frame.paragraphs:
                        font_size = shape.text_frame.paragraphs[0].font.size
                        if font_size and font_size.pt > 24:
                            is_title = True
                except Exception:
                    pass

                # 构建 bbox（使用 shape 的位置和尺寸）
                try:
                    bbox = BoundingBox(
                        l=shape.left,
                        t=shape.top,
                        r=shape.left + shape.width,
                        b=shape.top + shape.height,
                        coord_origin=CoordOrigin.TOPLEFT,
                    )
                except Exception:
                    bbox = BoundingBox(
                        l=0, t=0, r=0, b=0, coord_origin=CoordOrigin.TOPLEFT
                    )

                prov = ProvenanceItem(
                    page_no=slide_idx,
                    bbox=bbox,
                    charspan=(0, len(text)),
                )
                if is_title:
                    doc.add_title(text=text, orig=text, prov=prov, parent=slide_group)
                else:
                    doc.add_text(
                        label=DocItemLabel.TEXT,
                        text=text,
                        orig=text,
                        prov=prov,
                        parent=slide_group,
                    )

            # 图片
            if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                try:
                    bbox = BoundingBox(
                        l=shape.left,
                        t=shape.top,
                        r=shape.left + shape.width,
                        b=shape.top + shape.height,
                        coord_origin=CoordOrigin.TOPLEFT,
                    )
                except Exception:
                    bbox = BoundingBox(
                        l=0, t=0, r=0, b=0, coord_origin=CoordOrigin.TOPLEFT
                    )
                try:
                    prov = ProvenanceItem(
                        page_no=slide_idx, bbox=bbox, charspan=(0, 0)
                    )
                    doc.add_picture(prov=prov, parent=slide_group)
                except Exception:
                    pass

            # 表格
            if shape.has_table:
                try:
                    table = shape.table
                    html_content = _table_to_html(table)
                    rows = []
                    for row in table.rows:
                        row_data = []
                        for cell in row.cells:
                            row_data.append(_clean_text(cell.text))
                        rows.append(row_data)
                    table_data = TableData(
                        num_rows=len(rows),
                        num_cols=len(rows[0]) if rows else 0,
                        table_cells=[],
                    )
                    bbox = BoundingBox(
                        l=shape.left,
                        t=shape.top,
                        r=shape.left + shape.width,
                        b=shape.top + shape.height,
                        coord_origin=CoordOrigin.TOPLEFT,
                    )
                    prov = ProvenanceItem(
                        page_no=slide_idx, bbox=bbox, charspan=(0, 0)
                    )
                    prov.extra = {"html": html_content}
                    doc.add_table(data=table_data, prov=prov, parent=slide_group)
                except Exception:
                    pass

    return doc


def pptx_to_json(
    pptx_path: str | Path,
    output_json_path: str | Path | None = None,
    extract_images: bool = True,
    images_output_dir: str | Path | None = None,
) -> dict[str, Any]:
    """
    将 PPTX 转换为结构化 JSON。

    参数：
        pptx_path: 输入 PPTX 文件路径
        output_json_path: 输出 JSON 文件路径（可选）
        extract_images: 是否提取图片到本地
        images_output_dir: 图片输出目录（默认与 JSON 同级目录下的 images/）

    返回：
        结构化字典，可直接 json.dumps
    """
    pptx_path = Path(pptx_path)
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX 文件不存在: {pptx_path}")

    # 1. 转换
    doc = _convert_pptx_with_docling_core(pptx_path)

    # 2. 按幻灯片维度提取图片，建立映射
    images: list[dict[str, Any]] = []
    slide_image_map: dict[int, list[int]] = {}
    if extract_images:
        if images_output_dir is None:
            images_output_dir = pptx_path.parent / "images"
        images, slide_image_map = _extract_images_by_slide(pptx_path, images_output_dir)

    # 3. 构建 JSON 结构
    slides = _build_slide_tree(doc)
    
    # 4. 直接从 PPTX 提取包含样式的元素（补充 docling 的不足）
    elements_from_pptx = _extract_elements_from_pptx(pptx_path)
    
    # 5. 使用包含样式的元素替换原有元素
    for slide in slides:
        slide_idx = slide["slide_index"]
        
        # 添加图片
        if slide_idx in slide_image_map:
            img_indices = slide_image_map[slide_idx]
            slide["images"] = [images[i] for i in img_indices if i < len(images)]
        else:
            slide["images"] = []
        
        # 找到对应的PPTX元素数据
        pptx_slide_data = next((s for s in elements_from_pptx if s["slide_index"] == slide_idx), None)
        if pptx_slide_data:
            # 直接使用包含样式的元素
            slide["elements"] = pptx_slide_data["elements"]
            # 添加幻灯片背景色
            if "background_color" in pptx_slide_data:
                slide["background_color"] = pptx_slide_data["background_color"]

    # 提取表格信息
    tables = []
    for slide_data in elements_from_pptx:
        for elem in slide_data["elements"]:
            if elem.get("type") == "table":
                tables.append({
                    "slide_index": slide_data["slide_index"],
                    "html": elem.get("html"),
                    "table_type": elem.get("table_type"),
                    "shape_style": elem.get("shape_style"),
                    "cell_styles": elem.get("cell_styles"),
                })

    data: dict[str, Any] = {
        "meta": {
            "source_file": str(pptx_path.resolve()),
            "filename": pptx_path.name,
            "conversion_status": "SUCCESS",
            "slide_count": len(doc.pages),
        },
        "headings": _extract_headings(doc),
        "slides": slides,
        "texts_flat": _extract_text_flat(doc),
        "tables": tables,
        "images": images,
        "slide_image_map": slide_image_map,
    }

    # 4. 可选保存 JSON
    if output_json_path:
        output_json_path = Path(output_json_path)
        output_json_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"JSON 已保存: {output_json_path.resolve()}")

    return data


def main() -> None:
    parser = argparse.ArgumentParser(description="PPTX 转 JSON（基于 Docling）")
    parser.add_argument("--input", "-i", required=True, help="输入 PPTX 文件路径")
    parser.add_argument("--output", "-o", default=None, help="输出 JSON 文件路径（默认不保存文件，仅打印）")
    parser.add_argument("--images-dir", default=None, help="图片提取输出目录")
    parser.add_argument("--no-extract-images", action="store_true", help="跳过图片提取")
    args = parser.parse_args()

    data = pptx_to_json(
        pptx_path=args.input,
        output_json_path=args.output,
        extract_images=not args.no_extract_images,
        images_output_dir=args.images_dir,
    )

    if not args.output:
        print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
