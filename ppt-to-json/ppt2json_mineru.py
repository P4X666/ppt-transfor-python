# -*- coding: utf-8 -*-
"""
使用 MinerU 将 PowerPoint (PPTX) 文件转换为结构化 JSON。

功能：
- 调用 MinerU 的 do_parse 解析 PPTX
- 提取文本、图片、表格、标题层级
- 读取 MinerU 输出的 content_list_v2 和 middle.json
- 组织为结构清晰的 JSON 数据

运行示例：
    python ppt2json_mineru.py --input "i:/wefor/ai/knowledge/ppt/UA模版-封面目录结尾.pptx" --output "result_mineru.json"
"""

from __future__ import annotations

import argparse
import json
import logging
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Optional

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def _read_pptx_bytes(pptx_path: str | Path) -> bytes:
    """读取 PPTX 文件为字节流。"""
    pptx_path = Path(pptx_path)
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX 文件不存在: {pptx_path}")
    if pptx_path.suffix.lower() != ".pptx":
        raise ValueError(f"文件不是 PPTX 格式: {pptx_path.suffix}")
    return pptx_path.read_bytes()


# ============================================================
# 样式提取工具函数
# ============================================================


# DrawingML 命名空间
_A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'


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
    """获取颜色的十六进制表示（安全处理 _NoneColor）"""
    if color_obj is None:
        return None
    try:
        if hasattr(color_obj, 'rgb') and color_obj.rgb is not None:
            rgb = color_obj.rgb
            return '#{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])
        if hasattr(color_obj, 'theme_color'):
            tc = color_obj.theme_color
            if tc is not None:
                return str(tc)
    except (AttributeError, Exception):
        pass
    return None


def _extract_font_style(font, p_elem=None) -> dict[str, Any]:
    """
    提取字体样式。
    font: python-pptx Font 对象
    p_elem: 段落 lxml 元素（用于读取 defRPr 默认值）
    """
    style: dict[str, Any] = {}
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
    """提取形状样式（背景色、边框、位置坐标）"""
    style: dict[str, Any] = {}

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

    # 形状位置和大小（EMU 单位）
    try:
        style['left'] = shape.left
        style['top'] = shape.top
        style['width'] = shape.width
        style['height'] = shape.height
    except Exception:
        pass

    return style


def _extract_paragraph_style(paragraph, p_elem=None) -> dict[str, Any]:
    """提取段落样式（对齐、间距），支持从 defRPr/lstStyle 补充"""
    style: dict[str, Any] = {}

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


def _clean_text(text: Any) -> str:
    """清理文本中的特殊字符"""
    if text is None:
        return ""
    s = str(text)
    cleaned = s.replace('\u000b', ' ')
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


def _extract_layout_images(
    pptx_path: str | Path,
    output_dir: str | Path,
) -> tuple[list[dict[str, Any]], dict[int, list[int]]]:
    """
    MinerU 只提取 slide.shapes 中的图片，不处理母版/布局中的图片。
    此函数补充提取 slide master 和 slide layout 中的图片资源，
    并建立每张幻灯片与其关联图片的映射关系。

    返回：
        (images, slide_image_map)
        - images: 全局图片列表
        - slide_image_map: {slide_index(1-based): [image_index, ...]}
    """
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from PIL import Image
    import io
    import hashlib

    output_dir = Path(output_dir)
    images_dir = output_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    prs = Presentation(str(pptx_path))
    images: list[dict[str, Any]] = []
    seen_hashes: dict[str, int] = {}  # hash -> image_index

    def _process_shape_container(container, source: str) -> list[int]:
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

                filename = f"layout_img_{h}{ext}"
                out_path = images_dir / filename
                out_path.write_bytes(image_bytes)

                width_px = height_px = None
                try:
                    with Image.open(io.BytesIO(image_bytes)) as im:
                        width_px, height_px = im.size
                except Exception:
                    pass

                img_index = len(images)
                seen_hashes[h] = img_index
                images.append({
                    "image_index": img_index,
                    "filename": filename,
                    "extracted_path": str(out_path.resolve()),
                    "source": source,
                    "shape_name": shape.name,
                    "size_bytes": len(image_bytes),
                    "width_px": width_px,
                    "height_px": height_px,
                })
                added_indices.append(img_index)
            except Exception as exc:
                logger.debug(f"提取布局图片时跳过 shape {shape.name}: {exc}")
        return added_indices

    # 1. 先提取所有母版图片，建立全局索引
    master_image_indices: dict[int, list[int]] = {}  # master_idx -> [img_indices]
    for m_idx, master in enumerate(prs.slide_masters):
        master_image_indices[m_idx] = _process_shape_container(master, f"master:{master.name}")

    # 2. 再提取所有布局图片
    layout_image_indices: dict[int, list[int]] = {}  # layout_idx -> [img_indices]
    for l_idx, layout in enumerate(prs.slide_layouts):
        layout_image_indices[l_idx] = _process_shape_container(layout, f"layout:{layout.name}")

    # 3. 建立每张幻灯片与图片的映射
    slide_image_map: dict[int, list[int]] = {}
    master_list = list(prs.slide_masters)
    layout_list = list(prs.slide_layouts)
    for slide_idx, slide in enumerate(prs.slides, start=1):
        associated: set[int] = set()

        # 幻灯片继承其母版的图片
        master = slide.slide_layout.slide_master
        m_idx = next((i for i, m in enumerate(master_list) if m is master), -1)
        if m_idx >= 0:
            associated.update(master_image_indices.get(m_idx, []))

        # 幻灯片继承其布局的图片
        layout = slide.slide_layout
        l_idx = next((i for i, l in enumerate(layout_list) if l is layout), -1)
        if l_idx >= 0:
            associated.update(layout_image_indices.get(l_idx, []))

        if associated:
            slide_image_map[slide_idx] = sorted(associated)

    logger.info(f"从母版/布局中补充提取了 {len(images)} 张图片")
    return images, slide_image_map


def _parse_with_mineru(
    pptx_bytes: bytes,
    filename_stem: str,
    output_dir: str | Path,
    lang: str = "zh",
) -> dict[str, Any]:
    """
    使用 MinerU 的 do_parse 解析 PPTX 字节流。
    返回解析结果的字典，包含 content_list、middle_json 等。
    """
    from mineru.cli.common import do_parse

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        do_parse(
            output_dir=str(output_dir),
            pdf_file_names=[filename_stem],
            pdf_bytes_list=[pptx_bytes],
            p_lang_list=[lang],
            backend="pipeline",
            parse_method="auto",
            formula_enable=True,
            table_enable=True,
            f_dump_md=True,
            f_dump_content_list=True,
            f_dump_middle_json=True,
            f_dump_model_output=False,
            f_dump_orig_pdf=False,
            f_make_md_mode="MM_MD",
            start_page_id=0,
            end_page_id=None,
        )
    except Exception as exc:
        logger.error(f"MinerU 解析失败: {exc}", exc_info=True)
        raise RuntimeError(f"MinerU 解析失败: {exc}") from exc

    # 读取输出文件
    office_dir = output_dir / filename_stem / "office"
    result: dict[str, Any] = {
        "md_content": None,
        "content_list": None,
        "content_list_v2": None,
        "middle_json": None,
    }

    md_file = office_dir / f"{filename_stem}.md"
    if md_file.exists():
        result["md_content"] = md_file.read_text(encoding="utf-8")

    cl_file = office_dir / f"{filename_stem}_content_list.json"
    if cl_file.exists():
        try:
            result["content_list"] = json.loads(cl_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            logger.warning(f"content_list.json 解析失败: {exc}")

    cl_v2_file = office_dir / f"{filename_stem}_content_list_v2.json"
    if cl_v2_file.exists():
        try:
            result["content_list_v2"] = json.loads(cl_v2_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            logger.warning(f"content_list_v2.json 解析失败: {exc}")

    middle_file = office_dir / f"{filename_stem}_middle.json"
    if middle_file.exists():
        try:
            result["middle_json"] = json.loads(middle_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            logger.warning(f"middle.json 解析失败: {exc}")

    return result


def _flatten_content_list_v2(content_list_v2: list[Any]) -> list[dict[str, Any]]:
    """
    将 MinerU 的嵌套 content_list_v2 扁平化为统一元素列表。
    外层列表索引即为 page_id（从0开始）。
    """
    flat_items: list[dict[str, Any]] = []
    if not isinstance(content_list_v2, list):
        return flat_items

    for page_id, page_items in enumerate(content_list_v2):
        if not isinstance(page_items, list):
            continue
        for item in page_items:
            if not isinstance(item, dict):
                continue
            # 注入 page_id 便于后续关联幻灯片
            item["page_id"] = page_id
            flat_items.append(item)
    return flat_items


def _extract_text_from_item(item: dict[str, Any]) -> str | None:
    """从 content_list_v2 的元素中提取文本内容。"""
    item_type = item.get("type", "")
    content = item.get("content", {})

    if item_type == "paragraph" and isinstance(content, dict):
        para_content = content.get("paragraph_content", [])
        texts = []
        for pc in para_content:
            if isinstance(pc, dict) and pc.get("type") == "text":
                texts.append(pc.get("content", ""))
        return "".join(texts) if texts else None

    if item_type == "table" and isinstance(content, dict):
        # 表格的文本可以从 html 中提取，或返回 caption
        return content.get("table_caption", "") or None

    return None


def _build_slide_tree(
    flat_items: list[dict[str, Any]],
    styled_elements: dict[int, list[dict[str, Any]]] | None = None,
) -> list[dict[str, Any]]:
    """按 page_id 组织为幻灯片树结构，可选合并样式信息。"""
    slides_map: dict[int, dict[str, Any]] = {}

    for item in flat_items:
        page_id = item.get("page_id", 0)
        if page_id not in slides_map:
            slides_map[page_id] = {
                "slide_index": int(page_id) + 1,  # 转为1-based，更直观
                "elements": [],
            }

        element: dict[str, Any] = {
            "type": item.get("type", "unknown"),
        }

        text = _extract_text_from_item(item)
        if text is not None:
            element["text"] = text

        content = item.get("content", {})
        if isinstance(content, dict):
            if "html" in content:
                element["html"] = content["html"]
            if "img_path" in content:
                element["img_path"] = content["img_path"]
            if "table_type" in content:
                element["table_type"] = content["table_type"]

        slides_map[page_id]["elements"].append(element)

    # --- 合并 PPTX 中提取的样式信息 ---
    if styled_elements:
        for slide_data in slides_map.values():
            slide_idx = slide_data["slide_index"]
            pptx_data = styled_elements.get(slide_idx, None)
            pptx_elems = pptx_data.get("elements", []) if isinstance(pptx_data, dict) else []
            
            # 添加幻灯片背景色
            if isinstance(pptx_data, dict) and pptx_data.get("background_color"):
                slide_data["background_color"] = pptx_data["background_color"]

            # 建立文本 -> PPTX 元素的索引
            pptx_by_text: dict[str, dict[str, Any]] = {}
            pptx_by_type: dict[str, list[dict[str, Any]]] = {"table": [], "image": []}
            for pe in pptx_elems:
                t = pe.get("type", "")
                txt = pe.get("text", "")
                if t == "paragraph" and txt:
                    pptx_by_text[txt] = pe
                elif t in ("table", "image"):
                    pptx_by_type[t].append(pe)

            # 对每个 MinerU 元素，尝试合并样式
            table_idx = 0
            image_idx = 0
            for elem in slide_data["elements"]:
                etype = elem.get("type", "")
                etext = elem.get("text", "")

                if etype == "paragraph" and etext:
                    # 精确匹配
                    styled = pptx_by_text.get(etext)
                    if styled is None:
                        # 模糊匹配：文本片段
                        for k, v in pptx_by_text.items():
                            if etext in k or k in etext:
                                styled = v
                                break
                    if styled:
                        if "paragraphs" in styled:
                            elem["paragraphs"] = styled["paragraphs"]
                        if "shape_style" in styled:
                            elem["shape_style"] = styled["shape_style"]

                elif etype == "table":
                    if table_idx < len(pptx_by_type["table"]):
                        styled = pptx_by_type["table"][table_idx]
                        table_idx += 1
                        if "shape_style" in styled:
                            elem["shape_style"] = styled["shape_style"]
                        if "cell_styles" in styled:
                            elem["cell_styles"] = styled["cell_styles"]
                        # 优先用 PPTX 里更准确的 HTML
                        if "html" in styled and not elem.get("html"):
                            elem["html"] = styled["html"]

                elif etype == "image":
                    if image_idx < len(pptx_by_type["image"]):
                        styled = pptx_by_type["image"][image_idx]
                        image_idx += 1
                        if "shape_style" in styled:
                            elem["shape_style"] = styled["shape_style"]

    # 按 slide_index 排序
    return [slides_map[k] for k in sorted(slides_map.keys())]


def _extract_styled_elements_from_pptx(
    pptx_path: str | Path,
    slide_count: int,
) -> dict[int, list[dict[str, Any]]]:
    """
    直接从 PPTX 提取每张幻灯片的元素，包含完整样式信息。
    
    返回: {slide_index(1-based): [styled_elements]}
    每个元素包含 type、text、paragraphs（含 runs 和 style）、shape_style 等。
    """
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(str(pptx_path))
    slide_elements: dict[int, list[dict[str, Any]]] = {}

    for slide_idx, slide in enumerate(prs.slides, start=1):
        elements: list[dict[str, Any]] = []
        # 提取幻灯片背景色
        bg_color = _extract_slide_background_color(slide)

        for shape in slide.shapes:
            # ---- 文本框 ----
            if shape.has_text_frame:
                paragraphs_data = []
                for paragraph in shape.text_frame.paragraphs:
                    p_elem = paragraph._p
                    runs_data = []
                    for run in paragraph.runs:
                        font_style = _extract_font_style(run.font, p_elem)
                        runs_data.append({
                            "text": _clean_text(run.text),
                            "style": font_style,
                        })
                    para_style = _extract_paragraph_style(paragraph, p_elem)
                    paragraphs_data.append({
                        "runs": runs_data,
                        "style": para_style,
                    })

                shape_style = _extract_shape_style(shape)
                full_text = _clean_text(shape.text_frame.text)
                if full_text:
                    elements.append({
                        "type": "paragraph",
                        "text": full_text,
                        "paragraphs": paragraphs_data,
                        "shape_style": shape_style,
                    })

            # ---- 表格 ----
            elif shape.has_table:
                table = shape.table
                shape_style = _extract_shape_style(shape)

                # 构建 HTML 表格
                html_rows = []
                for row in table.rows:
                    html_cells = []
                    for cell in row.cells:
                        html_cells.append(_clean_text(cell.text))
                    html_rows.append(html_cells)
                html_parts = ["<table>"]
                for ri, row in enumerate(html_rows):
                    tag = "th" if ri == 0 else "td"
                    html_parts.append("  <tr>")
                    for cell in row:
                        html_parts.append(f"    <{tag}>{cell}</{tag}>")
                    html_parts.append("  </tr>")
                html_parts.append("</table>")
                html_content = "\n".join(html_parts)

                # 提取单元格样式
                cell_styles = []
                for row in table.rows:
                    row_styles = []
                    for cell in row.cells:
                        cs: dict[str, Any] = {}
                        try:
                            # 从 XML 直接读取单元格背景色
                            tc = cell._tc
                            bg = _parse_solid_fill_color(tc)
                            if bg:
                                cs["background_color"] = bg
                            # 字体样式：从 defRPr 读取
                            if cell.text:
                                for p in cell.text_frame.paragraphs:
                                    pe = p._p
                                    fs = _parse_defRPr(pe)
                                    if fs:
                                        cs["font_style"] = fs
                                        break
                        except Exception:
                            pass
                        row_styles.append(cs)
                    cell_styles.append(row_styles)

                elements.append({
                    "type": "table",
                    "html": html_content,
                    "table_type": "simple_table",
                    "shape_style": shape_style,
                    "cell_styles": cell_styles,
                })

            # ---- 图片 ----
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_style = _extract_shape_style(shape)
                elements.append({
                    "type": "image",
                    "shape_style": shape_style,
                })

        slide_elements[slide_idx] = {
            "elements": elements,
            "background_color": bg_color
        }

    return slide_elements


def _extract_headings(flat_items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """提取标题元素。MinerU PPTX 输出中标题通常也是 paragraph。"""
    headings: list[dict[str, Any]] = []
    for item in flat_items:
        item_type = item.get("type", "")
        text = _extract_text_from_item(item)
        if text is None:
            continue
        # MinerU 对 PPTX 的解析结果中，标题通常没有单独标记
        # 这里简单将第一个 paragraph 视为标题（实际可根据位置/字号判断）
        if item_type in ("title", "heading", "section_header"):
            headings.append({
                "text": text,
                "level": 1,
                "slide_index": int(item.get("page_id", 0)) + 1,
            })
    return headings


def _extract_images(flat_items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """提取图片引用。"""
    images: list[dict[str, Any]] = []
    for item in flat_items:
        if item.get("type") == "image":
            content = item.get("content", {})
            img_path = content.get("img_path") if isinstance(content, dict) else None
            images.append({
                "img_path": img_path,
                "slide_index": int(item.get("page_id", 0)) + 1,
            })
    return images


def _extract_tables(flat_items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """提取表格数据。"""
    tables: list[dict[str, Any]] = []
    for item in flat_items:
        if item.get("type") == "table":
            content = item.get("content", {})
            tables.append({
                "html": content.get("html") if isinstance(content, dict) else None,
                "table_type": content.get("table_type") if isinstance(content, dict) else None,
                "slide_index": int(item.get("page_id", 0)) + 1,
            })
    return tables


def _extract_texts_flat(flat_items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """扁平化提取所有文本元素。"""
    texts: list[dict[str, Any]] = []
    for item in flat_items:
        text = _extract_text_from_item(item)
        if text is not None:
            texts.append({
                "text": text,
                "type": item.get("type", "unknown"),
                "slide_index": int(item.get("page_id", 0)) + 1,
            })
    return texts


def pptx_to_json(
    pptx_path: str | Path,
    output_json_path: str | Path | None = None,
    mineru_output_dir: str | Path | None = None,
    keep_mineru_output: bool = False,
    lang: str = "zh",
) -> dict[str, Any]:
    """
    将 PPTX 转换为结构化 JSON（基于 MinerU）。

    参数：
        pptx_path: 输入 PPTX 文件路径
        output_json_path: 输出 JSON 文件路径（可选）
        mineru_output_dir: MinerU 中间输出目录（默认使用临时目录）
        keep_mineru_output: 是否保留 MinerU 原始输出文件
        lang: 语言提示（默认 zh）

    返回：
        结构化字典，可直接 json.dumps
    """
    pptx_path = Path(pptx_path)
    filename_stem = pptx_path.stem

    # 1. 读取文件
    logger.info(f"读取 PPTX 文件: {pptx_path}")
    pptx_bytes = _read_pptx_bytes(pptx_path)

    # 2. 准备输出目录（默认放在当前目录下的 mineru_output/ 中）
    if mineru_output_dir is None:
        mineru_output_dir = Path.cwd() / "mineru_output" / filename_stem
    else:
        mineru_output_dir = Path(mineru_output_dir)

    # 3. 调用 MinerU 解析
    logger.info("开始 MinerU 解析...")
    mineru_result = _parse_with_mineru(
        pptx_bytes=pptx_bytes,
        filename_stem=filename_stem,
        output_dir=mineru_output_dir,
        lang=lang,
    )
    logger.info("MinerU 解析完成。")

    # 4. 处理 content_list_v2（嵌套列表结构）
    content_list_v2_raw = mineru_result.get("content_list_v2") or []
    flat_items = _flatten_content_list_v2(content_list_v2_raw)
    slide_count = len(content_list_v2_raw) if isinstance(content_list_v2_raw, list) else 0

    # 5. 补充提取母版/布局中的图片（MinerU 不处理这部分）
    logger.info("开始补充提取母版/布局图片...")
    layout_images, slide_image_map = _extract_layout_images(pptx_path, mineru_output_dir)

    # 合并图片：MinerU 提取的 + 补充提取的
    mineru_images = _extract_images(flat_items)
    all_images = mineru_images + layout_images

    # 6. 从 PPTX 提取样式信息（背景色、字体颜色、坐标等）
    logger.info("开始提取 PPTX 样式信息...")
    styled_elements = _extract_styled_elements_from_pptx(pptx_path, slide_count)

    # 7. 构建幻灯片树，注入图片关联和样式信息
    slides = _build_slide_tree(flat_items, styled_elements)
    for slide in slides:
        slide_idx = slide["slide_index"]
        if slide_idx in slide_image_map:
            img_indices = slide_image_map[slide_idx]
            slide["images"] = [layout_images[i] for i in img_indices if i < len(layout_images)]
        else:
            slide["images"] = []

    data: dict[str, Any] = {
        "meta": {
            "source_file": str(pptx_path.resolve()),
            "filename": pptx_path.name,
            "conversion_status": "SUCCESS",
            "slide_count": slide_count,
            "mineru_output_dir": str(mineru_output_dir.resolve()),
        },
        "headings": _extract_headings(flat_items),
        "slides": slides,
        "texts_flat": _extract_texts_flat(flat_items),
        "images": all_images,
        "tables": _extract_tables(flat_items),
        "markdown": mineru_result.get("md_content"),
        "raw_content_list": mineru_result.get("content_list"),
        "raw_content_list_v2": content_list_v2_raw,
        "raw_middle_json": mineru_result.get("middle_json"),
        "layout_images": layout_images,
        "slide_image_map": slide_image_map,
    }

    # 5. 可选保存 JSON
    if output_json_path:
        output_json_path = Path(output_json_path)
        output_json_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        logger.info(f"JSON 已保存: {output_json_path.resolve()}")

    # 6. 清理临时目录
    if not keep_mineru_output:
        try:
            shutil.rmtree(mineru_output_dir, ignore_errors=True)
            logger.info(f"已清理 MinerU 临时输出目录: {mineru_output_dir}")
        except Exception as exc:
            logger.warning(f"清理临时目录失败: {exc}")

    return data


def main() -> None:
    parser = argparse.ArgumentParser(description="PPTX 转 JSON（基于 MinerU）")
    parser.add_argument("--input", "-i", required=True, help="输入 PPTX 文件路径")
    parser.add_argument("--output", "-o", default=None, help="输出 JSON 文件路径（默认不保存文件，仅打印）")
    parser.add_argument("--mineru-dir", default=None, help="MinerU 中间输出目录（默认使用临时目录）")
    parser.add_argument("--keep-mineru-output", action="store_true", help="保留 MinerU 原始输出文件")
    parser.add_argument("--lang", default="zh", help="语言提示（默认 zh）")
    args = parser.parse_args()

    try:
        data = pptx_to_json(
            pptx_path=args.input,
            output_json_path=args.output,
            mineru_output_dir=args.mineru_dir,
            keep_mineru_output=args.keep_mineru_output,
            lang=args.lang,
        )
    except Exception as exc:
        logger.error(f"转换失败: {exc}", exc_info=True)
        sys.exit(1)

    if not args.output:
        print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
