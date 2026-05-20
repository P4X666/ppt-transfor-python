#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPTX 转换服务

封装基于 python-pptx + docling-core 的 PPTX 文档转换功能
借鉴 translate_ppt2json 的优秀设计：
- 样式提取（字体、形状）
- URL远程文件支持
- 安全验证机制
- 递归群组形状处理
"""

from __future__ import annotations

import io
import json
import hashlib
import tempfile
import os
from pathlib import Path
from typing import Any, Tuple, Optional

import requests

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

# 安全配置
ALLOW_LOCAL_PATH = os.getenv("PPTX_CONVERTER_ALLOW_LOCAL", "1") == "1"
ALLOW_URL_DOWNLOAD = os.getenv("PPTX_CONVERTER_ALLOW_URL", "0") == "1"


def _safe_color_hex(font) -> Optional[str]:
    """安全获取颜色的十六进制值"""
    try:
        color = getattr(font, "color", None)
        if color is None:
            return None
        rgb = getattr(color, "rgb", None)
        if rgb is None:
            return None
        tup = tuple(int(b) for b in rgb)
        if len(tup) >= 3:
            return "".join(f"{b:02X}" for b in tup[:3])
    except Exception:
        pass
    return None


def _extract_runs_from_text_frame(text_frame) -> list[list[dict[str, Any]]]:
    """从文本框中提取所有run信息，包括文本内容和样式"""
    paras = []
    try:
        for p in text_frame.paragraphs:
            runs = []
            for r in p.runs:
                font = r.font
                runs.append({
                    "text": r.text,
                    "font_name": getattr(font, "name", None),
                    "font_size_pt": getattr(font, "size", None).pt if getattr(font, "size", None) else None,
                    "bold": getattr(font, "bold", None),
                    "italic": getattr(font, "italic", None),
                    "color": _safe_color_hex(font)
                })
            paras.append(runs)
    except Exception:
        pass
    return paras


def _extract_shape_style(shape) -> dict[str, Any]:
    """提取形状的样式属性"""
    style = {}
    from pptx.util import Pt

    # 填充样式
    try:
        fill = shape.fill
        if fill.type != 0:
            style["fill_type"] = fill.type
            if hasattr(fill, "fore_color"):
                fore_color = fill.fore_color
                if hasattr(fore_color, "rgb"):
                    style["fill_color"] = "".join(f"{b:02X}" for b in fore_color.rgb)
    except Exception:
        pass

    # 线条样式
    try:
        line = shape.line
        if line.width > Pt(0):
            style["line_width"] = line.width.pt
            if hasattr(line, "color") and hasattr(line.color, "rgb"):
                style["line_color"] = "".join(f"{b:02X}" for b in line.color.rgb)
    except Exception:
        pass

    return style


def _validate_public_http_url(url: str) -> str:
    """验证URL是否为安全的公共HTTP(S) URL"""
    import re
    
    if not (url.startswith("http://") or url.startswith("https://")):
        raise ValueError(f"不支持的协议类型: {url}")

    local_patterns = [
        r"^http://127\.",
        r"^http://localhost",
        r"^http://0\.0\.0\.0",
        r"^http://192\.168\.",
        r"^http://10\.",
        r"^http://172\.(1[6-9]|2[0-9]|3[0-1])\.",
        r"^http://\[::1\]",
    ]

    for pattern in local_patterns:
        if re.match(pattern, url, re.IGNORECASE):
            raise ValueError(f"禁止访问本地地址: {url}")

    return url


def _download_ppt_from_url(url: str) -> str:
    """从URL下载PPT文件到临时文件"""
    safe_url = _validate_public_http_url(url)
    response = requests.get(safe_url, stream=True, timeout=60, allow_redirects=False)
    response.raise_for_status()

    with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as temp_file:
        for chunk in response.iter_content(chunk_size=8192):
            if chunk:
                temp_file.write(chunk)
        return temp_file.name


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
    
    # 移除垂直制表符和其他控制字符
    cleaned = text.replace('\u000b', ' ')  # 垂直制表符替换为空格
    cleaned = cleaned.replace('\r', '')     # 移除回车符
    
    # 移除其他控制字符（保留基本的\n和\t）
    result = []
    for char in cleaned:
        if char == '\n' or char == '\t':
            result.append(char)
        elif ord(char) < 32:
            result.append(' ')
        else:
            result.append(char)
    
    return ''.join(result).strip()


def _table_to_html(table) -> str:
    """将PPT表格转换为HTML格式"""
    rows = []
    for row in table.rows:
        row_data = []
        for cell in row.cells:
            text = _clean_text(cell.text)
            row_data.append(text)
        rows.append(row_data)
    
    # 判断是否为表头（首行加粗等）
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
) -> Tuple[list[dict[str, Any]], dict[int, list[int]]]:
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
            except Exception:
                pass
        return added_indices

    master_image_indices: dict[int, list[int]] = {}
    for m_idx, master in enumerate(prs.slide_masters):
        master_image_indices[m_idx] = _process_shape_container(
            master, f"master:{master.name}"
        )

    layout_image_indices: dict[int, list[int]] = {}
    for l_idx, layout in enumerate(prs.slide_layouts):
        layout_image_indices[l_idx] = _process_shape_container(
            layout, f"layout:{layout.name}"
        )

    slide_image_map: dict[int, list[int]] = {}
    master_list = list(prs.slide_masters)
    layout_list = list(prs.slide_layouts)

    for slide_idx, slide in enumerate(prs.slides, start=1):
        associated: set[int] = set()

        own_indices = _process_shape_container(slide, f"slide:{slide_idx}")
        associated.update(own_indices)

        master = slide.slide_layout.slide_master
        m_idx = next((i for i, m in enumerate(master_list) if m is master), -1)
        if m_idx >= 0:
            associated.update(master_image_indices.get(m_idx, []))

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
    """
    slides: list[dict[str, Any]] = []
    current_slide: dict[str, Any] | None = None
    current_slide_no = 0

    for item, level in doc.iterate_items():
        provs = getattr(item, "prov", []) or []
        page_no = provs[0].page_no if provs else current_slide_no

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
        
        # 处理表格的HTML内容
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
    扁平化提取所有文本元素，保留标题层级信息。
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
        
        # 确定类型
        item_type = type(item).__name__
        elem_type = "paragraph" if item_type in ("TextItem", "TitleItem") else item_type.lower()
        
        texts.append({
            "text": _clean_text(text),
            "type": elem_type,
            "level": level,
            "label": label_name,
            "slide_index": int(page_no) if page_no is not None else None,
        })
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
        
        # 获取表格HTML内容
        if provs:
            prov = provs[0]
            if hasattr(prov, "extra") and prov.extra and "html" in prov.extra:
                table_data["html"] = prov.extra["html"]
                table_data["table_type"] = "simple_table"
        
        tables.append(table_data)
    
    return tables


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
            headings.append({
                "text": text,
                "level": level,
                "slide_index": int(page_no) if page_no is not None else None,
            })
    return headings


def _convert_pptx_with_docling_core(pptx_path: str | Path) -> DoclingDocument:
    """
    使用 docling_core 的能力将 PPTX 转为 DoclingDocument。
    """
    from pptx import Presentation

    prs = Presentation(str(pptx_path))
    doc = DoclingDocument(name=Path(pptx_path).stem)

    for slide_idx, slide in enumerate(prs.slides, start=1):
        try:
            width = prs.slide_width
            height = prs.slide_height
        except Exception:
            width = height = 0
        doc.add_page(page_no=slide_idx, size=Size(width=width, height=height))

        slide_group = GroupItem(
            name=f"slide_{slide_idx}",
            self_ref=f"#/groups/{len(doc.groups)}",
        )
        doc.groups.append(slide_group)
        doc.body.children.append(RefItem(cref=slide_group.self_ref))

        for shape in slide.shapes:
            if shape.has_text_frame:
                text = _clean_text(shape.text_frame.text)
                if not text:
                    continue
                is_title = False
                try:
                    if shape.text_frame.paragraphs:
                        font_size = shape.text_frame.paragraphs[0].font.size
                        if font_size and font_size.pt > 24:
                            is_title = True
                except Exception:
                    pass

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

            if shape.shape_type == 13:
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
                    # 将HTML内容存储在prov的额外属性中
                    prov.extra = {"html": html_content}
                    doc.add_table(data=table_data, prov=prov, parent=slide_group)
                except Exception:
                    pass

    return doc


def convert_pptx_to_json(
    pptx_path: str | Path,
    extract_images: bool = True,
    images_output_dir: Optional[str | Path] = None,
) -> dict[str, Any]:
    """
    将 PPTX 转换为结构化 JSON。

    参数：
        pptx_path: 输入 PPTX 文件路径或URL
        extract_images: 是否提取图片到本地
        images_output_dir: 图片输出目录

    返回：
        结构化字典，可直接 json.dumps
    """
    temp_file_path = None
    
    try:
        source_str = str(pptx_path)
        
        # 判断是URL还是本地路径
        if source_str.startswith('http://') or source_str.startswith('https://'):
            if not ALLOW_URL_DOWNLOAD:
                raise ValueError("未启用URL下载功能，请设置环境变量 PPTX_CONVERTER_ALLOW_URL=1")
            temp_file_path = _download_ppt_from_url(source_str)
            pptx_path = Path(temp_file_path)
        else:
            if not ALLOW_LOCAL_PATH:
                raise ValueError("未启用本地文件访问功能，请设置环境变量 PPTX_CONVERTER_ALLOW_LOCAL=1")
            pptx_path = Path(source_str)
            if not pptx_path.exists():
                raise FileNotFoundError(f"PPTX 文件不存在: {pptx_path}")

        doc = _convert_pptx_with_docling_core(pptx_path)

        images: list[dict[str, Any]] = []
        slide_image_map: dict[int, list[int]] = {}
        if extract_images:
            if images_output_dir is None:
                images_output_dir = pptx_path.parent / "images"
            images, slide_image_map = _extract_images_by_slide(pptx_path, images_output_dir)

        slides = _build_slide_tree(doc)
        for slide in slides:
            slide_idx = slide["slide_index"]
            if slide_idx in slide_image_map:
                img_indices = slide_image_map[slide_idx]
                slide["images"] = [images[i] for i in img_indices if i < len(images)]
            else:
                slide["images"] = []

        data: dict[str, Any] = {
            "meta": {
                "source_file": str(pptx_path.resolve()),
                "filename": pptx_path.name,
                "conversion_status": "SUCCESS",
                "slide_count": len(doc.pages),
                "is_remote_source": temp_file_path is not None,
            },
            "headings": _extract_headings(doc),
            "slides": slides,
            "texts_flat": _extract_text_flat(doc),
            "tables": _extract_tables(doc),
            "images": images,
            "slide_image_map": slide_image_map,
        }

        return data
    
    finally:
        # 清理临时文件
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.unlink(temp_file_path)
            except Exception:
                pass