"""
PPT转JSON核心模块
负责解析PPT文件并提取所有元素信息为结构化JSON数据
"""

import json
import os
import logging
from typing import Any

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

from utils.color_utils import rgb_color_to_hex, get_fill_color
from utils.image_utils import extract_image

logger = logging.getLogger(__name__)


# 主题颜色名称到索引的映射
THEME_COLOR_MAP = {
    "dk1": 1, "lt1": 2, "dk2": 3, "lt2": 4,
    "accent1": 5, "accent2": 6, "accent3": 7, "accent4": 8,
    "accent5": 9, "accent6": 10, "hlink": 11, "folHlink": 12
}


def extract_theme_colors(prs: Presentation) -> dict:
    """
    从PPT中提取主题颜色方案
    
    Args:
        prs: Presentation对象
        
    Returns:
        主题颜色字典，如 {"dk1": "#14141E", "lt1": "#B3B3BF", ...}
    """
    theme_colors = {}
    try:
        from lxml import etree
        master = prs.slide_masters[0]
        for rel in master.part.rels.values():
            if 'theme' in rel.reltype:
                theme_part = rel.target_part
                theme_xml = theme_part.blob
                root = etree.fromstring(theme_xml)
                for elem in root.iter():
                    if 'clrScheme' in elem.tag:
                        for child in elem:
                            tag = child.tag.split('}')[-1]
                            for sub in child:
                                sub_tag = sub.tag.split('}')[-1]
                                if sub_tag == 'srgbClr':
                                    val = sub.attrib.get('val', '')
                                    if val:
                                        theme_colors[tag] = f"#{val}"
                                elif sub_tag == 'sysClr':
                                    val = sub.attrib.get('lastClr', '')
                                    if val:
                                        theme_colors[tag] = f"#{val}"
                break
    except Exception as e:
        logger.warning(f"提取主题颜色失败: {e}")
    return theme_colors


class PPTToJsonConverter:
    """
    PPT转JSON转换器
    
    将PowerPoint文件解析为结构化的JSON数据，保留所有视觉元素信息
    """
    
    def __init__(self, pptx_path: str, output_dir: str):
        """
        初始化转换器
        
        Args:
            pptx_path: PPT文件路径
            output_dir: 输出目录
        """
        self.pptx_path = pptx_path
        self.output_dir = output_dir
        self.assets_dir = os.path.join(output_dir, "assets", "images")
        self.image_counter = 0
        self.image_map = {}  # 用于去重
        self.theme_colors = {}  # 主题颜色缓存
        
        os.makedirs(self.assets_dir, exist_ok=True)
        
    def convert(self) -> dict:
        """
        执行PPT到JSON的转换
        
        Returns:
            结构化的JSON数据字典
        """
        logger.info(f"开始解析PPT文件: {self.pptx_path}")
        
        try:
            prs = Presentation(self.pptx_path)
        except Exception as e:
            logger.error(f"无法打开PPT文件: {e}")
            raise
        
        # 提取主题颜色
        self.theme_colors = extract_theme_colors(prs)
        logger.info(f"提取到主题颜色: {self.theme_colors}")
        
        result = {
            "metadata": self._extract_metadata(prs),
            "theme_colors": self.theme_colors,
            "slides": []
        }
        
        for slide_idx, slide in enumerate(prs.slides, 1):
            logger.info(f"处理第 {slide_idx} 张幻灯片")
            slide_data = self._extract_slide(slide, slide_idx)
            result["slides"].append(slide_data)
        
        logger.info(f"PPT解析完成，共 {len(result['slides'])} 张幻灯片")
        return result
    
    def _extract_metadata(self, prs: Presentation) -> dict:
        """提取PPT元数据"""
        return {
            "file_path": self.pptx_path,
            "slide_width": prs.slide_width,
            "slide_height": prs.slide_height,
            "slide_width_inches": prs.slide_width / 914400,
            "slide_height_inches": prs.slide_height / 914400,
            "slide_count": len(prs.slides)
        }

    def _extract_slide(self, slide, slide_idx: int) -> dict:
        """提取单张幻灯片的所有信息"""
        slide_data = {
            "slide_index": slide_idx,
            "slide_id": slide.slide_id,
            "layout": self._extract_layout_info(slide.slide_layout),
            "background": self._extract_background(slide),
            "shapes": []
        }

        # 提取slide级别的形状
        for shape in slide.shapes:
            shape_data = self._extract_shape(shape)
            if shape_data:
                slide_data["shapes"].append(shape_data)

        # 提取layout级别的形状（如背景图片等）
        layout = slide.slide_layout
        for shape in layout.shapes:
            # 避免重复提取已在slide中出现的占位符
            shape_data = self._extract_shape(shape)
            if shape_data:
                shape_data["source"] = "layout"
                slide_data["shapes"].append(shape_data)

        # 提取master级别的形状（如区域背景等）
        master = slide.slide_layout.slide_master
        for shape in master.shapes:
            # 只提取auto_shape类型的master形状（如Rectangle背景区域）
            if hasattr(shape, 'shape_type') and shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                shape_data = self._extract_shape(shape)
                if shape_data:
                    shape_data["source"] = "master"
                    slide_data["shapes"].append(shape_data)

        return slide_data

    def _extract_layout_info(self, layout) -> dict:
        """提取幻灯片布局信息"""
        placeholders = []
        for placeholder in layout.placeholders:
            placeholders.append({
                "idx": placeholder.placeholder_format.idx,
                "type": str(placeholder.placeholder_format.type),
                "name": placeholder.name,
                "left": placeholder.left,
                "top": placeholder.top,
                "width": placeholder.width,
                "height": placeholder.height
            })

        return {
            "name": layout.name,
            "placeholder_count": len(layout.placeholders),
            "placeholders": placeholders
        }

    def _extract_background(self, slide) -> dict:
        """提取幻灯片背景信息（包含master继承链）"""
        bg = slide.background
        result = {
            "type": None,
            "fill": None
        }

        if bg.fill:
            result["fill"] = get_fill_color(bg.fill)
            if bg.fill.type:
                result["type"] = str(bg.fill.type)

        # 如果slide背景是BACKGROUND类型，需要提取master的实际背景色
        if bg.fill and bg.fill.type is not None:
            from pptx.enum.dml import MSO_FILL_TYPE
            if bg.fill.type == MSO_FILL_TYPE.BACKGROUND:
                master_bg = self._extract_master_background(slide)
                if master_bg:
                    result["master_background"] = master_bg

        return result

    def _extract_master_background(self, slide) -> dict | None:
        """提取slide master的背景信息"""
        try:
            master = slide.slide_layout.slide_master
            if not master or not master.background:
                return None

            fill = master.background.fill
            if not fill or fill.type is None:
                return None

            result = {"type": str(fill.type)}

            from pptx.enum.dml import MSO_FILL_TYPE
            if fill.type == MSO_FILL_TYPE.SOLID:
                try:
                    result["color"] = rgb_color_to_hex(fill.fore_color.rgb)
                except Exception:
                    pass
            elif fill.type == MSO_FILL_TYPE.GRADIENT:
                result["gradient_stops"] = []
                if hasattr(fill, 'gradient_stops'):
                    for stop in fill.gradient_stops:
                        stop_info = {
                            "position": stop.position,
                            "color": rgb_color_to_hex(stop.color.rgb) if stop.color else None
                        }
                        result["gradient_stops"].append(stop_info)

            return result
        except Exception:
            return None
    
    def _extract_shape(self, shape) -> dict | None:
        """提取形状元素信息"""
        shape_data = {
            "shape_id": shape.shape_id,
            "name": shape.name,
            "shape_type": str(shape.shape_type),
            "left": shape.left,
            "top": shape.top,
            "width": shape.width,
            "height": shape.height,
            "rotation": shape.rotation if hasattr(shape, 'rotation') else 0,
        }
        
        # 提取位置尺寸（英寸单位便于人工阅读）
        shape_data["position"] = {
            "left_inches": shape.left / 914400,
            "top_inches": shape.top / 914400,
            "width_inches": shape.width / 914400,
            "height_inches": shape.height / 914400
        }
        
        # 根据形状类型提取特定信息
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            shape_data.update(self._extract_picture(shape))
        elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
            shape_data.update(self._extract_text_frame(shape))
        elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            shape_data.update(self._extract_auto_shape(shape))
        elif shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
            shape_data.update(self._extract_placeholder(shape))
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            shape_data.update(self._extract_group(shape))
        elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            shape_data.update(self._extract_table(shape))
        elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
            shape_data.update(self._extract_chart(shape))
        elif shape.shape_type == MSO_SHAPE_TYPE.LINE:
            shape_data.update(self._extract_line(shape))
        elif shape.shape_type == MSO_SHAPE_TYPE.FREEFORM:
            shape_data.update(self._extract_freeform(shape))
        else:
            # 通用形状处理
            shape_data.update(self._extract_generic_shape(shape))
        
        return shape_data
    
    def _extract_picture(self, shape) -> dict:
        """提取图片信息"""
        result = {"element_type": "picture"}
        
        try:
            image = shape.image
            image_hash = hash(image.blob)
            
            # 检查是否已提取过相同图片
            if image_hash in self.image_map:
                result["image_info"] = self.image_map[image_hash]
            else:
                self.image_counter += 1
                image_info = extract_image(image, self.assets_dir, self.image_counter)
                self.image_map[image_hash] = image_info
                result["image_info"] = image_info
            
            # 图片裁剪信息
            if hasattr(shape, 'crop_left'):
                result["crop"] = {
                    "left": shape.crop_left,
                    "top": shape.crop_top,
                    "right": shape.crop_right,
                    "bottom": shape.crop_bottom
                }
            
        except Exception as e:
            logger.warning(f"提取图片时出错: {e}")
            result["error"] = str(e)
        
        return result
    
    def _extract_text_frame(self, shape) -> dict:
        """提取文本框信息"""
        result = {"element_type": "text_box"}
        
        if hasattr(shape, 'text_frame'):
            result["text_frame"] = self._extract_text_frame_data(shape.text_frame)
        
        if hasattr(shape, 'fill'):
            result["fill"] = get_fill_color(shape.fill)
        
        if hasattr(shape, 'line'):
            result["line"] = self._extract_line_format(shape.line)
        
        return result
    
    def _extract_text_frame_data(self, text_frame) -> dict:
        """提取文本框详细数据"""
        result = {
            "text": text_frame.text,
            "paragraphs": [],
            "margin_left": text_frame.margin_left,
            "margin_top": text_frame.margin_top,
            "margin_right": text_frame.margin_right,
            "margin_bottom": text_frame.margin_bottom,
            "word_wrap": text_frame.word_wrap,
            "auto_size": str(text_frame.auto_size) if text_frame.auto_size else None,
            "vertical_anchor": str(text_frame.vertical_anchor) if text_frame.vertical_anchor else None,
            "lst_style": self._extract_lst_style(text_frame)
        }
        
        for para in text_frame.paragraphs:
            para_data = self._extract_paragraph(para)
            result["paragraphs"].append(para_data)
        
        return result
    
    def _extract_lst_style(self, text_frame) -> dict | None:
        """提取文本框的列表样式(lstStyle)，包含默认段落/字体属性"""
        txBody = text_frame._txBody
        lstStyle = txBody.find(qn('a:lstStyle'))
        if lstStyle is None:
            return None
        
        result = {}
        for level in range(1, 10):
            lvl_pPr = lstStyle.find(qn(f'a:lvl{level}pPr'))
            if lvl_pPr is not None:
                level_data = {}
                algn = lvl_pPr.get('algn')
                if algn:
                    level_data['alignment'] = algn
                
                defTabSz = lvl_pPr.get('defTabSz')
                if defTabSz:
                    level_data['def_tab_sz'] = defTabSz
                
                lnSpc = lvl_pPr.find(qn('a:lnSpc'))
                if lnSpc is not None:
                    spcPct = lnSpc.find(qn('a:spcPct'))
                    if spcPct is not None:
                        level_data['line_spacing_pct'] = spcPct.get('val')
                
                defRPr = lvl_pPr.find(qn('a:defRPr'))
                if defRPr is not None:
                    rpr_data = {}
                    sz = defRPr.get('sz')
                    if sz:
                        rpr_data['size'] = int(sz)
                    cap = defRPr.get('cap')
                    if cap:
                        rpr_data['cap'] = cap
                    spc = defRPr.get('spc')
                    if spc:
                        rpr_data['spc'] = spc
                    
                    # 字体颜色
                    solidFill = defRPr.find(qn('a:solidFill'))
                    if solidFill is not None:
                        srgbClr = solidFill.find(qn('a:srgbClr'))
                        if srgbClr is not None:
                            rpr_data['color'] = f"#{srgbClr.get('val')}"
                        else:
                            schemeClr = solidFill.find(qn('a:schemeClr'))
                            if schemeClr is not None:
                                rpr_data['color_scheme'] = schemeClr.get('val')
                    
                    # 字体名称
                    latin = defRPr.find(qn('a:latin'))
                    if latin is not None:
                        rpr_data['font_name'] = latin.get('typeface')
                    
                    # 粗体判断：字体名包含 "Bold" 即视为粗体
                    if rpr_data.get('font_name') and 'Bold' in rpr_data['font_name']:
                        rpr_data['bold'] = True
                    
                    level_data['default_run_props'] = rpr_data
                
                result[f'level{level}'] = level_data
        
        return result if result else None
    
    def _extract_paragraph(self, paragraph) -> dict:
        """提取段落信息"""
        para_data = {
            "text": paragraph.text,
            "level": paragraph.level,
            "alignment": str(paragraph.alignment) if paragraph.alignment else None,
            "line_spacing": paragraph.line_spacing,
            "space_before": paragraph.space_before,
            "space_after": paragraph.space_after,
            "runs": []
        }
        
        for run in paragraph.runs:
            run_data = self._extract_run(run)
            para_data["runs"].append(run_data)
        
        return para_data
    
    def _extract_run(self, run) -> dict:
        """提取文本运行（run）信息"""
        font = run.font
        
        run_data = {
            "text": run.text,
            "font": {
                "name": font.name,
                "size": font.size,
                "size_pt": font.size / 12700 if font.size else None,
                "bold": font.bold,
                "italic": font.italic,
                "underline": font.underline,
                "strike": getattr(font, 'strike', None),
                "color": self._extract_font_color(font.color),
                "highlight_color": str(getattr(font, 'highlight_color', None)) if getattr(font, 'highlight_color', None) else None
            }
        }
        
        return run_data
    
    def _extract_font_color(self, font_color) -> dict:
        """提取字体颜色信息"""
        result = {
            "type": None,
            "rgb": None,
            "theme_color": None,
            "theme_color_hex": None
        }

        if font_color is None:
            return result

        try:
            if font_color.type is not None:
                result["type"] = str(font_color.type)

            # 先尝试获取rgb（对RGB类型有效）
            try:
                if font_color.rgb:
                    result["rgb"] = rgb_color_to_hex(font_color.rgb)
            except AttributeError:
                pass

            # 提取theme_color并解析为hex
            if hasattr(font_color, 'theme_color') and font_color.theme_color:
                from utils.color_utils import theme_color_to_string
                tc_str = theme_color_to_string(font_color.theme_color)
                result["theme_color"] = tc_str

                # 解析theme_color名称并查找对应的HEX值
                if tc_str and self.theme_colors:
                    # 处理格式如 "ACCENT_4 (8)" 或 "ACCENT_4"
                    tc_name = tc_str.split('(')[0].strip().lower()
                    # 将 accent_4 转换为 accent4 格式以匹配theme_colors的key
                    tc_name_normalized = tc_name.replace('_', '')
                    if tc_name in self.theme_colors:
                        result["theme_color_hex"] = self.theme_colors[tc_name]
                    elif tc_name_normalized in self.theme_colors:
                        result["theme_color_hex"] = self.theme_colors[tc_name_normalized]
        except Exception:
            pass

        return result
    
    def _extract_auto_shape(self, shape) -> dict:
        """提取自动形状信息"""
        result = {"element_type": "auto_shape"}

        # 标记是否为线条形状（prstGeom prst="line"的auto_shape需要重新分类为line）
        is_line_shape = False

        try:
            auto_shape_type_val = shape.auto_shape_type
            result["auto_shape_type"] = str(auto_shape_type_val)
        except Exception:
            # 某些auto_shape_type值（如'line'）在MSO_AUTO_SHAPE_TYPE中没有XML映射，
            # hasattr在Python 3.12+不再抑制ValueError，因此需要try/except
            # 尝试通过XML直接获取prstGeom属性值
            try:
                sp = shape._element
                # 使用'.//'搜索所有后代元素，因为prstGeom是spPr的子元素而非sp的直接子元素
                prstGeom = sp.find('.//' + qn('a:prstGeom'))
                if prstGeom is not None:
                    prst_val = prstGeom.get('prst', 'unknown')
                    result["auto_shape_type"] = prst_val
                    # 当prstGeom的prst属性为"line"时，该形状实际是线条，需要重新分类
                    if prst_val == 'line':
                        is_line_shape = True
                else:
                    result["auto_shape_type"] = "unknown"
                    # 没有prstGeom时，通过形状名称和尺寸判断是否为线条
                    shape_name = shape.name.lower() if hasattr(shape, 'name') else ''
                    if ('line' in shape_name or 'connector' in shape_name) and (shape.width <= 2 or shape.height <= 2):
                        is_line_shape = True
            except Exception:
                result["auto_shape_type"] = "unknown"
                # 异常时也尝试通过名称和尺寸判断
                try:
                    shape_name = shape.name.lower() if hasattr(shape, 'name') else ''
                    if ('line' in shape_name or 'connector' in shape_name) and (shape.width <= 2 or shape.height <= 2):
                        is_line_shape = True
                except Exception:
                    pass

        # 线条形状重新分类为element_type: "line"，并提取begin/end坐标
        if is_line_shape:
            result["element_type"] = "line"
            # 从位置信息计算线条的起止坐标
            # 水平线：height≈0，垂直线：width≈0
            begin_x = shape.left
            begin_y = shape.top
            end_x = shape.left + shape.width
            end_y = shape.top + shape.height
            result["begin"] = {"x": begin_x, "y": begin_y}
            result["end"] = {"x": end_x, "y": end_y}

        if hasattr(shape, 'text_frame') and shape.text_frame.text:
            result["text_frame"] = self._extract_text_frame_data(shape.text_frame)

        if hasattr(shape, 'fill'):
            result["fill"] = get_fill_color(shape.fill)
            # 对于使用theme_color的fill，也解析出实际HEX值
            fill_data = result["fill"]
            if fill_data and fill_data.get("theme_color") and self.theme_colors:
                tc_name = fill_data["theme_color"].split('(')[0].strip().lower()
                tc_name_normalized = tc_name.replace('_', '')
                if tc_name in self.theme_colors:
                    fill_data["theme_color_hex"] = self.theme_colors[tc_name]
                elif tc_name_normalized in self.theme_colors:
                    fill_data["theme_color_hex"] = self.theme_colors[tc_name_normalized]

        if hasattr(shape, 'line'):
            result["line"] = self._extract_line_format(shape.line)

        # adjustments属性在prstGeom为'line'等未映射类型时会抛出ValueError，
        # Python 3.12+的hasattr不再抑制ValueError，因此需要try/except
        # 线条形状不需要adjustments
        if not is_line_shape:
            try:
                result["adjustments"] = [float(adj) for adj in shape.adjustments]
            except Exception:
                pass

        return result
    
    def _extract_placeholder(self, shape) -> dict:
        """提取占位符信息"""
        result = {"element_type": "placeholder"}
        
        if hasattr(shape, 'placeholder_format'):
            result["placeholder_format"] = {
                "idx": shape.placeholder_format.idx,
                "type": str(shape.placeholder_format.type)
            }
        
        if hasattr(shape, 'text_frame') and shape.text_frame.text:
            result["text_frame"] = self._extract_text_frame_data(shape.text_frame)
        
        if hasattr(shape, 'fill'):
            result["fill"] = get_fill_color(shape.fill)
        
        return result
    
    def _extract_group(self, shape) -> dict:
        """提取组合形状信息"""
        result = {
            "element_type": "group",
            "shapes": []
        }
        
        if hasattr(shape, 'shapes'):
            for sub_shape in shape.shapes:
                sub_shape_data = self._extract_shape(sub_shape)
                if sub_shape_data:
                    result["shapes"].append(sub_shape_data)
        
        return result
    
    def _extract_table(self, shape) -> dict:
        """提取表格信息"""
        result = {"element_type": "table"}
        
        if hasattr(shape, 'table'):
            table = shape.table
            result["table"] = {
                "rows": len(table.rows),
                "columns": len(table.columns),
                "cells": []
            }
            
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    cell_fill = get_fill_color(cell.fill) if hasattr(cell, 'fill') else None
                    # 对fill的theme_color也解析hex
                    if cell_fill and cell_fill.get("theme_color") and self.theme_colors:
                        tc_name = cell_fill["theme_color"].split('(')[0].strip().lower()
                        tc_name_normalized = tc_name.replace('_', '')
                        if tc_name in self.theme_colors:
                            cell_fill["theme_color_hex"] = self.theme_colors[tc_name]
                        elif tc_name_normalized in self.theme_colors:
                            cell_fill["theme_color_hex"] = self.theme_colors[tc_name_normalized]
                    
                    # 提取单元格边框
                    borders = self._extract_cell_borders(cell) if hasattr(cell, 'borders') else None
                    
                    cell_data = {
                        "row": row_idx,
                        "column": col_idx,
                        "text": cell.text,
                        "text_frame": self._extract_text_frame_data(cell.text_frame) if hasattr(cell, 'text_frame') else None,
                        "fill": cell_fill,
                        "borders": borders
                    }
                    result["table"]["cells"].append(cell_data)
        
        return result
    
    def _extract_cell_borders(self, cell) -> dict | None:
        """提取单元格边框信息"""
        try:
            borders = {}
            for border_name in ['left', 'top', 'right', 'bottom']:
                border = getattr(cell.borders, border_name, None)
                if border:
                    border_data = {
                        "type": str(border.type) if border.type else None,
                        "width": border.width if hasattr(border, 'width') else None,
                        "color": None
                    }
                    if hasattr(border, 'color') and border.color:
                        try:
                            border_data["color"] = rgb_color_to_hex(border.color.rgb)
                        except AttributeError:
                            pass
                    borders[border_name] = border_data
            return borders
        except Exception:
            return None
    
    def _extract_chart(self, shape) -> dict:
        """提取图表信息（简化处理）"""
        result = {"element_type": "chart"}
        
        if hasattr(shape, 'chart'):
            chart = shape.chart
            result["chart_type"] = str(chart.chart_type) if hasattr(chart, 'chart_type') else None
            result["has_title"] = chart.has_title if hasattr(chart, 'has_title') else False
        
        return result
    
    def _extract_line(self, shape) -> dict:
        """提取线条信息"""
        result = {"element_type": "line"}
        
        if hasattr(shape, 'line'):
            result["line"] = self._extract_line_format(shape.line)
        
        if hasattr(shape, 'begin_x'):
            result["begin"] = {
                "x": shape.begin_x,
                "y": shape.begin_y
            }
            result["end"] = {
                "x": shape.end_x,
                "y": shape.end_y
            }
        
        return result
    
    def _extract_freeform(self, shape) -> dict:
        """提取自由形状信息"""
        result = {"element_type": "freeform"}
        
        if hasattr(shape, 'fill'):
            result["fill"] = get_fill_color(shape.fill)
        
        if hasattr(shape, 'line'):
            result["line"] = self._extract_line_format(shape.line)
        
        return result
    
    def _extract_generic_shape(self, shape) -> dict:
        """提取通用形状信息"""
        result = {"element_type": "generic"}
        
        if hasattr(shape, 'text_frame') and shape.text_frame.text:
            result["text_frame"] = self._extract_text_frame_data(shape.text_frame)
        
        if hasattr(shape, 'fill'):
            result["fill"] = get_fill_color(shape.fill)
        
        if hasattr(shape, 'line'):
            result["line"] = self._extract_line_format(shape.line)
        
        return result
    
    def _extract_line_format(self, line) -> dict:
        """提取线条格式信息"""
        result = {
            "type": None,
            "color": None,
            "width": None,
            "dash_style": None
        }
        
        if line is None:
            return result
        
        try:
            if line.fill:
                result["color"] = get_fill_color(line.fill)
            
            if line.width:
                result["width"] = line.width
                result["width_pt"] = line.width / 12700
            
            if line.dash_style:
                result["dash_style"] = str(line.dash_style)
        except Exception:
            pass
        
        return result
    
    def save_json(self, data: dict, json_path: str):
        """
        将JSON数据保存到文件
        
        Args:
            data: JSON数据字典
            json_path: 输出JSON文件路径
        """
        os.makedirs(os.path.dirname(json_path), exist_ok=True)
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"JSON数据已保存到: {json_path}")


def convert_ppt_to_json(pptx_path: str, output_dir: str, json_filename: str = "output.json") -> str:
    """
    便捷的PPT转JSON函数
    
    Args:
        pptx_path: PPT文件路径
        output_dir: 输出目录
        json_filename: JSON文件名
        
    Returns:
        生成的JSON文件路径
    """
    converter = PPTToJsonConverter(pptx_path, output_dir)
    data = converter.convert()
    
    json_path = os.path.join(output_dir, json_filename)
    converter.save_json(data, json_path)
    
    return json_path
