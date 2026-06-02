"""
JSON转PPT核心模块
负责读取JSON数据并使用python-pptx精确还原PPT内容
"""

import json
import os
import logging
import base64

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, MSO_CONNECTOR_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn
from lxml import etree

from utils.color_utils import hex_to_rgb_color, get_fill_color
from utils.image_utils import get_image_path_from_json
import plugins as shape_plugins

logger = logging.getLogger(__name__)


class JsonToPPTConverter:
    """
    JSON转PPT转换器
    
    读取结构化的JSON数据，精确还原为PowerPoint文件
    """
    
    def __init__(self, json_path: str, assets_dir: str):
        """
        初始化转换器
        
        Args:
            json_path: JSON文件路径
            assets_dir: 资源文件目录（包含图片等）
        """
        self.json_path = json_path
        self.assets_dir = assets_dir
        self.data = None
        
    def load_json(self) -> dict:
        """加载JSON数据"""
        try:
            with open(self.json_path, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
            logger.info(f"JSON数据加载成功: {self.json_path}")
            return self.data
        except Exception as e:
            logger.error(f"加载JSON文件失败: {e}")
            raise
    
    def convert(self, output_pptx_path: str):
        """
        执行JSON到PPT的转换
        
        Args:
            output_pptx_path: 输出PPT文件路径
        """
        if self.data is None:
            self.load_json()
        
        logger.info("开始生成PPT文件...")
        
        # 创建新的Presentation
        prs = Presentation()
        
        # 设置幻灯片尺寸
        metadata = self.data.get("metadata", {})
        if "slide_width" in metadata and "slide_height" in metadata:
            prs.slide_width = metadata["slide_width"]
            prs.slide_height = metadata["slide_height"]
        
        # 处理每一张幻灯片
        slides_data = self.data.get("slides", [])
        for slide_data in slides_data:
            self._create_slide(prs, slide_data)
        
        # 保存PPT
        os.makedirs(os.path.dirname(output_pptx_path), exist_ok=True)
        prs.save(output_pptx_path)
        logger.info(f"PPT文件已保存: {output_pptx_path}")
    
    def _create_slide(self, prs: Presentation, slide_data: dict):
        """创建单张幻灯片"""
        # 获取布局
        layout_name = slide_data.get("layout", {}).get("name", "Title Slide")
        
        # 尝试找到匹配的布局
        slide_layout = None
        for layout in prs.slide_layouts:
            if layout.name == layout_name:
                slide_layout = layout
                break
        
        # 如果没有找到匹配布局，使用第一个空白布局
        if slide_layout is None:
            # 优先使用空白布局
            for layout in prs.slide_layouts:
                if "blank" in layout.name.lower() or "空白" in layout.name:
                    slide_layout = layout
                    break
            if slide_layout is None:
                slide_layout = prs.slide_layouts[0] if prs.slide_layouts else None
        
        if slide_layout is None:
            logger.error("无法找到可用的幻灯片布局")
            return
        
        # 添加幻灯片
        slide_idx = len(prs.slides)
        slide = prs.slides.add_slide(slide_layout)
        
        # 设置背景
        self._apply_background(slide, slide_data.get("background", {}))
        
        # 计算默认字体颜色（基于背景色）
        default_font_color = self._get_default_font_color(slide_data.get("background", {}))
        
        # 收集浅色区域信息（用于判断形状是否在浅色背景上）
        light_areas = self._collect_light_areas(slide_data.get("shapes", []))
        
        # 按z-order添加形状：master(底层) → layout → slide(顶层)
        # 通过插件机制动态判断Master形状是否被Layout形状遮挡
        shapes = slide_data.get("shapes", [])
        master_shapes = [s for s in shapes if s.get("source") == "master"]
        layout_shapes = [s for s in shapes if s.get("source") == "layout"]
        slide_shapes = [s for s in shapes if not s.get("source")]
        
        # 检测slide中已有的placeholder类型，用于过滤layout中的重复placeholder
        slide_placeholder_types = set()
        for s in slide_shapes:
            if s.get("element_type") == "placeholder":
                pf = s.get("placeholder_format", {})
                ptype = str(pf.get("type", ""))
                slide_placeholder_types.add(ptype)
        
        # 默认占位符文本模式，这些文本是layout的引导文字，不应出现在最终PPT中
        default_placeholder_patterns = [
            "click to insert", "click to add", "click to enter",
            "click to edit", "click to type", "click icon to add"
        ]
        
        def _is_default_placeholder_text(text):
            """判断文本是否为默认占位符引导文字"""
            if not text:
                return False
            text_lower = text.strip().lower()
            return any(pattern in text_lower for pattern in default_placeholder_patterns)
        
        plugin_context = {
            "slide_index": slide_idx,
            "slide_count": len(prs.slides)
        }
        
        for shape_data in master_shapes:
            if shape_plugins.should_include_shape(shape_data, slide_data, plugin_context):
                self._add_shape(slide, shape_data, default_font_color, light_areas)
        for shape_data in layout_shapes:
            if shape_data.get("element_type") == "placeholder":
                pf = shape_data.get("placeholder_format", {})
                ptype = str(pf.get("type", ""))
                # 跳过layout中与slide同类型的placeholder
                if ptype in slide_placeholder_types:
                    continue
                # 跳过包含默认占位符引导文字的layout placeholder
                # 这些文字（如"Click to Insert title"）是编辑界面的引导提示，
                # 不应出现在最终PPT中
                tf = shape_data.get("text_frame", {})
                text = tf.get("text", "") if tf else ""
                if _is_default_placeholder_text(text):
                    continue
            self._add_shape(slide, shape_data, default_font_color, light_areas)
        for shape_data in slide_shapes:
            self._add_shape(slide, shape_data, default_font_color, light_areas)
    
    def _collect_light_areas(self, shapes: list) -> list:
        """收集浅色填充区域信息，用于判断文字颜色"""
        light_areas = []
        for shape_data in shapes:
            fill_data = shape_data.get("fill", {})
            if not fill_data or fill_data.get("type") is None:
                continue
            
            # 获取填充颜色
            fill_color = fill_data.get("theme_color_hex") or fill_data.get("color")
            if not fill_color:
                continue
            
            # 计算亮度
            try:
                hex_color = fill_color.lstrip("#")
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                brightness = 0.299 * r + 0.587 * g + 0.114 * b
                if brightness >= 128:
                    light_areas.append({
                        "left": shape_data.get("left", 0),
                        "top": shape_data.get("top", 0),
                        "width": shape_data.get("width", 0),
                        "height": shape_data.get("height", 0),
                        "color": fill_color
                    })
            except (ValueError, IndexError):
                pass
        
        return light_areas
    
    def _get_shape_default_color(self, shape_data: dict, slide_default: str, light_areas: list) -> str:
        """根据形状位置和浅色区域确定默认字体颜色"""
        if not light_areas:
            return slide_default
        
        shape_left = shape_data.get("left", 0)
        shape_top = shape_data.get("top", 0)
        shape_width = shape_data.get("width", 0)
        shape_height = shape_data.get("height", 0)
        shape_center_x = shape_left + shape_width / 2
        shape_center_y = shape_top + shape_height / 2
        
        for area in light_areas:
            area_right = area["left"] + area["width"]
            area_bottom = area["top"] + area["height"]
            if (area["left"] <= shape_center_x <= area_right and
                area["top"] <= shape_center_y <= area_bottom):
                return "#000000"
        
        return slide_default
    
    def _get_default_font_color(self, background_data: dict) -> str | None:
        """根据背景色确定默认字体颜色"""
        master_bg = background_data.get("master_background", {})
        fill_data = background_data.get("fill", {})
        
        # 获取背景色
        bg_color = None
        if master_bg and master_bg.get("color"):
            bg_color = master_bg.get("color")
        elif fill_data and fill_data.get("color"):
            bg_color = fill_data.get("color")
        
        if not bg_color:
            return None
        
        # 判断是否为深色背景（简单的亮度计算）
        try:
            hex_color = bg_color.lstrip("#")
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            # 亮度公式: 0.299*R + 0.587*G + 0.114*B
            brightness = 0.299 * r + 0.587 * g + 0.114 * b
            # 亮度低于128认为是深色背景，使用白色字体
            if brightness < 128:
                return "#FFFFFF"
        except (ValueError, IndexError):
            pass
        
        return None
    
    def _apply_background(self, slide, background_data: dict):
        """应用幻灯片背景（包含master背景色）"""
        fill_data = background_data.get("fill", {})
        master_bg = background_data.get("master_background", {})

        try:
            bg = slide.background
            fill = bg.fill

            # 优先使用master_background（当slide背景是BACKGROUND类型时）
            if master_bg and master_bg.get("type"):
                master_type = str(master_bg.get("type", ""))
                if "SOLID" in master_type:
                    fill.solid()
                    color = master_bg.get("color")
                    if color:
                        rgb = hex_to_rgb_color(color)
                        if rgb:
                            fill.fore_color.rgb = rgb
                    return
                elif "GRADIENT" in master_type:
                    fill.gradient()
                    gradient_stops = master_bg.get("gradient_stops", [])
                    for i, stop_data in enumerate(gradient_stops):
                        if i < len(fill.gradient_stops):
                            stop_color = stop_data.get("color")
                            if stop_color:
                                rgb = hex_to_rgb_color(stop_color)
                                if rgb:
                                    fill.gradient_stops[i].color.rgb = rgb
                    return

            # 回退到slide自身的fill
            if not fill_data or fill_data.get("type") is None:
                return

            fill_type = str(fill_data.get("type", ""))

            if "SOLID" in fill_type:
                fill.solid()
                color = fill_data.get("color")
                if color:
                    rgb = hex_to_rgb_color(color)
                    if rgb:
                        fill.fore_color.rgb = rgb
            elif "GRADIENT" in fill_type:
                fill.gradient()
                gradient_stops = fill_data.get("gradient_stops", [])
                for i, stop_data in enumerate(gradient_stops):
                    if i < len(fill.gradient_stops):
                        stop_color = stop_data.get("color")
                        if stop_color:
                            rgb = hex_to_rgb_color(stop_color)
                            if rgb:
                                fill.gradient_stops[i].color.rgb = rgb
        except Exception as e:
            logger.warning(f"应用背景时出错: {e}")
    
    def _add_shape(self, slide, shape_data: dict, default_font_color: str = None, light_areas: list = None):
        """添加形状到幻灯片"""
        element_type = shape_data.get("element_type", "generic")
        
        # 根据形状位置计算实际默认字体颜色
        shape_default_color = default_font_color
        if light_areas and element_type in ("text_box", "placeholder", "table"):
            shape_default_color = self._get_shape_default_color(shape_data, default_font_color, light_areas)
        
        try:
            if element_type == "picture":
                self._add_picture(slide, shape_data)
            elif element_type == "text_box":
                self._add_text_box(slide, shape_data, shape_default_color)
            elif element_type == "auto_shape":
                self._add_auto_shape(slide, shape_data, default_font_color)
            elif element_type == "placeholder":
                self._add_placeholder(slide, shape_data, shape_default_color)
            elif element_type == "group":
                self._add_group(slide, shape_data, shape_default_color, light_areas)
            elif element_type == "table":
                self._add_table(slide, shape_data, shape_default_color)
            elif element_type == "line":
                self._add_line(slide, shape_data)
            elif element_type == "freeform":
                self._add_freeform(slide, shape_data)
            else:
                self._add_generic_shape(slide, shape_data, default_font_color)
        except Exception as e:
            logger.warning(f"添加形状时出错 ({shape_data.get('name', 'unknown')}): {e}")
    
    def _get_position(self, shape_data: dict) -> tuple:
        """获取形状位置"""
        left = shape_data.get("left", 0)
        top = shape_data.get("top", 0)
        width = shape_data.get("width", 1000000)
        height = shape_data.get("height", 1000000)
        return left, top, width, height
    
    def _add_picture(self, slide, shape_data: dict):
        """添加图片"""
        image_info = shape_data.get("image_info", {})
        image_path = get_image_path_from_json(image_info, self.assets_dir)
        
        if not image_path or not os.path.exists(image_path):
            logger.warning(f"图片文件不存在，跳过: {image_info.get('filename')}")
            return
        
        left, top, width, height = self._get_position(shape_data)
        
        try:
            slide.shapes.add_picture(image_path, left, top, width, height)
        except Exception as e:
            logger.warning(f"添加图片失败: {e}")
    
    def _add_text_box(self, slide, shape_data: dict, default_font_color: str = None):
        """添加文本框"""
        left, top, width, height = self._get_position(shape_data)
        
        shape = slide.shapes.add_textbox(left, top, width, height)
        
        # 应用文本内容
        text_frame_data = shape_data.get("text_frame", {})
        if text_frame_data:
            self._apply_text_frame(shape.text_frame, text_frame_data, default_font_color)
        
        # 应用填充
        fill_data = shape_data.get("fill")
        if fill_data:
            self._apply_fill(shape.fill, fill_data)
        
        # 应用线条
        line_data = shape_data.get("line")
        if line_data:
            self._apply_line(shape.line, line_data)
        
        # 应用旋转
        rotation = shape_data.get("rotation", 0)
        if rotation:
            shape.rotation = rotation
    
    def _add_auto_shape(self, slide, shape_data: dict, default_font_color: str = None):
        """添加自动形状"""
        left, top, width, height = self._get_position(shape_data)
        
        # 获取形状类型
        auto_shape_type_str = shape_data.get("auto_shape_type", "RECTANGLE")
        shape_type = self._parse_shape_type(auto_shape_type_str)
        
        shape = slide.shapes.add_shape(shape_type, left, top, width, height)
        
        # 应用文本
        text_frame_data = shape_data.get("text_frame", {})
        if text_frame_data and text_frame_data.get("text"):
            self._apply_text_frame(shape.text_frame, text_frame_data, default_font_color)
        
        # 应用填充
        fill_data = shape_data.get("fill")
        if fill_data:
            self._apply_fill(shape.fill, fill_data)
        
        # 应用线条
        line_data = shape_data.get("line")
        if line_data:
            self._apply_line(shape.line, line_data)
        
        # 应用旋转
        rotation = shape_data.get("rotation", 0)
        if rotation:
            shape.rotation = rotation
        
        # 应用调整值
        adjustments = shape_data.get("adjustments", [])
        if adjustments and hasattr(shape, 'adjustments'):
            for i, adj in enumerate(adjustments):
                if i < len(shape.adjustments):
                    shape.adjustments[i] = adj
    
    def _add_placeholder(self, slide, shape_data: dict, default_font_color: str = None):
        """添加占位符（作为文本框处理）"""
        # 占位符通常在添加幻灯片时已存在，这里尝试匹配或添加新形状
        left, top, width, height = self._get_position(shape_data)
        
        text_frame_data = shape_data.get("text_frame", {})
        if text_frame_data and text_frame_data.get("text"):
            shape = slide.shapes.add_textbox(left, top, width, height)
            self._apply_text_frame(shape.text_frame, text_frame_data, default_font_color)
            
            fill_data = shape_data.get("fill")
            if fill_data:
                self._apply_fill(shape.fill, fill_data)
    
    def _add_group(self, slide, shape_data: dict, default_font_color: str = None, light_areas: list = None):
        """添加组合形状（python-pptx不支持直接创建组合，逐个添加）"""
        # python-pptx 不直接支持创建组合形状
        # 这里将子形状逐个添加
        for sub_shape_data in shape_data.get("shapes", []):
            self._add_shape(slide, sub_shape_data, default_font_color, light_areas)
    
    def _add_table(self, slide, shape_data: dict, default_font_color: str = None):
        """添加表格"""
        table_data = shape_data.get("table", {})
        if not table_data:
            return
        
        rows = table_data.get("rows", 1)
        cols = table_data.get("columns", 1)
        left, top, width, height = self._get_position(shape_data)
        
        shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = shape.table
        
        # 设置所有单元格边框为黑色
        self._set_table_borders_black(table, rows, cols)
        
        # 填充单元格内容
        cells_data = table_data.get("cells", [])
        for cell_data in cells_data:
            row = cell_data.get("row", 0)
            col = cell_data.get("column", 0)
            
            if row < rows and col < cols:
                cell = table.cell(row, col)
                text = cell_data.get("text", "")
                if text:
                    cell.text = text
                
                # 应用单元格填充（表格默认黑色背景）
                fill_data = cell_data.get("fill")
                if fill_data and fill_data.get("type"):
                    self._apply_fill(cell.fill, fill_data)
                else:
                    # 默认设置单元格背景为黑色
                    try:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(0, 0, 0)
                    except Exception:
                        pass
                
                # 应用单元格文本格式（带默认颜色）
                text_frame_data = cell_data.get("text_frame")
                if text_frame_data:
                    self._apply_text_frame(cell.text_frame, text_frame_data, default_font_color)
                
                # 应用单元格边框
                borders_data = cell_data.get("borders")
                if borders_data:
                    self._apply_cell_borders(cell, borders_data)
    
    def _set_table_borders_black(self, table, rows: int, cols: int):
        """通过XML操作设置表格所有单元格边框为黑色"""
        from lxml import etree
        
        nsmap = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }
        
        for row_idx in range(rows):
            for col_idx in range(cols):
                cell = table.cell(row_idx, col_idx)
                tc = cell._tc
                tcPr = tc.find(qn('a:tcPr'))
                if tcPr is None:
                    tcPr = etree.SubElement(tc, qn('a:tcPr'))
                
                for border_name in ['lnL', 'lnR', 'lnT', 'lnB']:
                    ln = tcPr.find(qn(f'a:{border_name}'))
                    if ln is None:
                        ln = etree.SubElement(tcPr, qn(f'a:{border_name}'))
                    ln.set('w', '0')
                    solidFill = ln.find(qn('a:solidFill'))
                    if solidFill is None:
                        solidFill = etree.SubElement(ln, qn('a:solidFill'))
                    srgbClr = solidFill.find(qn('a:srgbClr'))
                    if srgbClr is None:
                        srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
                    srgbClr.set('val', '000000')
    
    def _add_line(self, slide, shape_data: dict):
        """添加线条"""
        begin = shape_data.get("begin", {})
        end = shape_data.get("end", {})
        
        if begin and end:
            # 使用add_connector添加连接线，需要MSO_CONNECTOR_TYPE枚举
            shape = slide.shapes.add_connector(
                MSO_CONNECTOR_TYPE.STRAIGHT,
                begin.get("x", 0), begin.get("y", 0),
                end.get("x", 0), end.get("y", 0)
            )
        else:
            # 没有begin/end坐标时，从位置信息推断起止点
            left, top, width, height = self._get_position(shape_data)
            begin_x = left
            begin_y = top
            end_x = left + width
            end_y = top + height
            shape = slide.shapes.add_connector(
                MSO_CONNECTOR_TYPE.STRAIGHT,
                begin_x, begin_y, end_x, end_y
            )
        
        line_data = shape_data.get("line")
        if line_data:
            self._apply_line(shape.line, line_data)
    
    def _add_freeform(self, slide, shape_data: dict):
        """添加自由形状（使用矩形近似）"""
        left, top, width, height = self._get_position(shape_data)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        
        fill_data = shape_data.get("fill")
        if fill_data:
            self._apply_fill(shape.fill, fill_data)
        
        line_data = shape_data.get("line")
        if line_data:
            self._apply_line(shape.line, line_data)
    
    def _add_generic_shape(self, slide, shape_data: dict, default_font_color: str = None):
        """添加通用形状"""
        left, top, width, height = self._get_position(shape_data)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        
        text_frame_data = shape_data.get("text_frame", {})
        if text_frame_data and text_frame_data.get("text"):
            self._apply_text_frame(shape.text_frame, text_frame_data, default_font_color)
        
        fill_data = shape_data.get("fill")
        if fill_data:
            self._apply_fill(shape.fill, fill_data)
        
        line_data = shape_data.get("line")
        if line_data:
            self._apply_line(shape.line, line_data)
        
        rotation = shape_data.get("rotation", 0)
        if rotation:
            shape.rotation = rotation
    
    def _apply_text_frame(self, text_frame, text_frame_data: dict, default_font_color: str = None):
        """应用文本框内容"""
        paragraphs_data = text_frame_data.get("paragraphs", [])
        
        if not paragraphs_data:
            return
        
        # 清除默认段落
        text_frame.clear()
        
        # 应用边距
        margin_left = text_frame_data.get("margin_left")
        if margin_left is not None:
            text_frame.margin_left = margin_left
        margin_top = text_frame_data.get("margin_top")
        if margin_top is not None:
            text_frame.margin_top = margin_top
        margin_right = text_frame_data.get("margin_right")
        if margin_right is not None:
            text_frame.margin_right = margin_right
        margin_bottom = text_frame_data.get("margin_bottom")
        if margin_bottom is not None:
            text_frame.margin_bottom = margin_bottom
        
        # 应用自动换行
        word_wrap = text_frame_data.get("word_wrap")
        if word_wrap is not None:
            text_frame.word_wrap = word_wrap
        elif word_wrap is None:
            # 原PPT未显式设置wrap；先用python-pptx去掉"none"默认值
            text_frame.word_wrap = True
            # 再通过XML操作移除wrap属性以精确匹配原PPT（无wrap=允许换行）
            bodyPr = text_frame._txBody.find(qn('a:bodyPr'))
            if bodyPr is not None:
                bodyPr.attrib.pop('wrap', None)
        
        # 应用垂直对齐
        vertical_anchor = text_frame_data.get("vertical_anchor")
        if vertical_anchor:
            try:
                # 支持 "MIDDLE (3)" 和 "MIDDLE" 两种格式
                anchor_name = vertical_anchor.split('(')[0].strip()
                text_frame.vertical_anchor = MSO_ANCHOR[anchor_name]
            except (KeyError, AttributeError):
                pass
        
        # 应用自动缩放（normAutofit=文字适配形状 / spAutoFit=形状适配文字）
        auto_size = text_frame_data.get("auto_size")
        if auto_size:
            try:
                size_name = auto_size.split('(')[0].strip()
                text_frame.auto_size = MSO_AUTO_SIZE[size_name]
            except (KeyError, AttributeError):
                pass
        
        # 应用列表样式（lstStyle），设置默认段落/字体属性
        self._apply_lst_style(text_frame, text_frame_data.get("lst_style"))
        
        # 添加段落
        for para_idx, para_data in enumerate(paragraphs_data):
            if para_idx == 0:
                paragraph = text_frame.paragraphs[0]
            else:
                paragraph = text_frame.add_paragraph()
            
            self._apply_paragraph(paragraph, para_data, default_font_color)
    
    def _apply_lst_style(self, text_frame, lst_style_data: dict):
        """通过XML操作将lstStyle应用到文本框"""
        if not lst_style_data:
            return
        
        txBody = text_frame._txBody
        
        # 移除已有的lstStyle（避免重复）
        existing = txBody.find(qn('a:lstStyle'))
        if existing is not None:
            txBody.remove(existing)
        
        lstStyle = etree.SubElement(txBody, qn('a:lstStyle'))
        
        for level_key, level_data in lst_style_data.items():
            # level_key 格式为 "level1"，需转为 XML 缩写 "lvl1pPr"
            level_num = level_key.replace('level', '')
            tag = f'a:lvl{level_num}pPr'
            lvl_pPr = etree.SubElement(lstStyle, qn(tag))
            
            alignment = level_data.get('alignment')
            if alignment:
                lvl_pPr.set('algn', alignment)
            
            def_tab_sz = level_data.get('def_tab_sz')
            if def_tab_sz:
                lvl_pPr.set('defTabSz', def_tab_sz)
            
            line_spacing_pct = level_data.get('line_spacing_pct')
            if line_spacing_pct:
                lnSpc = etree.SubElement(lvl_pPr, qn('a:lnSpc'))
                spcPct = etree.SubElement(lnSpc, qn('a:spcPct'))
                spcPct.set('val', line_spacing_pct)
            
            defRPr_data = level_data.get('default_run_props')
            if defRPr_data:
                defRPr = etree.SubElement(lvl_pPr, qn('a:defRPr'))
                
                size = defRPr_data.get('size')
                if size:
                    defRPr.set('sz', str(size))
                
                cap = defRPr_data.get('cap')
                if cap:
                    defRPr.set('cap', cap)
                
                spc = defRPr_data.get('spc')
                if spc:
                    defRPr.set('spc', spc)
                
                bold = defRPr_data.get('bold')
                if bold:
                    defRPr.set('b', '1')
                
                color = defRPr_data.get('color')
                if color:
                    solidFill = etree.SubElement(defRPr, qn('a:solidFill'))
                    srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
                    srgbClr.set('val', color.lstrip('#'))
                
                font_name = defRPr_data.get('font_name')
                if font_name:
                    for tag in ['latin', 'ea', 'cs']:
                        elem = etree.SubElement(defRPr, qn(f'a:{tag}'))
                        elem.set('typeface', font_name)
    
    def _apply_paragraph(self, paragraph, para_data: dict, default_font_color: str = None):
        """应用段落格式"""
        # 应用层级
        level = para_data.get("level", 0)
        paragraph.level = level
        
        # 应用对齐方式
        alignment = para_data.get("alignment")
        if alignment:
            try:
                align_name = alignment.split('.')[-1]
                paragraph.alignment = PP_ALIGN[align_name]
            except (KeyError, AttributeError):
                pass
        
        # 应用行距
        line_spacing = para_data.get("line_spacing")
        if line_spacing is not None:
            paragraph.line_spacing = line_spacing
        
        # 应用段前段后间距
        space_before = para_data.get("space_before")
        if space_before is not None:
            paragraph.space_before = space_before
        space_after = para_data.get("space_after")
        if space_after is not None:
            paragraph.space_after = space_after
        
        # 添加文本runs
        runs_data = para_data.get("runs", [])
        for run_idx, run_data in enumerate(runs_data):
            if run_idx == 0:
                run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            else:
                run = paragraph.add_run()
            
            self._apply_run(run, run_data, default_font_color)
    
    def _apply_run(self, run, run_data: dict, default_color: str = None):
        """应用文本运行格式
        
        Args:
            run: 文本运行对象
            run_data: 运行数据字典
            default_color: 默认字体颜色（当颜色为继承色时使用）
        """
        text = run_data.get("text", "")
        run.text = text
        
        font_data = run_data.get("font", {})
        if not font_data:
            return
        
        font = run.font
        
        # 字体名称
        name = font_data.get("name")
        if name:
            font.name = name
        
        # 字体大小
        size = font_data.get("size")
        if size:
            font.size = Pt(size / 12700)
        
        # 粗体
        bold = font_data.get("bold")
        if bold is not None:
            font.bold = bold
        
        # 斜体
        italic = font_data.get("italic")
        if italic is not None:
            font.italic = italic
        
        # 下划线
        underline = font_data.get("underline")
        if underline is not None:
            font.underline = underline
        
        # 删除线
        strike = font_data.get("strike")
        if strike is not None:
            font.strike = strike
        
        # 字体颜色
        color_data = font_data.get("color", {})
        if color_data:
            # 优先使用theme_color_hex（主题颜色的实际RGB值）
            theme_color_hex = color_data.get("theme_color_hex")
            rgb_hex = color_data.get("rgb")
            
            color_applied = False
            if theme_color_hex:
                rgb = hex_to_rgb_color(theme_color_hex)
                if rgb:
                    font.color.rgb = rgb
                    color_applied = True
            elif rgb_hex:
                rgb = hex_to_rgb_color(rgb_hex)
                if rgb:
                    font.color.rgb = rgb
                    color_applied = True
            
            # 如果颜色为None（继承色）且提供了默认颜色，则使用默认颜色
            if not color_applied and default_color:
                rgb = hex_to_rgb_color(default_color)
                if rgb:
                    font.color.rgb = rgb
    
    def _apply_fill(self, fill, fill_data: dict):
        """应用填充样式"""
        if not fill_data:
            return

        fill_type = fill_data.get("type")
        if not fill_type:
            return

        fill_type_str = str(fill_type) if fill_type else ""

        try:
            if "SOLID" in fill_type_str:
                fill.solid()
                # 优先使用theme_color_hex
                color = fill_data.get("theme_color_hex") or fill_data.get("color")
                if color:
                    rgb = hex_to_rgb_color(color)
                    if rgb:
                        fill.fore_color.rgb = rgb
            elif "GRADIENT" in fill_type_str:
                fill.gradient()
                gradient_stops = fill_data.get("gradient_stops", [])
                for i, stop_data in enumerate(gradient_stops):
                    if i < len(fill.gradient_stops):
                        stop_color = stop_data.get("color")
                        if stop_color:
                            rgb = hex_to_rgb_color(stop_color)
                            if rgb:
                                fill.gradient_stops[i].color.rgb = rgb
            elif "BACKGROUND" in fill_type_str:
                fill.background()
        except Exception as e:
            logger.warning(f"应用填充时出错: {e}")
    
    def _apply_line(self, line, line_data: dict):
        """应用线条样式"""
        try:
            color_data = line_data.get("color", {})
            if color_data and color_data.get("color"):
                line.fill.solid()
                rgb = hex_to_rgb_color(color_data["color"])
                if rgb:
                    line.fill.fore_color.rgb = rgb
            
            width = line_data.get("width")
            if width:
                line.width = Pt(width / 12700)
            
            dash_style = line_data.get("dash_style")
            if dash_style:
                try:
                    dash_name = dash_style.split('.')[-1]
                    line.dash_style = MSO_LINE_DASH_STYLE[dash_name]
                except (KeyError, AttributeError):
                    pass
        except Exception as e:
            logger.warning(f"应用线条时出错: {e}")
    
    def _apply_cell_borders(self, cell, borders_data: dict):
        """应用单元格边框样式"""
        try:
            for border_name, border_data in borders_data.items():
                border = getattr(cell.borders, border_name, None)
                if border and border_data:
                    # 设置边框颜色
                    color = border_data.get("color")
                    if color:
                        border.color.rgb = hex_to_rgb_color(color)
                    
                    # 设置边框宽度
                    width = border_data.get("width")
                    if width:
                        border.width = Pt(width / 12700)
        except Exception as e:
            logger.warning(f"应用单元格边框时出错: {e}")
    
    def _parse_shape_type(self, shape_type_str: str) -> MSO_SHAPE:
        """解析形状类型字符串为MSO_SHAPE枚举"""
        if not shape_type_str:
            return MSO_SHAPE.RECTANGLE
        
        # 提取枚举名称：支持 "OVAL (9)" 和 "MSO_SHAPE.OVAL (9)" 等格式
        enum_name = shape_type_str.split('.')[-1]
        # 去除括号及括号内的内容，如 "OVAL (9)" → "OVAL"
        enum_name = enum_name.split('(')[0].strip()
        
        if not enum_name:
            return MSO_SHAPE.RECTANGLE
        
        # "line"类型不在MSO_SHAPE枚举中，线条应通过element_type: "line"路由到_add_line，
        # 此处作为保护性处理，返回矩形（实际不应走到这里）
        if enum_name.lower() == 'line':
            logger.warning(f"线条形状通过_parse_shape_type处理，应通过element_type路由")
            return MSO_SHAPE.RECTANGLE
        
        # "unknown"类型无法确定具体形状，返回矩形作为兜底
        if enum_name.lower() == 'unknown':
            return MSO_SHAPE.RECTANGLE
        
        try:
            return MSO_SHAPE[enum_name]
        except KeyError:
            logger.warning(f"未知的形状类型: {shape_type_str}，使用默认矩形")
            return MSO_SHAPE.RECTANGLE


def convert_json_to_ppt(json_path: str, output_pptx_path: str, assets_dir: str = None) -> str:
    """
    便捷的JSON转PPT函数
    
    Args:
        json_path: JSON文件路径
        output_pptx_path: 输出PPT文件路径
        assets_dir: 资源目录，如果为None则从JSON路径推断
        
    Returns:
        生成的PPT文件路径
    """
    if assets_dir is None:
        assets_dir = os.path.join(os.path.dirname(json_path), "assets", "images")
    
    converter = JsonToPPTConverter(json_path, assets_dir)
    converter.convert(output_pptx_path)
    
    return output_pptx_path
