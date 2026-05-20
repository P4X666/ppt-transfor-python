"""
颜色处理工具模块
提供PPT颜色与RGB/十六进制颜色之间的转换功能
"""

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR


def rgb_color_to_hex(color) -> str | None:
    """
    将RGBColor对象转换为十六进制颜色字符串
    
    Args:
        color: RGBColor对象或None
        
    Returns:
        十六进制颜色字符串，如 '#FF0000'，如果颜色为None则返回None
    """
    if color is None:
        return None
    return f"#{color[0]:02X}{color[1]:02X}{color[2]:02X}"


def hex_to_rgb_color(hex_color: str) -> RGBColor | None:
    """
    将十六进制颜色字符串转换为RGBColor对象
    
    Args:
        hex_color: 十六进制颜色字符串，如 '#FF0000'
        
    Returns:
        RGBColor对象，如果输入无效则返回None
    """
    if not hex_color or hex_color == "None":
        return None
    
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return None
    
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return RGBColor(r, g, b)
    except ValueError:
        return None


def theme_color_to_string(theme_color) -> str | None:
    """
    将主题颜色枚举转换为字符串表示
    
    Args:
        theme_color: MSO_THEME_COLOR枚举值
        
    Returns:
        主题颜色名称字符串
    """
    if theme_color is None or theme_color == MSO_THEME_COLOR.NOT_THEME_COLOR:
        return None
    return str(theme_color)


def get_fill_color(fill) -> dict:
    """
    从填充对象中提取颜色信息
    
    Args:
        fill: pptx填充对象
        
    Returns:
        包含颜色信息的字典
    """
    from pptx.enum.dml import MSO_FILL_TYPE
    
    result = {
        "type": None,
        "color": None,
        "theme_color": None,
        "brightness": None
    }
    
    if fill is None:
        return result
    
    try:
        fill_type = fill.type
        if fill_type is None:
            return result
            
        result["type"] = str(fill_type)
        
        if fill_type == MSO_FILL_TYPE.SOLID:
            if hasattr(fill, 'fore_color') and fill.fore_color:
                fore_color = fill.fore_color
                if fore_color.type is not None:
                    # 尝试获取RGB颜色（RGB类型）
                    try:
                        result["color"] = rgb_color_to_hex(fore_color.rgb)
                    except AttributeError:
                        pass
                    # 获取主题颜色（SCHEME类型）
                    if hasattr(fore_color, 'theme_color') and fore_color.theme_color:
                        result["theme_color"] = theme_color_to_string(fore_color.theme_color)
                    # 获取亮度调整值
                    if hasattr(fore_color, 'brightness'):
                        result["brightness"] = fore_color.brightness
        elif fill_type == MSO_FILL_TYPE.GRADIENT:
            result["gradient_stops"] = []
            if hasattr(fill, 'gradient_stops'):
                for stop in fill.gradient_stops:
                    stop_info = {
                        "position": stop.position,
                        "color": rgb_color_to_hex(stop.color.rgb) if stop.color else None
                    }
                    result["gradient_stops"].append(stop_info)
        elif fill_type == MSO_FILL_TYPE.BACKGROUND:
            result["type"] = "BACKGROUND"
    except Exception:
        pass
    
    return result
