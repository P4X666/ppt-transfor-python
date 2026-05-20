"""
图片处理工具模块
提供PPT图片的提取、保存和加载功能
"""

import os
import hashlib
import base64
from pathlib import Path


def extract_image(image_part, output_dir: str, image_index: int) -> dict:
    """
    从PPT中提取图片并保存到指定目录
    
    Args:
        image_part: PPT图片部件对象
        output_dir: 图片输出目录
        image_index: 图片索引，用于生成文件名
        
    Returns:
        包含图片信息的字典
    """
    image_info = {
        "filename": None,
        "original_ext": None,
        "width": None,
        "height": None,
        "hash": None,
        "base64": None
    }
    
    try:
        # 获取图片二进制数据
        image_bytes = image_part.blob
        
        # 计算图片哈希值，用于去重
        image_hash = hashlib.md5(image_bytes).hexdigest()
        image_info["hash"] = image_hash
        
        # 获取图片格式
        content_type = getattr(image_part, 'content_type', '')
        ext = content_type.split('/')[-1] if content_type else 'png'
        if ext == 'jpeg':
            ext = 'jpg'
        image_info["original_ext"] = ext
        
        # 生成文件名
        filename = f"image_{image_index:04d}_{image_hash[:8]}.{ext}"
        image_info["filename"] = filename
        
        # 保存图片
        output_path = os.path.join(output_dir, filename)
        with open(output_path, 'wb') as f:
            f.write(image_bytes)
        
        # 同时保存base64编码，便于JSON嵌入
        image_info["base64"] = base64.b64encode(image_bytes).decode('utf-8')
        
        # 尝试获取图片尺寸
        try:
            from PIL import Image as PILImage
            with PILImage.open(output_path) as img:
                image_info["width"] = img.width
                image_info["height"] = img.height
        except Exception:
            pass
            
    except Exception as e:
        image_info["error"] = str(e)
    
    return image_info


def load_image_from_base64(base64_str: str, output_path: str) -> bool:
    """
    从base64字符串加载图片并保存
    
    Args:
        base64_str: base64编码的图片数据
        output_path: 输出文件路径
        
    Returns:
        是否成功保存
    """
    try:
        image_bytes = base64.b64decode(base64_str)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'wb') as f:
            f.write(image_bytes)
        return True
    except Exception:
        return False


def get_image_path_from_json(image_info: dict, assets_dir: str) -> str | None:
    """
    从JSON图片信息中获取图片的本地路径
    
    Args:
        image_info: JSON中的图片信息字典
        assets_dir: 资源目录
        
    Returns:
        图片的本地路径，如果不可用则返回None
    """
    if not image_info:
        return None
    
    # 优先使用已保存的文件
    filename = image_info.get("filename")
    if filename:
        filepath = os.path.join(assets_dir, filename)
        if os.path.exists(filepath):
            return filepath
    
    # 如果没有文件但有base64数据，则临时保存
    base64_data = image_info.get("base64")
    if base64_data and filename:
        filepath = os.path.join(assets_dir, filename)
        if load_image_from_base64(base64_data, filepath):
            return filepath
    
    return None
