"""
从PPT指定页面提取底部区域图标资源

功能说明:
    从PPT文件的指定页面中提取位于底部区域的图标/图片元素，
    按页面编号分别保存，并生成提取结果报告。

使用方法:
    python extract_page_icons.py <pptx文件路径> <页码列表> <输出目录>
    
示例:
    python extract_page_icons.py "ppt/sample.pptx" "7,8" "output"
    python extract_page_icons.py "ppt/sample.pptx" "1,2,3" "icons"
"""

import os
import sys
import json
import hashlib
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def extract_image(image_part, output_dir: str, image_index: int) -> dict:
    """
    从PPT中提取图片并保存到指定目录，保持原始格式和质量
    
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
        "size_bytes": None
    }
    
    try:
        # 获取图片二进制数据
        image_bytes = image_part.blob
        image_info["size_bytes"] = len(image_bytes)
        
        # 计算图片哈希值，用于去重
        image_hash = hashlib.md5(image_bytes).hexdigest()
        image_info["hash"] = image_hash
        
        # 获取图片原始格式
        content_type = getattr(image_part, 'content_type', '')
        ext = content_type.split('/')[-1] if content_type else 'png'
        if ext == 'jpeg':
            ext = 'jpg'
        image_info["original_ext"] = ext
        
        # 生成文件名
        filename = f"icon_{image_index:04d}_{image_hash[:8]}.{ext}"
        image_info["filename"] = filename
        
        # 保存图片（保持原始二进制数据，无损）
        output_path = os.path.join(output_dir, filename)
        with open(output_path, 'wb') as f:
            f.write(image_bytes)
        
        # 尝试获取图片尺寸
        try:
            from PIL import Image as PILImage
            with PILImage.open(output_path) as img:
                image_info["width"] = img.width
                image_info["height"] = img.height
                image_info["format"] = img.format
                image_info["mode"] = img.mode
        except Exception:
            pass
            
    except Exception as e:
        image_info["error"] = str(e)
    
    return image_info


def is_in_bottom_area(shape, slide_height_emu: int, threshold_ratio: float = 0.6) -> bool:
    """
    判断形状是否位于页面底部区域
    
    判断逻辑:
    1. 如果形状的顶部在阈值以下（shape.top >= threshold），认为是底部元素
    2. 如果形状覆盖了整个页面高度（如全页背景图），也认为包含底部区域
    
    Args:
        shape: PPT形状对象
        slide_height_emu: 幻灯片高度（EMU单位）
        threshold_ratio: 底部区域阈值比例（默认60%，即底部40%区域）
        
    Returns:
        是否在底部区域
    """
    # 计算底部区域阈值位置
    threshold = slide_height_emu * threshold_ratio
    
    # 条件1: 形状的顶部位置在阈值以下
    if shape.top >= threshold:
        return True
    
    # 条件2: 形状覆盖了整个页面高度（如全页背景图），且覆盖底部区域
    shape_bottom = shape.top + shape.height
    if shape.top <= 0 and shape_bottom >= slide_height_emu:
        # 全页覆盖的图片，检查其是否覆盖底部区域
        return True
    
    # 条件3: 形状的底部跨越了阈值线
    if shape.top < threshold and shape_bottom > threshold:
        return True
    
    return False


def extract_page_bottom_icons(
    pptx_path: str, 
    page_number: int, 
    output_dir: str,
    bottom_threshold: float = 0.6
) -> list:
    """
    从PPT指定页面的底部区域提取图标资源
    
    Args:
        pptx_path: PPT文件路径
        page_number: 页码（从1开始）
        output_dir: 输出目录
        bottom_threshold: 底部区域阈值比例（默认60%）
        
    Returns:
        提取的图标信息列表
    """
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 打开PPT文件
    prs = Presentation(pptx_path)
    slide_height = prs.slide_height
    slide_width = prs.slide_width
    
    # 检查页码是否有效
    total_slides = len(prs.slides)
    if page_number < 1 or page_number > total_slides:
        raise ValueError(f"页码 {page_number} 超出范围，PPT共有 {total_slides} 页")
    
    # 获取指定页面（索引从0开始）
    slide_index = page_number - 1
    slide = prs.slides[slide_index]
    
    extracted_icons = []
    image_counter = 0
    image_hash_map = {}  # 用于去重
    
    # 底部区域阈值（EMU单位）
    threshold_emu = slide_height * bottom_threshold
    
    def extract_shapes(shapes, parent_name=""):
        nonlocal image_counter
        
        for shape in shapes:
            # 构建完整的形状名称
            shape_name = f"{parent_name}/{shape.name}" if parent_name else shape.name
            
            # 处理图片形状
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    # 检查是否在底部区域
                    if not is_in_bottom_area(shape, slide_height, bottom_threshold):
                        continue
                    
                    image = shape.image
                    image_hash = hash(image.blob)
                    
                    # 检查是否已提取过相同图片
                    if image_hash not in image_hash_map:
                        image_counter += 1
                        image_info = extract_image(image, output_dir, image_counter)
                        image_info["shape_name"] = shape_name
                        image_info["position"] = {
                            "left_emu": shape.left,
                            "top_emu": shape.top,
                            "width_emu": shape.width,
                            "height_emu": shape.height,
                            "left_inches": shape.left / 914400,
                            "top_inches": shape.top / 914400,
                            "width_inches": shape.width / 914400,
                            "height_inches": shape.height / 914400
                        }
                        image_info["source"] = parent_name if parent_name else "slide"
                        image_hash_map[image_hash] = image_info
                        extracted_icons.append(image_info)
                        
                        width = image_info.get('width', '?')
                        height = image_info.get('height', '?')
                        size_kb = image_info.get('size_bytes', 0) / 1024
                        print(f"  ✓ 提取图标: {image_info['filename']} ({width}x{height} px, {size_kb:.1f} KB)")
                    else:
                        print(f"  ⊘ 跳过重复图标: {shape_name}")
                except Exception as e:
                    print(f"  ✗ 提取图片失败 {shape_name}: {e}")
            
            # 处理组合形状（递归）
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                if hasattr(shape, 'shapes'):
                    extract_shapes(shape.shapes, shape_name)
    
    print(f"\n开始提取第 {page_number} 页的底部区域图标...")
    print(f"  页面尺寸: {slide_width/914400:.2f} x {slide_height/914400:.2f} 英寸")
    print(f"  底部区域阈值: 从 {threshold_emu/914400:.2f} 英寸处开始（页面高度的{bottom_threshold*100:.0f}%）")
    print(f"  输出目录: {output_dir}")
    
    # 提取slide级别的形状
    extract_shapes(slide.shapes)
    
    # 提取layout级别的形状
    layout = slide.slide_layout
    extract_shapes(layout.shapes, "layout")
    
    # 提取master级别的形状
    master = slide.slide_layout.slide_master
    extract_shapes(master.shapes, "master")
    
    print(f"  提取完成！共提取 {len(extracted_icons)} 个图标")
    return extracted_icons


def generate_report(results: dict, report_path: str):
    """
    生成提取结果报告
    
    Args:
        results: 提取结果字典
        report_path: 报告输出路径
    """
    report = {
        "report_title": "PPT底部区域图标提取报告",
        "generated_at": datetime.now().isoformat(),
        "source_file": results["source_file"],
        "total_pages_processed": len(results["pages"]),
        "total_icons_extracted": sum(len(p["icons"]) for p in results["pages"]),
        "output_base_dir": results["output_base_dir"],
        "pages": []
    }
    
    for page_result in results["pages"]:
        page_report = {
            "page_number": page_result["page_number"],
            "output_dir": page_result["output_dir"],
            "icons_count": len(page_result["icons"]),
            "icons": []
        }
        
        for icon in page_result["icons"]:
            icon_report = {
                "filename": icon.get("filename"),
                "format": icon.get("original_ext"),
                "dimensions": f"{icon.get('width', '?')}x{icon.get('height', '?')} px",
                "size_kb": round(icon.get("size_bytes", 0) / 1024, 2) if icon.get("size_bytes") else None,
                "hash": icon.get("hash"),
                "position": icon.get("position"),
                "source": icon.get("source"),
                "shape_name": icon.get("shape_name")
            }
            page_report["icons"].append(icon_report)
        
        report["pages"].append(page_report)
    
    # 保存JSON报告
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    
    # 生成文本报告
    text_report_path = report_path.replace('.json', '.txt')
    with open(text_report_path, 'w', encoding='utf-8') as f:
        f.write("=" * 70 + "\n")
        f.write("           PPT底部区域图标提取报告\n")
        f.write("=" * 70 + "\n\n")
        f.write(f"生成时间: {report['generated_at']}\n")
        f.write(f"源文件: {report['source_file']}\n")
        f.write(f"输出目录: {report['output_base_dir']}\n")
        f.write(f"处理页数: {report['total_pages_processed']}\n")
        f.write(f"提取图标总数: {report['total_icons_extracted']}\n")
        f.write("\n" + "-" * 70 + "\n\n")
        
        for page in report["pages"]:
            f.write(f"【第 {page['page_number']} 页】\n")
            f.write(f"  输出目录: {page['output_dir']}\n")
            f.write(f"  提取图标数: {page['icons_count']}\n")
            f.write(f"  图标列表:\n")
            
            for i, icon in enumerate(page["icons"], 1):
                f.write(f"    {i}. {icon['filename']}\n")
                f.write(f"       格式: {icon['format']}, 尺寸: {icon['dimensions']}, 大小: {icon['size_kb']} KB\n")
                if icon['position']:
                    pos = icon['position']
                    f.write(f"       位置: ({pos.get('left_inches', '?'):.2f}, {pos.get('top_inches', '?'):.2f}) 英寸\n")
                f.write(f"       来源: {icon['source']}\n")
            
            f.write("\n")
        
        f.write("=" * 70 + "\n")
        f.write("提取完成！\n")
    
    return report_path, text_report_path


def main():
    if len(sys.argv) != 4:
        print("用法: python extract_page_icons.py <pptx文件路径> <页码列表> <输出目录>")
        print("示例: python extract_page_icons.py \"ppt/sample.pptx\" \"7,8\" \"output\"")
        print("      python extract_page_icons.py \"ppt/sample.pptx\" \"1,2,3\" \"icons\"")
        sys.exit(1)
    
    pptx_path = sys.argv[1]
    page_numbers_str = sys.argv[2]
    output_base_dir = sys.argv[3]
    
    # 解析页码列表
    try:
        page_numbers = [int(p.strip()) for p in page_numbers_str.split(',')]
    except ValueError:
        print("错误: 页码格式不正确，请使用逗号分隔的数字，如 '7,8'")
        sys.exit(1)
    
    # 转换为绝对路径
    pptx_path = os.path.abspath(pptx_path)
    output_base_dir = os.path.abspath(output_base_dir)
    
    print("=" * 70)
    print("           PPT底部区域图标提取工具")
    print("=" * 70)
    print(f"PPT文件: {pptx_path}")
    print(f"目标页码: {', '.join(map(str, page_numbers))}")
    print(f"输出目录: {output_base_dir}")
    print("-" * 70)
    
    # 验证PPT文件
    if not os.path.exists(pptx_path):
        print(f"错误: PPT文件不存在: {pptx_path}")
        sys.exit(1)
    
    # 创建输出目录
    os.makedirs(output_base_dir, exist_ok=True)
    
    # 存储所有结果
    all_results = {
        "source_file": pptx_path,
        "output_base_dir": output_base_dir,
        "pages": []
    }
    
    try:
        for page_number in page_numbers:
            # 为每个页面创建单独的子目录
            page_output_dir = os.path.join(output_base_dir, f"page_{page_number:02d}")
            os.makedirs(page_output_dir, exist_ok=True)
            
            # 提取图标
            extracted_icons = extract_page_bottom_icons(
                pptx_path, 
                page_number, 
                page_output_dir,
                bottom_threshold=0.6  # 底部40%区域
            )
            
            all_results["pages"].append({
                "page_number": page_number,
                "output_dir": page_output_dir,
                "icons": extracted_icons
            })
        
        # 生成报告
        report_path = os.path.join(output_base_dir, "extraction_report.json")
        json_path, txt_path = generate_report(all_results, report_path)
        
        # 打印总结
        print("\n" + "=" * 70)
        print("           提取完成总结")
        print("=" * 70)
        total_icons = sum(len(p["icons"]) for p in all_results["pages"])
        print(f"处理页数: {len(all_results['pages'])}")
        print(f"提取图标总数: {total_icons}")
        print("\n各页面详情:")
        for page in all_results["pages"]:
            print(f"  第 {page['page_number']} 页: {len(page['icons'])} 个图标 -> {page['output_dir']}")
        print(f"\n报告文件:")
        print(f"  JSON报告: {json_path}")
        print(f"  文本报告: {txt_path}")
        print("=" * 70)
        
    except Exception as e:
        print(f"\n提取失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()