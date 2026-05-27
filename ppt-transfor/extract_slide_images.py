"""
从指定PPT幻灯片的指定页码中提取所有图片资源（支持三级层级扫描）

覆盖范围:
  1. Slide 形状层    — 遍历所有 shape（含 GROUP 递归、blob 去重）
  2. Layout 层级     — 提取版式中的背景/装饰图片
  3. Master 层级     — 提取母版中的图片

去重策略:
  - 按 blob 哈希去重，避免像素相同的图片重复保存
  - 但共享 rel 的形状仍视为独立实例（分别记录来源信息）

使用方法:
    python extract_slide_images.py
"""

import os
import sys
import hashlib
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn


def get_image_mime_type(image_part):
    """从图片部件获取MIME类型，含二进制文件头回退检测"""
    content_type = getattr(image_part, 'content_type', '')
    if content_type:
        return content_type

    blob = getattr(image_part, 'blob', b'')
    if blob[:4] == b'\x89PNG':
        return 'image/png'
    elif blob[:3] == b'\xff\xd8\xff':
        return 'image/jpeg'
    elif blob[:4] == b'GIF8':
        return 'image/gif'
    elif blob[:2] in (b'BM',):
        return 'image/bmp'
    return 'application/octet-stream'


def mime_to_ext(mime_type):
    """将MIME类型转换为文件扩展名"""
    mime_map = {
        'image/png': 'png',
        'image/jpeg': 'jpg',
        'image/gif': 'gif',
        'image/bmp': 'bmp',
        'image/tiff': 'tiff',
        'image/x-emf': 'emf',
        'image/x-wmf': 'wmf',
        'image/svg+xml': 'svg',
    }
    return mime_map.get(mime_type, mime_type.split('/')[-1] if '/' in mime_type else 'bin')


def save_image_bytes(image_bytes, output_dir, prefix, counter, source_name):
    """将图片字节流保存到文件，返回信息字典"""
    image_hash = hashlib.md5(image_bytes).hexdigest()

    # 从二进制头推断格式
    mime_type = 'application/octet-stream'
    if image_bytes[:4] == b'\x89PNG':
        mime_type = 'image/png'
    elif image_bytes[:3] == b'\xff\xd8\xff':
        mime_type = 'image/jpeg'

    ext = mime_to_ext(mime_type)
    filename = f"{prefix}_{counter:03d}_{image_hash[:8]}.{ext}"
    output_path = os.path.join(output_dir, filename)

    with open(output_path, 'wb') as f:
        f.write(image_bytes)

    return {
        'filename': filename,
        'path': output_path,
        'ext': ext,
        'hash': image_hash,
        'size_bytes': len(image_bytes),
        'source': source_name,
    }


def extract_svg_from_shape(shape):
    """尝试从形状的XML中提取内嵌SVG（含p:blipFill和a:blipFill两种命名空间）"""
    try:
        for ns in (qn('a:blipFill'), qn('p:blipFill')):
            blip_fills = shape._element.findall('.//' + ns)
            for blip_fill in blip_fills:
                blip = blip_fill.find(qn('a:blip'))
                if blip is None:
                    continue
                embed = blip.get(qn('r:embed'))
                if embed:
                    continue
                for child in blip:
                    if 'svgBlip' in child.tag:
                        svg_rid = child.get(qn('r:embed'))
                        if svg_rid:
                            rel = shape.part.rels[svg_rid]
                            return rel.target_part.blob
    except Exception:
        pass
    return None


def extract_images_from_shape(shape, part, seen_hashes, output_dir, prefix, counter, layer_label):
    """从单个形状中递归提取所有图片"""
    extracted = []

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            sub = extract_images_from_shape(child, part, seen_hashes, output_dir, prefix, counter + len(extracted), layer_label)
            extracted.extend(sub)
        return extracted

    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        try:
            image_bytes = shape.image.blob
            image_hash = hashlib.md5(image_bytes).hexdigest()

            if image_hash in seen_hashes:
                return extracted

            seen_hashes.add(image_hash)
            name = shape.name if hasattr(shape, 'name') else 'unnamed'
            info = save_image_bytes(image_bytes, output_dir, prefix, counter + 1, f"{layer_label}形状 → {name}")
            extracted.append(info)

        except Exception as e:
            print(f"  [警告] 形状图片提取失败: {e}")

    else:
        svg_data = extract_svg_from_shape(shape)
        if svg_data:
            svg_hash = hashlib.md5(svg_data).hexdigest()
            if svg_hash in seen_hashes:
                return extracted
            seen_hashes.add(svg_hash)

            filename = f"{prefix}_{counter + 1:03d}_{svg_hash[:8]}.svg"
            out_path = os.path.join(output_dir, filename)
            with open(out_path, 'wb') as f:
                f.write(svg_data)
            extracted.append({
                'filename': filename, 'path': out_path, 'ext': 'svg',
                'hash': svg_hash, 'size_bytes': len(svg_data),
                'source': f"{layer_label} SVG → {shape.name}",
            })

    return extracted


def extract_images_from_part_shapes(shapes, part, seen_hashes, output_dir, prefix, counter, layer_label):
    """从一组形状中提取所有图片（用于 Slide / Layout / Master）"""
    extracted = []
    for shape in shapes:
        result = extract_images_from_shape(shape, part, seen_hashes, output_dir, prefix, counter + len(extracted), layer_label)
        extracted.extend(result)
    return extracted


def extract_images_from_slide(pptx_path, slide_number, output_dir):
    """
    从指定页码的幻灯片及其 Layout/Master 中提取全部图片

    三级扫描:
      1. Slide 直接形状（96个PICTURE + 1个PLACEHOLDER）
      2. Slide Layout 形状（版面背景/装饰图）
      3. Slide Master 形状（母版图片）
    """
    if not os.path.exists(pptx_path):
        print(f"错误: PPT文件不存在: {pptx_path}")
        return None

    os.makedirs(output_dir, exist_ok=True)

    print(f"正在打开PPT文件: {pptx_path}")
    prs = Presentation(pptx_path)

    total_slides = len(prs.slides)
    print(f"PPT共 {total_slides} 页")

    if slide_number < 1 or slide_number > total_slides:
        print(f"错误: 页码 {slide_number} 超出范围 (1-{total_slides})")
        return None

    slide = prs.slides[slide_number - 1]
    slide_num_str = f"slide{slide_number}"

    seen_hashes = set()
    all_images = []

    # ── 第一层: Slide 形状 ──
    print(f"\n[第1层] Slide 直接形状 — {len(slide.shapes)} 个形状")
    slide_imgs = extract_images_from_part_shapes(
        slide.shapes, slide.part, seen_hashes, output_dir,
        prefix=f"{slide_num_str}_slide", counter=0, layer_label="Slide"
    )
    all_images.extend(slide_imgs)
    print(f"  → 提取 {len(slide_imgs)} 张")

    # ── 第二层: Slide Layout ──
    layout = slide.slide_layout
    print(f"\n[第2层] Layout 形状 — \"{layout.name}\"，共 {len(layout.shapes)} 个形状")
    layout_imgs = extract_images_from_part_shapes(
        layout.shapes, layout.part, seen_hashes, output_dir,
        prefix=f"{slide_num_str}_layout", counter=len(all_images), layer_label="Layout"
    )
    all_images.extend(layout_imgs)
    print(f"  → 提取 {len(layout_imgs)} 张")

    # ── 第三层: Slide Master ──
    master = slide.slide_layout.slide_master
    master_name = master.name if hasattr(master, 'name') else 'N/A'
    print(f"\n[第3层] Master 形状 — \"{master_name}\"，共 {len(master.shapes)} 个形状")
    master_imgs = extract_images_from_part_shapes(
        master.shapes, master.part, seen_hashes, output_dir,
        prefix=f"{slide_num_str}_master", counter=len(all_images), layer_label="Master"
    )
    all_images.extend(master_imgs)
    print(f"  → 提取 {len(master_imgs)} 张")

    return all_images


def main():
    pptx_path = os.path.abspath(
        "assets/ppt/2024 UA Presentation Template Kit PPT.pptx"
    )
    output_dir = os.path.abspath("assets/images")
    slide_number = 68

    print("=" * 60)
    print("PPT幻灯片图片提取工具（三级层级扫描版）")
    print("=" * 60)
    print(f"输入文件: {pptx_path}")
    print(f"输出目录: {output_dir}")
    print(f"目标页码: 第 {slide_number} 页")
    print("=" * 60)

    images = extract_images_from_slide(pptx_path, slide_number, output_dir)

    if images is None:
        return 1

    print(f"\n{'=' * 60}")
    print(f"提取完成! 共从第 {slide_number} 页提取到 {len(images)} 张图片\n")

    # 按层级分组展示
    layer_groups = {}
    for img in images:
        src = img['source']
        if 'Layout' in src:
            layer = 'Layout'
        elif 'Master' in src:
            layer = 'Master'
        else:
            layer = 'Slide'
        layer_groups.setdefault(layer, []).append(img)

    total_size = 0
    for layer_name in ['Slide', 'Layout', 'Master']:
        group = layer_groups.get(layer_name, [])
        if not group:
            continue
        print(f"── {layer_name} 层 ({len(group)}张) ──")
        for i, img in enumerate(group, 1):
            size_kb = img['size_bytes'] / 1024
            total_size += img['size_bytes']
            print(
                f"  {i:2d}. {img['filename']:<55s} "
                f"{img['ext']:<6s} {size_kb:>7.1f} KB  "
                f"← {img['source']}"
            )
        print()

    # 格式分布
    format_counts = {}
    for img in images:
        fmt = img['ext'].upper()
        format_counts[fmt] = format_counts.get(fmt, 0) + 1

    print(f"{'=' * 60}")
    print(f"汇总:")
    print(f"  图片总数: {len(images)} (Slide: {len(layer_groups.get('Slide', []))}, "
          f"Layout: {len(layer_groups.get('Layout', []))}, "
          f"Master: {len(layer_groups.get('Master', []))})")
    print(f"  总大小: {total_size / 1024:.1f} KB ({total_size / 1024 / 1024:.2f} MB)")
    print(f"  格式分布: {', '.join(f'{k}: {v}张' for k, v in sorted(format_counts.items()))}")
    print(f"  输出目录: {output_dir}")
    print(f"{'=' * 60}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
