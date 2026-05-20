"""
分区检测插件

通过分析幻灯片中内容形状的布局位置、尺寸和类型，自动识别Master级别的
区域背景形状是否真正构成可见分区。

核心逻辑：若幻灯片中存在大型内容形状（如表格）明确位于Rectangle右侧，
且与Rectangle边缘留有间隔，则Rectangle构成活跃分区。
小型标签、页脚等不应触发分区识别。
"""

from plugins.base import SlideShapeFilterPlugin


class PartitionDetectorPlugin(SlideShapeFilterPlugin):
    """
    分区检测插件

    检测逻辑：
    1. 对于 source="master" 的 auto_shape（Rectangle），计算其包围盒
    2. 扫描幻灯片中非Master的内容形状，按以下规则判断：
       a. 内容形状左边界需在Rectangle右侧阈值之外（右边界+5%宽度，保证有间隔）
       b. 仅检查 table 和 text_box 类型（排除 placeholder、picture 等非内容元素）
       c. 形状宽度≥阈值（max(Rectangle宽度×40%, 2000000 EMU)），排除小型标签
        d. 内容形状垂直范围需与Rectangle有交集
    3. 满足条件的形状表明页面围绕Rectangle布局 → Rectangle为活跃分区 → 保留
    4. 否则 Rectangle 被布局遮盖 → 排除

    适用范围：
    - 左侧浅色背景 + 右侧表格内容的分区页面
    - 自动适应不同布局、不同形状组合，无需指定页面索引
    """

    @property
    def priority(self) -> int:
        return 10

    def should_include_shape(self, shape_data: dict, slide_data: dict, context: dict) -> bool:
        source = shape_data.get("source")
        if source != "master":
            return True

        fill_data = shape_data.get("fill", {})
        if not fill_data or not fill_data.get("type"):
            return True

        element_type = shape_data.get("element_type")
        if element_type != "auto_shape":
            return True

        if not self._is_active_partition(shape_data, slide_data):
            return False

        return True

    def _is_active_partition(self, master_shape: dict, slide_data: dict) -> bool:
        """
        判断Master Rectangle是否构成活跃分区

        活跃分区条件：
        1. 幻灯片中存在非Master的 table 或 text_box 形状
        2. 形状左边界在Rectangle右侧且留有5%宽度的间隔
        3. 形状宽度足够大（≥max(Rectangle宽×40%, 2000000 EMU)），排除小型标签
        4. 形状垂直范围与Rectangle有交集
        """
        rect_left = master_shape.get("left", 0)
        rect_top = master_shape.get("top", 0)
        rect_width = master_shape.get("width", 0)
        rect_height = master_shape.get("height", 0)

        if rect_width <= 0 or rect_height <= 0:
            return False

        rect_right = rect_left + rect_width
        rect_bottom = rect_top + rect_height
        right_threshold = rect_right + rect_width * 0.05
        # 最小宽度：Rectangle宽度的40%且不小于2000000 EMU，排除小型标签和页脚
        min_width = max(rect_width * 0.4, 2000000)

        for shape in slide_data.get("shapes", []):
            shape_source = shape.get("source", "")
            if shape_source == "master":
                continue

            element_type = shape.get("element_type", "")
            if element_type not in ("table", "text_box"):
                continue

            shape_left = shape.get("left", 0)
            shape_top = shape.get("top", 0)
            shape_width = shape.get("width", 0)
            shape_height = shape.get("height", 0)

            if shape_width <= 0 or shape_height <= 0:
                continue

            # 排除小型形状（页脚标签、logo等），它们不构成分区的证据
            if shape_width < min_width:
                continue

            shape_bottom = shape_top + shape_height

            # 形状左边界在Rectangle右侧阈值之外 → 围绕Rectangle布局
            if shape_left > right_threshold:
                if shape_bottom > rect_top and shape_top < rect_bottom:
                    return True

        return False
