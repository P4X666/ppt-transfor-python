"""
幻灯片处理插件基类

提供可扩展的插件接口，允许通过添加新插件来改变形状过滤行为。
每个插件实现 should_include_shape 方法来决定某个形状是否应被包含在还原的幻灯片中。
"""


class SlideShapeFilterPlugin:
    """
    幻灯片形状过滤插件基类

    子类重写 should_include_shape 来实现自定义过滤逻辑。
    插件通过 _plugin_registry.json 或代码注册表进行注册。
    """

    @property
    def name(self) -> str:
        """插件名称，用于日志和调试"""
        return self.__class__.__name__

    @property
    def priority(self) -> int:
        """
        插件优先级，数值越小越先执行。
        默认 100，内置核心插件应使用较小值。
        """
        return 100

    def should_include_shape(self, shape_data: dict, slide_data: dict, context: dict) -> bool:
        """
        判断形状是否应该被包含在幻灯片中

        Args:
            shape_data: 当前形状的JSON数据
            slide_data: 当前幻灯片的完整JSON数据
            context:   上下文信息，包含 slide_index、slide_count 等

        Returns:
            True 表示保留该形状，False 表示排除
        """
        return True
