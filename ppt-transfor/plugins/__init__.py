"""
插件注册与管理模块

提供插件的自动发现、注册和调用机制。
内置插件按优先级排序执行，自定义插件可通过 register() 扩展。
"""

import logging

from plugins.base import SlideShapeFilterPlugin
from plugins.partition_detector import PartitionDetectorPlugin

logger = logging.getLogger(__name__)

_builtin_plugins: list[SlideShapeFilterPlugin] = []
_custom_plugins: list[SlideShapeFilterPlugin] = []
_initialized = False


def _init():
    """初始化内置插件（惰性加载）"""
    global _builtin_plugins, _initialized
    if _initialized:
        return
    _builtin_plugins = [
        PartitionDetectorPlugin(),
    ]
    _initialized = True


def register(plugin: SlideShapeFilterPlugin):
    """
    注册自定义插件

    Args:
        plugin: 实现了 SlideShapeFilterPlugin 接口的插件实例
    """
    global _custom_plugins
    if not isinstance(plugin, SlideShapeFilterPlugin):
        raise TypeError(f"插件必须继承 SlideShapeFilterPlugin，收到: {type(plugin)}")
    _custom_plugins.append(plugin)
    _custom_plugins.sort(key=lambda p: p.priority)
    logger.info(f"已注册插件: {plugin.name} (优先级: {plugin.priority})")


def unregister(plugin_name: str):
    """注销指定名称的插件"""
    global _custom_plugins
    _custom_plugins = [p for p in _custom_plugins if p.name != plugin_name]


def get_plugins() -> list[SlideShapeFilterPlugin]:
    """获取所有已注册的插件（按优先级排序）"""
    _init()
    all_plugins = _builtin_plugins + _custom_plugins
    all_plugins.sort(key=lambda p: p.priority)
    return all_plugins


def should_include_shape(shape_data: dict, slide_data: dict, context: dict) -> bool:
    """
    通过所有已注册插件判断形状是否应被包含

    所有插件都返回 True 时，形状才被包含；
    任一插件返回 False，形状被排除。

    Args:
        shape_data: 形状JSON数据
        slide_data: 幻灯片JSON数据
        context:   上下文信息 (slide_index, slide_count 等)

    Returns:
        True 表示包含该形状
    """
    for plugin in get_plugins():
        if not plugin.should_include_shape(shape_data, slide_data, context):
            logger.debug(f"插件 {plugin.name} 排除了形状 {shape_data.get('name')}")
            return False
    return True
