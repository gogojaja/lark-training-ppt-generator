"""
PPT原语组件库

基于 jingmei-ppt 方法论的L3原语层，提供基础构建块。

== 原语列表 ==
- TextPrimitive: 文本原语
- PanelPrimitive: 面板原语
- DividerPrimitive: 分隔线原语
- BadgePrimitive: 徽章原语
- SourceNotePrimitive: 来源注释原语
- StepNumberPrimitive: 步骤编号原语

== 用法 ==
    from primitives import TextPrimitive, PanelPrimitive
    
    text = TextPrimitive(text="标题", role="claim")
    panel = PanelPrimitive(width=10, height=5)
"""

from .text import TextPrimitive
from .panel import PanelPrimitive
from .divider import DividerPrimitive
from .badge import BadgePrimitive
from .source_note import SourceNotePrimitive
from .step_number import StepNumberPrimitive

__all__ = [
    "TextPrimitive",
    "PanelPrimitive",
    "DividerPrimitive",
    "BadgePrimitive",
    "SourceNotePrimitive",
    "StepNumberPrimitive",
]
