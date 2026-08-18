"""
PanelPrimitive - 面板原语

支持带背景色和边框的面板容器。

== 用法 ==
    from primitives import PanelPrimitive
    
    # 基础面板
    panel = PanelPrimitive(width=10, height=5)
    
    # 带背景色
    panel = PanelPrimitive(width=10, height=5, bg_color="F8FAFC")
    
    # 带边框
    panel = PanelPrimitive(width=10, height=5, border_color="E0E0E0")
"""


class PanelPrimitive:
    """面板原语"""
    
    def __init__(self, width=10, height=5, bg_color=None, border_color=None,
                 border_width=1, border_radius=0, shadow=False):
        """
        初始化面板原语
        
        Args:
            width: 宽度（cm）
            height: 高度（cm）
            bg_color: 背景色（6位HEX）
            border_color: 边框颜色（6位HEX）
            border_width: 边框宽度（px）
            border_radius: 圆角半径（px）
            shadow: 是否显示阴影
        """
        self.width = width
        self.height = height
        self.bg_color = bg_color
        self.border_color = border_color
        self.border_width = border_width
        self.border_radius = border_radius
        self.shadow = shadow
    
    def to_dict(self):
        """转换为字典"""
        return {
            "type": "panel",
            "width": self.width,
            "height": self.height,
            "bg_color": self.bg_color,
            "border_color": self.border_color,
            "border_width": self.border_width,
            "border_radius": self.border_radius,
            "shadow": self.shadow,
        }
    
    def to_xml(self, x=0, y=0):
        """生成PPTX XML"""
        # 将cm转换为EMU
        width_emu = int(self.width * 360000)
        height_emu = int(self.height * 360000)
        
        # 背景色
        fill_xml = ""
        if self.bg_color:
            fill_xml = f'<a:solidFill><a:srgbClr val="{self.bg_color}"/></a:solidFill>'
        else:
            fill_xml = '<a:noFill/>'
        
        # 边框
        border_xml = ""
        if self.border_color:
            border_width_emu = int(self.border_width * 12700)
            border_xml = f'<a:ln w="{border_width_emu}"><a:solidFill><a:srgbClr val="{self.border_color}"/></a:solidFill></a:ln>'
        else:
            border_xml = '<a:ln><a:noFill/></a:ln>'
        
        # 圆角
        av_lst = ""
        if self.border_radius > 0:
            # 圆角半径转换为pptx格式（1/100000）
            adj = int(self.border_radius * 100000 / min(self.width, self.height) / 360000)
            av_lst = f'<a:avLst><a:gd name="adj" fval="{adj}"/></a:avLst>'
        else:
            av_lst = '<a:avLst/>'
        
        # 阴影
        effect_xml = ""
        if self.shadow:
            effect_xml = '''<a:effectLst>
  <a:outerShdw blurRad="50800" dist="38100" dir="5400000" algn="t" rotWithShape="0">
    <a:srgbClr val="000000"><a:alpha val="23000"/></a:srgbClr>
  </a:outerShdw>
</a:effectLst>'''
        
        xml = f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name="panel"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>
    <a:prstGeom prst="roundRect">{av_lst}</a:prstGeom>
    {fill_xml}
    {border_xml}
    {effect_xml}
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''
        return xml
    
    def __repr__(self):
        return f"PanelPrimitive(width={self.width}, height={self.height})"
