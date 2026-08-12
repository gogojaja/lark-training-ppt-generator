"""
DividerPrimitive - 分隔线原语

支持水平和垂直分隔线。

== 用法 ==
    from primitives import DividerPrimitive
    
    # 水平分隔线
    divider = DividerPrimitive(width=10, orientation="horizontal")
    
    # 垂直分隔线
    divider = DividerPrimitive(height=5, orientation="vertical")
"""


class DividerPrimitive:
    """分隔线原语"""
    
    def __init__(self, width=None, height=None, orientation="horizontal",
                 color="E0E0E0", thickness=1):
        """
        初始化分隔线原语
        
        Args:
            width: 宽度（cm），水平分隔线必须指定
            height: 高度（cm），垂直分隔线必须指定
            orientation: 方向（horizontal/vertical）
            color: 颜色（6位HEX）
            thickness: 厚度（px）
        """
        self.orientation = orientation
        self.color = color
        self.thickness = thickness
        
        if orientation == "horizontal":
            self.width = width or 10
            self.height = thickness * 0.01  # 转换为cm
        else:
            self.width = thickness * 0.01
            self.height = height or 5
    
    def to_dict(self):
        """转换为字典"""
        return {
            "type": "divider",
            "width": self.width,
            "height": self.height,
            "orientation": self.orientation,
            "color": self.color,
            "thickness": self.thickness,
        }
    
    def to_xml(self, x=0, y=0):
        """生成PPTX XML"""
        width_emu = int(self.width * 360000)
        height_emu = int(self.height * 360000)
        thickness_emu = int(self.thickness * 12700)
        
        xml = f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name="divider"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{self.color}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>'''
        return xml
    
    def __repr__(self):
        return f"DividerPrimitive(orientation='{self.orientation}')"
