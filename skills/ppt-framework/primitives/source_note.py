"""
SourceNotePrimitive - 来源注释原语

支持数据来源、方法论说明等注释信息。

== 用法 ==
    from primitives import SourceNotePrimitive
    
    # 基础来源注释
    note = SourceNotePrimitive(source="数据来源：2026年Q2报告")
    
    # 带方法论说明
    note = SourceNotePrimitive(source="数据来源：央行统计", method="基于公开数据整理")
"""


class SourceNotePrimitive:
    """来源注释原语"""
    
    def __init__(self, source, method=None, width=10, font_size=10,
                 color="999999", align="left"):
        """
        初始化来源注释原语
        
        Args:
            source: 来源文本
            method: 方法论说明（可选）
            width: 宽度（cm）
            font_size: 字号（pt）
            color: 颜色（6位HEX）
            align: 对齐方式（left/center/right）
        """
        self.source = source
        self.method = method
        self.width = width
        self.font_size = font_size
        self.color = color
        self.align = align
    
    def to_dict(self):
        """转换为字典"""
        return {
            "type": "source_note",
            "source": self.source,
            "method": self.method,
            "width": self.width,
            "font_size": self.font_size,
            "color": self.color,
            "align": self.align,
        }
    
    def to_xml(self, x=0, y=0):
        """生成PPTX XML"""
        width_emu = int(self.width * 360000)
        height_emu = int(self.font_size * 12700 * 2)  # 预估高度
        
        # 构建文本
        if self.method:
            text = f"{self.source} | {self.method}"
        else:
            text = self.source
        
        # 对齐方式
        align_map = {"left": "l", "center": "ctr", "right": "r"}
        align_val = align_map.get(self.align, "l")
        
        xml = f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name="source_note"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="t"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="{align_val}"/>
      <a:r>
        <a:rPr lang="zh-CN" sz="{self.font_size * 100}" b="0" dirty="0">
          <a:solidFill><a:srgbClr val="{self.color}"/></a:solidFill>
        </a:rPr>
        <a:t>{text}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'''
        return xml
    
    def __repr__(self):
        return f"SourceNotePrimitive(source='{self.source}')"
