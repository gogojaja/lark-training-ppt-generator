"""
BadgePrimitive - 徽章原语

支持圆形/方形徽章，常用于编号、状态标识。

== 用法 ==
    from primitives import BadgePrimitive
    
    # 圆形徽章
    badge = BadgePrimitive(text="01", shape="circle")
    
    # 方形徽章
    badge = BadgePrimitive(text="P0", shape="square", bg_color="C00000")
"""


class BadgePrimitive:
    """徽章原语"""
    
    def __init__(self, text, shape="circle", size=1.5, bg_color="2E75B6",
                 text_color="FFFFFF", font_size=12, bold=True):
        """
        初始化徽章原语
        
        Args:
            text: 徽章文本
            shape: 形状（circle/square）
            size: 尺寸（cm）
            bg_color: 背景色（6位HEX）
            text_color: 文字颜色（6位HEX）
            font_size: 字号（pt）
            bold: 是否加粗
        """
        self.text = text
        self.shape = shape
        self.size = size
        self.bg_color = bg_color
        self.text_color = text_color
        self.font_size = font_size
        self.bold = bold
    
    def to_dict(self):
        """转换为字典"""
        return {
            "type": "badge",
            "text": self.text,
            "shape": self.shape,
            "size": self.size,
            "bg_color": self.bg_color,
            "text_color": self.text_color,
            "font_size": self.font_size,
            "bold": self.bold,
        }
    
    def to_xml(self, x=0, y=0):
        """生成PPTX XML"""
        size_emu = int(self.size * 360000)
        
        # 形状
        if self.shape == "circle":
            prst = "ellipse"
        else:
            prst = "roundRect"
        
        # 字重
        bold_val = "1" if self.bold else "0"
        
        xml = f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name="badge"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{size_emu}" cy="{size_emu}"/></a:xfrm>
    <a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{self.bg_color}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="zh-CN" sz="{self.font_size * 100}" b="{bold_val}" dirty="0">
          <a:solidFill><a:srgbClr val="{self.text_color}"/></a:solidFill>
        </a:rPr>
        <a:t>{self.text}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'''
        return xml
    
    def __repr__(self):
        return f"BadgePrimitive(text='{self.text}', shape='{self.shape}')"
