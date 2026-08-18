"""
StepNumberPrimitive - 步骤编号原语

支持带编号的圆形/方形步骤标识。

== 用法 ==
    from primitives import StepNumberPrimitive
    
    # 圆形步骤编号
    step = StepNumberPrimitive(number=1, size=1.5)
    
    # 方形步骤编号
    step = StepNumberPrimitive(number=2, shape="square", bg_color="C00000")
"""


class StepNumberPrimitive:
    """步骤编号原语"""
    
    def __init__(self, number, shape="circle", size=1.5, bg_color="2E75B6",
                 text_color="FFFFFF", font_size=14, bold=True, prefix=""):
        """
        初始化步骤编号原语
        
        Args:
            number: 步骤编号
            shape: 形状（circle/rounded_rect）
            size: 尺寸（cm）
            bg_color: 背景色（6位HEX）
            text_color: 文字颜色（6位HEX）
            font_size: 字号（pt）
            bold: 是否加粗
            prefix: 编号前缀（如"步骤"）
        """
        self.number = number
        self.shape = shape
        self.size = size
        self.bg_color = bg_color
        self.text_color = text_color
        self.font_size = font_size
        self.bold = bold
        self.prefix = prefix
    
    def to_dict(self):
        """转换为字典"""
        return {
            "type": "step_number",
            "number": self.number,
            "shape": self.shape,
            "size": self.size,
            "bg_color": self.bg_color,
            "text_color": self.text_color,
            "font_size": self.font_size,
            "bold": self.bold,
            "prefix": self.prefix,
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
        
        # 文本
        text = f"{self.prefix}{self.number}"
        
        xml = f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name="step_number"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
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
        <a:t>{text}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'''
        return xml
    
    def __repr__(self):
        return f"StepNumberPrimitive(number={self.number})"
