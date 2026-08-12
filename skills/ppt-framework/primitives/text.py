"""
TextPrimitive - 文本原语

支持角色驱动的文本渲染，自动应用对应的字号、字重、行高。

== 用法 ==
    from primitives import TextPrimitive
    
    # 角色驱动
    text = TextPrimitive(text="页面标题", role="claim")
    
    # 自定义样式
    text = TextPrimitive(text="自定义", font_size=20, bold=True, color="FF0000")
"""

# 角色配置（基于 styles.md）
ROLE_CONFIG = {
    "claim": {"font_size": 32, "bold": True, "line_height": 1.2, "color": "1F3864"},
    "sectionLabel": {"font_size": 12, "bold": False, "line_height": 1.3, "color": "999999"},
    "body": {"font_size": 14, "bold": False, "line_height": 1.5, "color": "333333"},
    "annotation": {"font_size": 12, "bold": False, "line_height": 1.4, "color": "666666"},
    "source": {"font_size": 10, "bold": False, "line_height": 1.3, "color": "999999"},
    # 向后兼容
    "H1": {"font_size": 36, "bold": True, "line_height": 1.2, "color": "1F3864"},
    "H2": {"font_size": 28, "bold": True, "line_height": 1.3, "color": "1F3864"},
    "H3": {"font_size": 22, "bold": True, "line_height": 1.4, "color": "333333"},
    "H4": {"font_size": 18, "bold": True, "line_height": 1.4, "color": "333333"},
    "Body": {"font_size": 14, "bold": False, "line_height": 1.5, "color": "333333"},
    "Caption": {"font_size": 12, "bold": False, "line_height": 1.4, "color": "666666"},
    "Small": {"font_size": 10, "bold": False, "line_height": 1.3, "color": "999999"},
}


class TextPrimitive:
    """文本原语"""
    
    def __init__(self, text, role=None, font_size=None, bold=None, 
                 color=None, font_family="Microsoft YaHei", align="left"):
        """
        初始化文本原语
        
        Args:
            text: 文本内容
            role: 角色（claim/sectionLabel/body/annotation/source）
            font_size: 字号（pt），覆盖角色默认值
            bold: 是否加粗，覆盖角色默认值
            color: 颜色（6位HEX），覆盖角色默认值
            font_family: 字体家族
            align: 对齐方式（left/center/right）
        """
        self.text = text
        self.role = role
        self.font_family = font_family
        self.align = align
        
        # 获取角色配置
        role_config = ROLE_CONFIG.get(role, ROLE_CONFIG["body"])
        
        # 应用配置（自定义值覆盖角色值）
        self.font_size = font_size if font_size is not None else role_config["font_size"]
        self.bold = bold if bold is not None else role_config["bold"]
        self.color = color if color is not None else role_config["color"]
        self.line_height = role_config.get("line_height", 1.5)
    
    def to_dict(self):
        """转换为字典"""
        return {
            "type": "text",
            "text": self.text,
            "role": self.role,
            "font_size": self.font_size,
            "bold": self.bold,
            "color": self.color,
            "font_family": self.font_family,
            "align": self.align,
            "line_height": self.line_height,
        }
    
    def to_xml(self, x=0, y=0, width=None, height=None):
        """生成PPTX XML"""
        # 计算尺寸
        if width is None:
            width = len(self.text) * self.font_size * 100  # 粗略估计
        if height is None:
            height = int(self.font_size * self.line_height * 12700 * 1.5)
        
        # 对齐方式
        align_map = {"left": "l", "center": "ctr", "right": "r"}
        align_val = align_map.get(self.align, "l")
        
        # 字重
        bold_val = "1" if self.bold else "0"
        
        xml = f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name="text_{self.role or 'custom'}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="{align_val}"/>
      <a:r>
        <a:rPr lang="zh-CN" sz="{self.font_size * 100}" b="{bold_val}" dirty="0">
          <a:solidFill><a:srgbClr val="{self.color}"/></a:solidFill>
          <a:latin typeface="{self.font_family}"/>
          <a:ea typeface="{self.font_family}"/>
        </a:rPr>
        <a:t>{self.text}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>'''
        return xml
    
    def __repr__(self):
        return f"TextPrimitive(text='{self.text}', role='{self.role}')"
