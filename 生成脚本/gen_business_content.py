#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_business_content.py — 业务内容页生成（合并版）

支持从Word文档提取业务规则、要点和关键信息，生成专业的业务内容PPT页面。

用法：
  py -3 gen_business_content.py input.docx --out output.pptx
  py -3 gen_business_content.py input.docx --type rules --out rules.pptx
  py -3 gen_business_content.py input.docx --type points --out points.pptx
"""
import argparse
import json
import os
import re
import sys
import zipfile

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SLIDE_W = 12192000
SLIDE_H = 6858000

# 配色方案
COLORS = {
    "restriction": {"title_bg": "C00000", "mark": "C00000", "content_bg": "FCE4EC"},
    "requirement": {"title_bg": "006100", "mark": "006100", "content_bg": "E2EFDA"},
    "approval": {"title_bg": "7F6000", "mark": "7F6000", "content_bg": "FFF2CC"},
    "point_category": {"title_bg": "1F3864", "mark": "2E75B6", "content_bg": "FFFFFF"},
}

# 图标映射
ICONS = {
    "restriction": "🚫",
    "requirement": "✅",
    "approval": "⚠️",
    "point_category": "📋",
}


def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def rpr(sz, b, color):
    return ('<a:rPr lang="zh-CN" sz="%d" b="%d" dirty="0">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:rPr>'
            % (sz, 1 if b else 0, color))


def run(sz, b, color, text):
    paras = []
    for line in text.split("\n"):
        line = line.strip()
        paras.append('<a:p><a:pPr algn="ctr"/><a:r>%s<a:t>%s</a:t></a:r></a:p>'
                     % (rpr(sz, b, color), esc(line)))
    return "".join(paras)


def box_sp(sid, name, x, y, w, h, text, fill, color, sz=1500, bold=True, prst="roundRect"):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="%s">' % (sid, esc(name), x, y, w, h, prst)
        + ('<a:avLst><a:gd name="adj" fval="8000"/></a:avLst>' if prst in ("round", "roundRect") else "")
        + '</a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln>'
        '</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="91440" rIns="91440" tIns="45720" bIns="45720"/>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
    ) % (fill, run(sz, bold, color, text))


def extract_business_content(docx_path):
    """从Word文档提取业务内容"""
    try:
        from docx import Document
    except ImportError:
        print("错误：需要安装 python-docx 库")
        print("运行：pip install python-docx")
        sys.exit(1)

    doc = Document(docx_path)
    content = {"rules": [], "points": []}
    
    # 规则类关键词
    rule_keywords = ["不允许", "禁止", "不得", "必须", "需要", "要求", "限制", "条件", "审批", "审核"]
    # 要点类关键词
    point_keywords = ["业务要点", "关键信息", "重要提示", "核心规则", "注意事项"]
    
    current_section = None
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        # 检查是否是标题
        if para.style.name.startswith("Heading"):
            current_section = text
            continue
        
        # 检查是否包含规则类关键词
        if any(keyword in text for keyword in rule_keywords):
            content["rules"].append(text)
        # 检查是否包含要点类关键词
        elif any(keyword in text for keyword in point_keywords):
            content["points"].append(text)
        # 如果在业务内容章节下，也提取
        elif current_section and any(kw in current_section for kw in ["规则", "要点", "限制", "要求"]):
            content["rules"].append(text)
    
    return content


def generate_rules_slide(title, subtitle, rules):
    """生成规则类页面"""
    shapes = []
    
    # 标题栏
    shapes.append(
        '<p:sp><p:nvSpPr><p:cNvPr id="1" name="title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="600000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr"/>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
        % (SLIDE_W, run(2400, True, "FFFFFF", title))
    )
    
    # 副标题
    shapes.append(
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="subtitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="500000" y="800000"/><a:ext cx="11192000" cy="400000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr"/>'
        '<a:p><a:r>%s<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
        % (rpr(1800, False, "333333"), esc(subtitle))
    )
    
    # 规则内容
    y = 1400000
    for i, rule in enumerate(rules[:10]):  # 最多10条
        shapes.append(
            '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="rule%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="500000" y="%d"/><a:ext cx="11192000" cy="300000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="FCE4EC"/></a:solidFill>'
            '<a:ln w="9525"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln>'
            '</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="91440"/>'
            '<a:p><a:r>%s<a:t>🚫 %s</a:t></a:r></a:p></p:txBody></p:sp>'
            % (100 + i, y, rpr(1400, False, "333333"), esc(rule[:50]))
        )
        y += 350000
    
    return '\n'.join(shapes)


def generate_points_slide(title, subtitle, points):
    """生成要点类页面"""
    shapes = []
    
    # 标题栏
    shapes.append(
        '<p:sp><p:nvSpPr><p:cNvPr id="1" name="title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="600000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="006100"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr"/>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
        % (SLIDE_W, run(2400, True, "FFFFFF", title))
    )
    
    # 副标题
    shapes.append(
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="subtitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="500000" y="800000"/><a:ext cx="11192000" cy="400000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr"/>'
        '<a:p><a:r>%s<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
        % (rpr(1800, False, "333333"), esc(subtitle))
    )
    
    # 要点卡片（2x2布局）
    icons = ["📋", "🏦", "💰", "⚙️"]
    x_positions = [500000, 6000000]
    y_positions = [1400000, 3800000]
    
    for i, point in enumerate(points[:4]):
        if i >= 4:
            break
        x = x_positions[i % 2]
        y = y_positions[i // 2]
        icon = icons[i % len(icons)]
        
        # 卡片背景
        shapes.append(
            '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="card%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="5200000" cy="2000000"/></a:xfrm>'
            '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fval="8000"/></a:avLst></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            '<a:ln w="9525"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln>'
            '</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="91440"/>'
            '<a:p><a:r>%s<a:t>%s %s</a:t></a:r></a:p></p:txBody></p:sp>'
            % (200 + i, x, y, rpr(1600, True, "333333"), icon, esc(point[:20]))
        )
    
    return '\n'.join(shapes)


def build_pptx(xml_content, output_path):
    """构建PPTX文件"""
    slide_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
        <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
      {xml_content}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

    presentation_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>'''

    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
</Types>'''

    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>'''

    slide_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>'''

    slide_master_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001">
        <a:schemeClr val="bg1"/>
      </p:bgRef>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
        <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap>
    <a:dk1/><a:lt1/><a:dk2/><a:lt2/>
    <a:accent1/><a:accent2/><a:accent3/><a:accent4/>
    <a:accent5/><a:accent6/>
    <a:hlink/><a:folHlink/>
  </p:clrMap>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>'''

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types_xml)
        zf.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
        zf.writestr('ppt/presentation.xml', presentation_xml)
        zf.writestr('ppt/_rels/presentation.xml.rels', rels_xml)
        zf.writestr('ppt/slides/slide1.xml', slide_xml)
        zf.writestr('ppt/slides/_rels/slide1.xml.rels', slide_rels_xml)
        zf.writestr('ppt/slideMasters/slideMaster1.xml', slide_master_xml)
        zf.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')


def main():
    parser = argparse.ArgumentParser(description="业务内容页生成（合并版）")
    parser.add_argument("input", help="输入Word文档")
    parser.add_argument("--out", "-o", default="business_content.pptx", help="输出PPTX文件")
    parser.add_argument("--type", choices=["rules", "points", "mixed", "auto"], default="auto", help="内容类型")
    parser.add_argument("--title", default="业务内容", help="页面标题")
    parser.add_argument("--subtitle", default="", help="副标题")

    args = parser.parse_args()

    # 提取业务内容
    content = extract_business_content(args.input)
    
    # 确定内容类型
    if args.type == "auto":
        if content["rules"] and content["points"]:
            content_type = "mixed"
        elif content["rules"]:
            content_type = "rules"
        else:
            content_type = "points"
    else:
        content_type = args.type
    
    # 生成页面
    if content_type == "rules":
        xml_content = generate_rules_slide(args.title, args.subtitle or "办理条件与限制", content["rules"])
    elif content_type == "points":
        xml_content = generate_points_slide(args.title, args.subtitle or "关键信息汇总", content["points"])
    else:
        # 混合类型：先生成规则，再生成要点
        xml_content = generate_rules_slide(args.title, args.subtitle or "规则与要点", content["rules"])
    
    # 生成PPTX
    build_pptx(xml_content, args.out)
    
    print(f"已生成业务内容页：{args.out}")
    print(f"内容类型：{content_type}")
    print(f"规则数量：{len(content['rules'])}")
    print(f"要点数量：{len(content['points'])}")


if __name__ == "__main__":
    main()
