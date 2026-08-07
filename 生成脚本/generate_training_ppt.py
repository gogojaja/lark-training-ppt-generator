#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""generate_training_ppt.py — 从Word文档生成培训PPT（参数化）

根据Word文档内容，生成包含以下页面的PPT：
1. 封面页（标题）
2. 场景说明页（淡绿色背景）
3. 业务规则页（淡绿色背景）
4. 业务办理流程页（纵向流程图）

用法:
    py -3 generate_training_ppt.py 输入文档/个人批量开户.docx --out 个人批量开户.pptx
    py -3 generate_training_ppt.py 输入文档/个人批量开户.docx --config config.json --out 个人批量开户.pptx
"""
import argparse
import json
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

SLIDE_W = 12192000
SLIDE_H = 6858000

# 默认颜色配置
DEFAULT_COLORS = {
    "cover_bg": "1F3864",      # 封面背景色（深蓝）
    "cover_text": "FFFFFF",    # 封面文字色（白色）
    "section_bg": "C6EFCE",    # 章节背景色（淡绿）
    "section_text": "006100",  # 章节文字色（深绿）
    "flow_main": "C6EFCE",     # 流程主框色（淡绿）
    "flow_decision": "FFF2CC", # 流程判断框色（淡黄）
    "flow_branch": "DDEBF7",   # 流程分支色（淡蓝）
    "flow_error": "FCE4EC",    # 流程异常色（淡红）
}


def extract_text_from_docx(path):
    """从docx提取文本段落。"""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    paras = []
    for p in root.iter("{%s}p" % NS_W):
        texts = [t.text or "" for t in p.iter("{%s}t" % NS_W)]
        line = "".join(texts).strip()
        if line:
            paras.append(line)
    return paras


def parse_document(paras):
    """解析文档结构，提取场景说明、业务规则、业务办理流程。"""
    result = {
        "title": "",
        "scene_description": [],
        "business_rules": [],
        "flow_steps": []
    }
    
    current_section = None
    for i, p in enumerate(paras):
        # 提取标题
        if i == 0:
            result["title"] = p
            continue
        
        # 识别章节
        if "场景说明" in p:
            current_section = "scene"
            continue
        elif "业务规则" in p:
            current_section = "rules"
            continue
        elif "业务办理流程" in p:
            current_section = "flow"
            continue
        elif p.startswith("步骤"):
            current_section = "flow"
        
        # 根据当前章节添加内容
        if current_section == "scene":
            result["scene_description"].append(p)
        elif current_section == "rules":
            result["business_rules"].append(p)
        elif current_section == "flow":
            # 提取流程步骤
            if p.startswith("步骤"):
                # 提取步骤文本（去掉"步骤X:"前缀）
                step_text = p.split(":", 1)[1] if ":" in p else p
                result["flow_steps"].append(step_text.strip())
    
    return result


def esc(text):
    """XML转义。"""
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))


def rpr(sz, b, color):
    """构造run属性。"""
    return ('<a:rPr lang="zh-CN" sz="%d" b="%d" dirty="0">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:rPr>'
            % (sz, 1 if b else 0, color))


def run(sz, b, color, text):
    """构造多行段落。"""
    paras = []
    for line in text.split("\n"):
        line = line.strip()
        paras.append('<a:p><a:pPr algn="ctr"/><a:r>%s<a:t>%s</a:t></a:r></a:p>'
                     % (rpr(sz, b, color), esc(line)))
    return "".join(paras)


def shape_text(sid, name, x, y, w, h, fill, text, font_size=1400,
               bold=True, color="FFFFFF", prst="roundRect"):
    """构造文本框形状。"""
    if prst == "diamond":
        av_lst = '<a:avLst/>'
        ln_color = "BF9000"
    else:
        av_lst = '<a:avLst><a:gd name="adj" fval="8000"/></a:avLst>'
        ln_color = "1F3864"
    
    return (
        '<p:sp>'
        '  <p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '  <p:spPr>'
        '    <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '    <a:prstGeom prst="%s">%s</a:prstGeom>'
        '    <a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '    <a:ln w="9525"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>'
        '  </p:spPr>'
        '  <p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="91440" rIns="91440" tIns="45720" bIns="45720"/><a:lstStyle/>'
        '    <a:p><a:pPr algn="ctr"/><a:r>%s<a:t>%s</a:t></a:r></a:p></p:txBody>'
        '</p:sp>'
    ) % (sid, esc(name), x, y, w, h, prst, av_lst, fill, ln_color,
         rpr(font_size, bold, color), esc(text))


def shape_text_multiline(sid, name, x, y, w, h, fill, text, font_size=1200,
                         bold=False, color="000000"):
    """构造多行文本框形状。"""
    lines = text.split("\n")
    para_xml = ""
    for line in lines:
        if line.strip():
            para_xml += '<a:p><a:pPr algn="l" marL="91440"/><a:r>%s<a:t>%s</a:t></a:r></a:p>' % (
                rpr(font_size, bold, color), esc(line.strip()))
    
    return (
        '<p:sp>'
        '  <p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '  <p:spPr>'
        '    <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '    <a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '    <a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln>'
        '  </p:spPr>'
        '  <p:txBody><a:bodyPr wrap="square" anchor="t" lIns="91440" rIns="91440" tIns="91440" bIns="91440"/>'
        '  <a:lstStyle/>%s</p:txBody>'
        '</p:sp>'
    ) % (sid, esc(name), x, y, w, h, fill, para_xml)


def arrow_shape(sid, x, y, cy):
    """竖向下箭头连接线。"""
    return (
        '<p:cxnSp>'
        '  <p:nvCxnSpPr><p:cNvPr id="%d" name="arrow%d"/><p:cNvCxnSpPr>'
        '    <a:stCxn/><a:endCxn/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>'
        '  <p:spPr>'
        '    <a:xfrm flipV="0" rot="5400000"><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        '    <a:ln w="28575"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>'
        '      <a:headEnd type="none"/><a:tailEnd type="triangle" w="med" len="med"/>'
        '    </a:ln>'
        '  </p:spPr>'
        '  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:cxnSp>'
    ) % (sid, sid, x, y, cy, cy)


def build_cover_slide(title, colors):
    """构建封面页。"""
    shapes = []
    sid = 1
    
    # 标题
    shapes.append(shape_text(sid, "title", 1000000, 2500000, SLIDE_W - 2000000, 1500000,
                             colors["cover_bg"], title, 4000, True, colors["cover_text"]))
    
    body = "".join(shapes)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '  <p:cSld>'
        '    <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
        '    <p:spTree>'
        '      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>'
        '        <p:nvPr/></p:nvGrpSpPr>'
        '      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '        <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '      %s'
        '    </p:spTree>'
        '  </p:cSld>'
        '  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sld>'
    ) % (NS_A, NS_R, NS_P, body)


def build_text_slide(title, content, colors, is_scene=True):
    """构建文本页（场景说明或业务规则）。"""
    shapes = []
    sid = 1
    
    # 标题
    shapes.append(shape_text(sid, "title", 0, 0, SLIDE_W, 800000,
                             colors["section_bg"], title, 2800, True, colors["section_text"]))
    sid += 1
    
    # 内容
    content_text = "\n".join(content)
    shapes.append(shape_text_multiline(sid, "content", 500000, 1000000, 
                                        SLIDE_W - 1000000, SLIDE_H - 1200000,
                                        colors["section_bg"], content_text, 
                                        1400, False, colors["section_text"]))
    
    body = "".join(shapes)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '  <p:cSld>'
        '    <p:spTree>'
        '      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>'
        '        <p:nvPr/></p:nvGrpSpPr>'
        '      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '        <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '      %s'
        '    </p:spTree>'
        '  </p:cSld>'
        '  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sld>'
    ) % (NS_A, NS_R, NS_P, body)


def build_flow_slide(title, steps, colors):
    """构建流程图页。"""
    shapes = []
    sid = 1
    
    # 标题
    shapes.append(shape_text(sid, "title", 0, 0, SLIDE_W, 800000,
                             "1F3864", title, 2800, True, "FFFFFF"))
    sid += 1
    
    # 流程步骤
    box_w = 7000000
    gap = 300000
    top = 1000000
    bottom_margin = 500000
    avail_h = SLIDE_H - top - bottom_margin
    
    box_h = (avail_h - gap * (len(steps) - 1)) // len(steps)
    if box_h > 600000:
        box_h = 600000
    box_x = (SLIDE_W - box_w) // 2
    
    y = top
    for i, step in enumerate(steps):
        # 判断是否为判断节点
        is_decision = "？" in step or "是否" in step
        
        if i == 0:
            fill, color = "C00000", "FFFFFF"   # 开始：红
            prst = "roundRect"
        elif i == len(steps) - 1:
            fill, color = "1E7145", "FFFFFF"   # 结束：绿
            prst = "roundRect"
        elif is_decision:
            fill, color = colors["flow_decision"], "7F6000"   # 判断：浅黄
            prst = "diamond"
        else:
            fill, color = colors["flow_main"], "006100"   # 步骤：浅绿
            prst = "roundRect"
        
        h = box_h * 1.3 if is_decision else box_h
        
        shapes.append(shape_text(sid, "step%d" % (i + 1), box_x, y, box_w,
                                 int(h), fill, step, 1400, True, color, prst))
        sid += 1
        
        # 连接线
        if i < len(steps) - 1:
            ax = box_x + box_w // 2
            shapes.append(arrow_shape(sid, ax, y + int(h) + gap // 2, gap))
            sid += 1
            y += int(h) + gap
    
    body = "".join(shapes)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '  <p:cSld>'
        '    <p:spTree>'
        '      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>'
        '        <p:nvPr/></p:nvGrpSpPr>'
        '      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '        <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '      %s'
        '    </p:spTree>'
        '  </p:cSld>'
        '  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sld>'
    ) % (NS_A, NS_R, NS_P, body)


# ---------- PPTX静态文件 ----------

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
    '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    '<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    '<Override PartName="/ppt/slides/slide4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    '</Types>'
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '</Relationships>'
)

PRESENTATION = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
    '  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
    '  <p:sldIdLst>'
    '    <p:sldId id="256" r:id="rId2"/>'
    '    <p:sldId id="257" r:id="rId3"/>'
    '    <p:sldId id="258" r:id="rId4"/>'
    '    <p:sldId id="259" r:id="rId5"/>'
    '  </p:sldIdLst>'
    '  <p:sldSz cx="%d" cy="%d"/>'
    '  <p:notesSz cx="6858000" cy="9144000"/>'
    '</p:presentation>'
) % (NS_A, NS_R, NS_P, SLIDE_W, SLIDE_H)

PRES_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>'
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>'
    '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide4.xml"/>'
    '</Relationships>'
)

SLIDE_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '</Relationships>'
)

MASTER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
    '  <p:cSld><p:spTree>'
    '    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '      <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '  </p:spTree></p:cSld>'
    '  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"'
    '    accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"'
    '    hlink="hlink" folHlink="folHlink"/>'
    '  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    '</p:sldMaster>'
) % (NS_A, NS_R, NS_P)

MASTER_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
    '</Relationships>'
)

LAYOUT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" type="blank" preserve="1">'
    '  <p:cSld name="Blank"><p:spTree>'
    '    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '      <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '  </p:spTree></p:cSld>'
    '  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
    '</p:sldLayout>'
) % (NS_A, NS_R, NS_P)

LAYOUT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
    '</Relationships>'
)

CORE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
    ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '  <dc:title>培训PPT</dc:title>'
    '  <dc:creator>lark-training-ppt-generator</dc:creator>'
    '</cp:coreProperties>'
)

APP = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '  <Application>lark-training-ppt-generator</Application>'
    '</Properties>'
)

THEME = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<a:theme xmlns:a="%s" name="Office">'
    '  <a:themeElements>'
    '    <a:clrScheme name="Office">'
    '      <a:dk1><a:srgbClr val="000000"/></a:dk1>'
    '      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
    '      <a:dk2><a:srgbClr val="44546A"/></a:dk2>'
    '      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
    '      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>'
    '      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
    '      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>'
    '      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
    '      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>'
    '      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
    '      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
    '      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
    '    </a:clrScheme>'
    '    <a:fontScheme name="Office">'
    '      <a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:majorFont>'
    '      <a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:minorFont>'
    '    </a:fontScheme>'
    '    <a:fmtScheme name="Office">'
    '      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
    '      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
    '      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    '      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
    '    </a:fmtScheme>'
    '  </a:themeElements>'
    '</a:theme>'
) % NS_A


def build_pptx(slides, out_path):
    """构建PPTX文件。"""
    files = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "docProps/core.xml": CORE,
        "docProps/app.xml": APP,
        "ppt/presentation.xml": PRESENTATION,
        "ppt/_rels/presentation.xml.rels": PRES_RELS,
        "ppt/slideMasters/slideMaster1.xml": MASTER,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": MASTER_RELS,
        "ppt/slideLayouts/slideLayout1.xml": LAYOUT,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": LAYOUT_RELS,
        "ppt/theme/theme1.xml": THEME,
    }
    
    # 添加幻灯片
    for i, slide_xml in enumerate(slides):
        files["ppt/slides/slide%d.xml" % (i + 1)] = slide_xml
        files["ppt/slides/_rels/slide%d.xml.rels" % (i + 1)] = SLIDE_RELS
    
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data.encode("utf-8"))
    
    print("已生成: %s (%d 页)" % (out_path, len(slides)))


def main(argv=None):
    ap = argparse.ArgumentParser(description="从Word文档生成培训PPT（参数化）")
    ap.add_argument("docx", help="输入Word文档路径")
    ap.add_argument("--config", help="配置文件路径（JSON格式）")
    ap.add_argument("--out", default="培训PPT.pptx", help="输出PPTX路径")
    ap.add_argument("--title", default=None, help="PPT标题（覆盖文档标题）")
    ap.add_argument("--scene-title", default="场景说明", help="场景说明页标题")
    ap.add_argument("--rules-title", default="业务规则", help="业务规则页标题")
    ap.add_argument("--flow-title", default="业务办理流程", help="流程图页标题")
    args = ap.parse_args(argv)
    
    # 读取配置文件
    colors = DEFAULT_COLORS.copy()
    if args.config and os.path.isfile(args.config):
        with open(args.config, encoding="utf-8") as f:
            config = json.load(f)
            if "colors" in config:
                colors.update(config["colors"])
    
    # 提取文档内容
    paras = extract_text_from_docx(args.docx)
    doc = parse_document(paras)
    
    # 使用命令行参数覆盖
    if args.title:
        doc["title"] = args.title
    
    print("文档标题: %s" % doc["title"])
    print("场景说明: %d 段" % len(doc["scene_description"]))
    print("业务规则: %d 段" % len(doc["business_rules"]))
    print("业务流程: %d 步" % len(doc["flow_steps"]))
    
    # 构建幻灯片
    slides = []
    
    # 1. 封面页
    slides.append(build_cover_slide(doc["title"], colors))
    
    # 2. 场景说明页
    if doc["scene_description"]:
        slides.append(build_text_slide(args.scene_title, doc["scene_description"], colors, is_scene=True))
    
    # 3. 业务规则页
    if doc["business_rules"]:
        slides.append(build_text_slide(args.rules_title, doc["business_rules"], colors, is_scene=False))
    
    # 4. 流程图页
    if doc["flow_steps"]:
        slides.append(build_flow_slide(args.flow_title, doc["flow_steps"], colors))
    
    # 生成PPTX
    build_pptx(slides, args.out)


if __name__ == "__main__":
    sys.exit(main())
