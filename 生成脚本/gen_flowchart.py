#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_flowchart.py — Word 流程操作 → 一页纵向流程图 PPT（零依赖）

用 Python 标准库直接读写 PPTX/docx 的 XML（两者本质都是 zip+xml），
无需 node / pptxgenjs / python-pptx / pip，本机 py -3 即可运行。

用法:
    py -3 gen_flowchart.py 流程操作.docx --out 流程图.pptx
    py -3 gen_flowchart.py --steps "登录系统; 录入信息; 提交审核" --title 开户流程 --out out.pptx

从 docx 提取规则（保持最简）:
    - 第一个非空段落作为 PPT 标题；
    - 其余非空段落按顺序作为流程步骤（一段一步）。
"""
import argparse
import io
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

# 命名空间
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 16:9 幻灯片尺寸（EMU）
SLIDE_W = 12192000
SLIDE_H = 6858000


def extract_steps_from_docx(path):
    """从 docx 提取 [标题, 步骤...]。零依赖：zip 读 word/document.xml 解析段落文本。"""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    paras = []
    for p in root.iter("{%s}p" % NS_W):
        texts = [t.text or "" for t in p.iter("{%s}t" % NS_W)]
        line = "".join(texts).strip()
        if line:
            paras.append(line)
    if not paras:
        raise SystemExit("未在 docx 中解析到任何文本段落。")
    title = paras[0]
    steps = paras[1:]
    return title, steps


def esc(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))


def shape_text(shape_id, name, x, y, cx, cy, fill, text, font_size=1400,
               bold=True, color="FFFFFF"):
    """构造一个文本框/圆角矩形 shape XML。"""
    return (
        '<p:sp>'
        '  <p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '  <p:spPr>'
        '    <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fval="8000"/></a:avLst></a:prstGeom>'
        '    <a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '    <a:ln w="9525"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:ln>'
        '  </p:spPr>'
        '  <p:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>'
        '    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="%d" b="%d" dirty="0">'
        '      <a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:rPr>'
        '      <a:t>%s</a:t></a:r></a:p></p:txBody>'
        '</p:sp>'
    ) % (shape_id, esc(name), x, y, cx, cy, fill, font_size, 1 if bold else 0,
         color, esc(text))


def arrow_shape(shape_id, x, y, cy):
    """竖向下箭头连接线（tailEnd triangle）。"""
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
    ) % (shape_id, shape_id, x, y, cy, cy)


def build_slide_xml(title, steps):
    """按纵向布局生成一页流程图 slide XML。"""
    # 纵向布局：标题区 + 步骤框（含开始/结束）+ 箭头
    box_w = 7000000
    gap = 300000
    top = 1200000
    bottom_margin = 500000
    avail_h = SLIDE_H - top - bottom_margin

    box_h = (avail_h - gap * (len(steps) - 1)) // len(steps)
    if box_h > 620000:
        box_h = 620000
    box_x = (SLIDE_W - box_w) // 2

    shapes = [shape_text(1, "title", 600000, 250000, SLIDE_W - 1200000, 600000,
                         "FFFFFF", title, 2800, True, "1F3864")]

    sid = 2
    y = top
    for i, step in enumerate(steps):
        if i == 0:
            fill, color = "C00000", "FFFFFF"   # 开始：红
        elif i == len(steps) - 1:
            fill, color = "1E7145", "FFFFFF"   # 结束：绿
        else:
            fill, color = "2E75B6", "FFFFFF"   # 步骤：蓝
        shapes.append(shape_text(sid, "step%d" % (i + 1), box_x, y, box_w,
                                 box_h, fill, step, 1600, True, color))
        sid += 1
        if i < len(steps) - 1:
            ax = box_x + box_w // 2
            shapes.append(arrow_shape(sid, ax, y + box_h + gap // 2, gap))
            sid += 1
            y += box_h + gap

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


# ---------- 最小合法 PPTX 的静态文件 ----------

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
    '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
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
    '  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
    '  <p:sldSz cx="%d" cy="%d"/>'
    '  <p:notesSz cx="6858000" cy="9144000"/>'
    '</p:presentation>'
) % (NS_A, NS_R, NS_P, SLIDE_W, SLIDE_H)

PRES_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>'
    '</Relationships>'
)

SLIDE_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '</Relationships>'
)

SLIDE_MASTER = (
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

SLIDE_MASTER_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
    '</Relationships>'
)

SLIDE_LAYOUT = (
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

SLIDE_LAYOUT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
    '</Relationships>'
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

CORE_PROPS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
    ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '  <dc:title>流程图</dc:title>'
    '  <dc:creator>lark-training-ppt-generator</dc:creator>'
    '</cp:coreProperties>'
)

APP_PROPS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '  <Application>lark-training-ppt-generator</Application>'
    '</Properties>'
)


def build_pptx(slide_xml, out_path):
    files = {
        "[Content_Types].xml": CONTENT_TYPES.encode("utf-8"),
        "_rels/.rels": ROOT_RELS.encode("utf-8"),
        "docProps/core.xml": CORE_PROPS.encode("utf-8"),
        "docProps/app.xml": APP_PROPS.encode("utf-8"),
        "ppt/presentation.xml": PRESENTATION.encode("utf-8"),
        "ppt/_rels/presentation.xml.rels": PRES_RELS.encode("utf-8"),
        "ppt/slides/slide1.xml": slide_xml.encode("utf-8"),
        "ppt/slides/_rels/slide1.xml.rels": SLIDE_RELS.encode("utf-8"),
        "ppt/slideMasters/slideMaster1.xml": SLIDE_MASTER.encode("utf-8"),
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": SLIDE_MASTER_RELS.encode("utf-8"),
        "ppt/slideLayouts/slideLayout1.xml": SLIDE_LAYOUT.encode("utf-8"),
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": SLIDE_LAYOUT_RELS.encode("utf-8"),
        "ppt/theme/theme1.xml": THEME.encode("utf-8"),
    }
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    print("已生成: %s (%d 个步骤)" % (out_path, len([x for x in slide_xml.split('name="step')]) - 1))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Word 流程操作 → 一页纵向流程图 PPT（零依赖）")
    ap.add_argument("docx", nargs="?", help="输入 Word 文档路径（含流程操作段落）")
    ap.add_argument("--steps", help="直接指定步骤，用分号或换行分隔（优先于 docx）")
    ap.add_argument("--title", default=None, help="PPT 标题（默认取 docx 第一段）")
    ap.add_argument("--out", default="流程图.pptx", help="输出 PPTX 路径")
    args = ap.parse_args(argv)

    if args.steps:
        title = args.title or "流程操作"
        steps = [s.strip() for s in args.steps.replace("\n", ";").split(";") if s.strip()]
        if not steps:
            raise SystemExit("--steps 未解析到任何步骤。")
    elif args.docx:
        if not os.path.isfile(args.docx):
            raise SystemExit("输入文件不存在: %s" % args.docx)
        auto_title, steps = extract_steps_from_docx(args.docx)
        title = args.title or auto_title
        if not steps:
            raise SystemExit("docx 中除标题外没有流程步骤段落。")
    else:
        ap.print_help()
        raise SystemExit("需要提供 docx 路径或 --steps。")

    if args.out.lower().endswith(".pptx"):
        out = args.out
    else:
        out = args.out + ".pptx"

    slide_xml = build_slide_xml(title, steps)
    build_pptx(slide_xml, out)


if __name__ == "__main__":
    sys.exit(main())
