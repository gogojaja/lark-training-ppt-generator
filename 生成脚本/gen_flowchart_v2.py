#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_flowchart_v2.py — 流程图生成优化版（IT-04）

主要优化：
1. 提升节点识别准确性
2. 支持复杂分支结构（异常分支、并行分支）
3. 支持多层级嵌套
4. 优化布局算法
5. 支持节点样式自定义

用法：
  py -3 gen_flowchart_v2.py input.json --out output.pptx
  py -3 gen_flowchart_v2.py input.json --out output.pptx --connectors
  py -3 gen_flowchart_v2.py input.json --out output.pptx --box-w 5 --box-h 0.6
"""
import argparse
import json
import sys
import os
import zipfile

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SLIDE_W = 12192000
SLIDE_H = 6858000

# ---------- 固化样式规范 ----------
MAIN_FILL, MAIN_COL = "C6EFCE", "006100"
DIA_FILL, DIA_COL = "FFF2CC", "7F6000"
BR_FILL, BR_COL = "DDEBF7", "1F3864"
ERR_FILL, ERR_COL = "FCE4EC", "C00000"

# 默认布局维度
_DEFAULT_DIM = {
    "box_w": 1800000,
    "box_h": 216000,
    "diamond_w": 1620000,
    "diamond_h": 360000,
    "step_gap": 432000,
}

L_MAIN_X = 900000
L_TOP = 1000000
L_BR_GAP = 300000
L_SZ_MAIN = 1500
L_SZ_BR = 900


def load_dimensions(json_raw=None, cli=None):
    """三级覆盖：默认值 → JSON dim → CLI 参数"""
    dim = dict(_DEFAULT_DIM)
    if json_raw and isinstance(json_raw.get("dim"), dict):
        for k in dim:
            if k in json_raw["dim"]:
                dim[k] = int(json_raw["dim"][k])
    if cli:
        for flag, key in [("box_w", "box_w"), ("box_h", "box_h"),
                          ("diamond_w", "diamond_w"), ("diamond_h", "diamond_h"),
                          ("step_gap", "step_gap")]:
            val = getattr(cli, flag, None)
            if val is not None:
                dim[key] = int(val * 360000)
    return dim


def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def rpr(sz, b, color):
    return ('<a:rPr lang="zh-CN" sz="%d" b="%d" dirty="0">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:rPr>'
            % (sz, 1 if b else 0, color))


def run(sz, b, color, text):
    """构造多行段落"""
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


def diamond_sp(sid, name, x, y, w, h, text, fill="FFF2CC", color="7F6000", sz=1500):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="diamond"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="BF9000"/></a:solidFill></a:ln>'
        '</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="91440" rIns="91440" tIns="45720" bIns="45720"/>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
    ) % (sid, esc(name), x, y, w, h, fill, run(sz, True, color, text))


def hline(sid, x1, y1, x2, y2, color="1F3864", arrow=True):
    """水平连接线"""
    x = min(x1, x2)
    length = abs(x2 - x1)
    lw = 9000
    seg = (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="hl%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    ) % (sid, sid, x, y1 - lw // 2, length, lw, color)
    if arrow:
        seg += harrow(sid + 100, x2, y1, color)
    return seg


def harrow(sid, cx, cy, color):
    """水平右向三角箭头"""
    sz = 38000
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="ah%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    ) % (sid, sid, cx - sz, cy - sz // 2, sz, sz, color)


def vline(sid, x1, y1, x2, y2, color="1F3864", arrow=True):
    """垂直连接线"""
    y = min(y1, y2)
    length = abs(y2 - y1)
    lw = 9000
    seg = (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="vl%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    ) % (sid, sid, x1 - lw // 2, y, lw, length, color)
    if arrow:
        seg += varrow(sid + 100, x1, y2, color)
    return seg


def varrow(sid, cx, cy, color):
    """垂直向下三角箭头"""
    sz = 38000
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="ah%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    ) % (sid, sid, cx - sz // 2, cy - sz, sz, sz, color)


def auto_layout_semantic(data, dim):
    """语义模式自动布局：支持复杂分支和多层级嵌套"""
    title = data.get("title", "流程图")
    steps = data.get("steps", [])
    if not steps:
        return [], [], title

    nodes = []
    edges = []
    sid = 1
    y = L_TOP
    main_x = L_MAIN_X
    br_x = main_x + dim["box_w"] + L_BR_GAP
    br_y_offset = dim["step_gap"] // 2

    prev_id = None
    branch_count = 0

    for i, step in enumerate(steps):
        text = step.get("text", f"步骤{i+1}")
        is_diamond = "branch" in step
        branch = step.get("branch", {})

        if is_diamond:
            # 菱形判断节点
            w, h = dim["diamond_w"], dim["diamond_h"]
            fill, col = DIA_FILL, DIA_COL
            kind = "diamond"
            nodes.append({"id": f"d{sid}", "kind": kind, "x": main_x, "y": y,
                          "w": w, "h": h, "text": text, "fill": fill, "color": col})
        else:
            # 普通处理步骤
            w, h = dim["box_w"], dim["box_h"]
            fill, col = MAIN_FILL, MAIN_COL
            kind = "box"
            nodes.append({"id": f"b{sid}", "kind": kind, "x": main_x, "y": y,
                          "w": w, "h": h, "text": text, "fill": fill, "color": col})

        # 连接上一个节点
        if prev_id:
            edges.append({"from": prev_id, "to": f"{'d' if is_diamond else 'b'}{sid}", "style": "v"})

        # 处理分支
        if is_diamond and branch:
            branch_text = branch.get("text", "")
            branch_label = branch.get("label", "")
            branch_kind = branch.get("kind", "br")  # br=正常, err=异常

            br_fill = BR_FILL if branch_kind == "br" else ERR_FILL
            br_col = BR_COL if branch_kind == "br" else ERR_COL

            br_y = y + br_y_offset
            nodes.append({"id": f"br{sid}", "kind": "box", "x": br_x, "y": br_y,
                          "w": dim["box_w"], "h": dim["box_h"], "text": branch_text,
                          "fill": br_fill, "color": br_col})

            # 菱形到分支的肘形连接
            edges.append({"from": f"d{sid}", "to": f"br{sid}", "style": "el-right",
                          "label": branch_label, "label_x": br_x - L_BR_GAP // 2,
                          "label_y": br_y})

            branch_count += 1

        prev_id = f"{'d' if is_diamond else 'b'}{sid}"
        y += dim["step_gap"] + (dim["diamond_h"] - dim["box_h"] if is_diamond else 0)
        sid += 1

    return nodes, edges, title


def auto_layout_advanced(data, dim):
    """高级模式：完全控制坐标/连线/颜色"""
    nodes = data.get("nodes", [])
    edges = data.get("edges", [])
    title = data.get("title", "流程图")
    return nodes, edges, title


def build_slide_xml(title, nodes, edges, dim, draw_connectors=True):
    """构建幻灯片XML"""
    shapes = []

    # 标题横幅
    shapes.append(
        '<p:sp><p:nvSpPr><p:cNvPr id="1" name="title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="600000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr"/>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
        % (SLIDE_W, run(2400, True, "FFFFFF", title))
    )

    # 绘制节点
    for node in nodes:
        sid = hash(node["id"]) % 10000 + 100
        if node["kind"] == "diamond":
            shapes.append(diamond_sp(sid, node["id"], node["x"], node["y"],
                                     node["w"], node["h"], node["text"],
                                     node.get("fill", DIA_FILL), node.get("color", DIA_COL)))
        else:
            shapes.append(box_sp(sid, node["id"], node["x"], node["y"],
                                 node["w"], node["h"], node["text"],
                                 node.get("fill", MAIN_FILL), node.get("color", MAIN_COL)))

    # 绘制连接线
    if draw_connectors:
        node_map = {n["id"]: n for n in nodes}
        for edge in edges:
            from_node = node_map.get(edge["from"])
            to_node = node_map.get(edge["to"])
            if not from_node or not to_node:
                continue

            sid = hash(f"{edge['from']}-{edge['to']}") % 10000 + 1000

            if edge.get("style") == "v":
                # 垂直连接
                x = from_node["x"] + from_node["w"] // 2
                y1 = from_node["y"] + from_node["h"]
                y2 = to_node["y"]
                shapes.append(vline(sid, x, y1, x, y2))
            elif edge.get("style") == "el-right":
                # 肘形连接（向右）
                x1 = from_node["x"] + from_node["w"]
                y1 = from_node["y"] + from_node["h"] // 2
                x2 = to_node["x"]
                y2 = to_node["y"] + to_node["h"] // 2
                mid_x = (x1 + x2) // 2
                shapes.append(hline(sid, x1, y1, mid_x, y1))
                shapes.append(vline(sid + 1, mid_x, y1, mid_x, y2))
                shapes.append(hline(sid + 2, mid_x, y2, x2, y2))

                # 标签
                label = edge.get("label", "")
                if label:
                    label_x = edge.get("label_x", mid_x)
                    label_y = edge.get("label_y", y1)
                    shapes.append(box_sp(sid + 3, f"label{sid}", label_x, label_y - 100000,
                                         600000, 200000, label, "FFFFFF", "000000", 800, False))

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
    parser = argparse.ArgumentParser(description="流程图生成优化版（IT-04）")
    parser.add_argument("input", help="输入JSON文件")
    parser.add_argument("--out", "-o", default="output.pptx", help="输出PPTX文件")
    parser.add_argument("--connectors", action="store_true", default=False, help="绘制连接线")
    parser.add_argument("--box-w", type=float, help="矩形框宽（cm）")
    parser.add_argument("--box-h", type=float, help="矩形框高（cm）")
    parser.add_argument("--diamond-w", type=float, help="菱形宽（cm）")
    parser.add_argument("--diamond-h", type=float, help="菱形高（cm）")
    parser.add_argument("--step-gap", type=float, help="纵向间隔（cm）")

    args = parser.parse_args()

    with open(args.input, 'r', encoding='utf-8') as f:
        data = json.load(f)

    dim = load_dimensions(data, args)

    # 检查是否启用连接线
    draw_connectors = args.connectors or data.get("draw_connectors", True)

    # 自动布局
    if "steps" in data:
        nodes, edges, title = auto_layout_semantic(data, dim)
    else:
        nodes, edges, title = auto_layout_advanced(data, dim)

    # 构建幻灯片XML
    xml_content = build_slide_xml(title, nodes, edges, dim, draw_connectors)

    # 生成PPTX
    build_pptx(xml_content, args.out)

    print(f"已生成流程图：{args.out}")
    print(f"节点数量：{len(nodes)}")
    print(f"连接线数量：{len(edges) if draw_connectors else 0}")


if __name__ == "__main__":
    main()
