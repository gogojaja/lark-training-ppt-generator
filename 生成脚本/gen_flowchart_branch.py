#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_flowchart_branch.py — 带分支/判断的一页流程图 PPT（零依赖）

复刻样例风格：顶部深蓝标题横幅 + 方框/菱形判断 + 肘形分支 + 分支标注。
读入一个 JSON（nodes + edges），输出 16:9 一页 PPTX。

== 两种输入模式 ==

模式 A · 语义模式（推荐）：JSON 只描述步骤与分支，工具自动布局与配色。
{
  "title": "个人批量开户业务办理流程（纵向）",
  "steps": [
    {"text": "登录系统进入场景"},                          // 主流程步骤（浅绿）
    {"text": "客户信息是否齐全",                           // 含 branch 自动变菱形判断（浅黄）
     "branch": {"text": "跳转客户信息维护", "label": "否", "kind": "err"}},  // kind: err=异常(浅红) / br=正常(浅蓝)
    {"text": "企业身份核实"}
  ]
}
自动布局规则（固化）：纵向单列，判断右侧分支（肘形 + 是/否标签）。

模式 B · 高级模式：完全控制坐标/连线/颜色（nodes + edges）。
{
  "title": "个人批量开户流程",
  "nodes": [
    {"id":"n1","kind":"box","x":..,"y":..,"w":..,"h":..,"text":".."},   // kind: box/short/diamond
  ],
  "edges": [
    {"from":"n1","to":"n2","style":"v"},                       // 垂直
    {"from":"d1","to":"b1","style":"el-right","label":"否",     // 肘形分往右侧
     "label_x":..,"label_y":..},
  ]
}
坐标单位 EMU; 颜色为六位 RGB 十六进制（不含 #）。

== 固化样式规范（自动套用，也可在 JSON 中覆盖）==
主流程框:  浅绿 C6EFCE / 深绿字 006100
菱形判断:  浅黄 FFF2CC / 深黄字 7F6000
正常分支:  浅蓝 DDEBF7 / 深蓝字 1F3864
异常分支:  浅红 FCE4EC / 红字 C00000
"""
import argparse
import json
import sys
import zipfile

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SLIDE_W = 12192000
SLIDE_H = 6858000

# ---------- 固化样式规范 ----------
# 配色：主流程浅绿 / 菱形浅黄 / 正常分支浅蓝 / 异常分支浅红
MAIN_FILL, MAIN_COL = "C6EFCE", "006100"     # 主流程框
DIA_FILL, DIA_COL = "FFF2CC", "7F6000"       # 菱形判断
BR_FILL, BR_COL = "DDEBF7", "1F3864"         # 正常分支
ERR_FILL, ERR_COL = "FCE4EC", "C00000"       # 异常分支

# 纵向紧凑布局参数（语义模式自动排版）
L_MAIN_X = 900000     # 主流程列 X
L_MAIN_W = 2600000    # 主流程框宽
L_BOX_H = 216000      # 主流程框高（0.6cm = 0.6 * 360000）
L_DIA_H = 216000      # 菱形高
L_STEP_H = 275000     # 行节距（略大于框高留出间隙）
L_TOP = 1000000       # 首行 Y
L_BR_W = 2700000      # 分支框宽
L_BR_GAP = 300000     # 主流程与分支列间距
L_SZ_MAIN = 1000      # 主流程字号
L_SZ_BR = 900         # 分支字号


def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def rpr(sz, b, color):
    return ('<a:rPr lang="zh-CN" sz="%d" b="%d" dirty="0">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:rPr>'
            % (sz, 1 if b else 0, color))


def run(sz, b, color, text):
    """构造多行段落：每个 \n 拆成一个 <a:p>。"""
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
        + ('<a:avLst><a:gd name="adj" fval="8000"/></a:avLst>' if prst == "round" or prst == "roundRect" else "")
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


def _arrowhead_sp(sid, cx, cy_tip, direction, color):
    """在 (cx, cy_tip) 画三角箭头。direction: 'up'|'down'|'left'|'right'。
    用三角形 autoshape，旋转 + 平移定位。"""
    size = 38000
    rot = {"down": "0", "up": "5400000", "right": "0", "left": "5400000"}[direction]
    if direction in ("down", "up"):
        x = cx - size // 2
        y = cy_tip - size if direction == "down" else cy_tip
    else:
        x = cx - size if direction == "right" else cx
        y = cy_tip - size // 2
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="ah%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm rot="%s"><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    ) % (sid, sid, rot, x, y, size, size, color)


def hline(sid, x1, y1, x2, y2, color="1F3864", arrow=True):
    """水平连接线（细矩形），从 (x1,y1) 到 (x2,y2)。可选右侧三角箭头。"""
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
        tip_x = x2
        seg += _arrowhead_sp(sid + 1, tip_x, y1, "right", color)
    return seg


def vline(sid, x1, y1, x2, y2, color="1F3864", arrow=True):
    """垂直连接线（细矩形），从 (x1,y1) 到 (x2,y2)。可选底部三角箭头。"""
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
        tip_y = y2
        seg += _arrowhead_sp(sid + 1, x1, tip_y, "down", color)
    return seg


def _v(sid, na, nb, color):
    x1 = na["x"] + na["w"] // 2
    y1 = na["y"] + na["h"]
    x2 = nb["x"] + nb["w"] // 2
    y2 = nb["y"]
    return vline(sid, x1, y1, x2, y2, color)


def _elbow(sid, na, nb, color):
    """源到右侧目标：水平到目标 X，再垂直向下/上到目标。"""
    x1 = na["x"] + na["w"]
    y1 = na["y"] + na["h"] // 2
    x2 = nb["x"]
    y2 = nb["y"] + nb["h"] // 2
    segs = []
    segs.append(hline(sid, x1, y1, x2, y1, color, arrow=False))
    segs.append(vline(sid + 1, x2, y1, x2, y2, color))
    return segs


def _elbow_left(sid, na, nb, color):
    """从源向左侧目标：先垂直到目标 Y，再水平往左。"""
    x1 = na["x"]
    y1 = na["y"] + na["h"] // 2
    x2 = nb["x"] + nb["w"]
    y2 = nb["y"] + nb["h"] // 2
    segs = []
    segs.append(vline(sid, x1, y1, x1, y2, color, arrow=False))
    segs.append(hline(sid + 1, x1, y2, x2, y2, color))
    return segs


def label_sp(sid, x, y, text, color="C00000", sz=1100):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="lb%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="500000" cy="350000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
        '<a:r>%s<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
    ) % (sid, sid, x, y, rpr(sz, True, color), esc(text))


def auto_layout(flow):
    """语义模式：steps → 纵向紧凑布局 nodes+edges（固化配色与间距）。

    step 结构: {"text":.., "branch": {"text":.., "label":.., "kind":"err"|"br"}}
    - 有 branch 的步骤渲染为菱形判断（浅黄），branch 渲染为右侧分支框
    - branch.kind="err" → 异常分支（浅红）；"br" → 正常分支（浅蓝）
    """
    steps = flow.get("steps", [])
    if not steps:
        raise SystemExit("语义模式需提供 steps 列表。")
    nodes, edges = [], []
    prev = None
    for i, st in enumerate(steps):
        text = st.get("text", "")
        has_branch = bool(st.get("branch")) or st.get("type") == "diamond"
        y = L_TOP + i * L_STEP_H
        h = L_DIA_H if has_branch else L_BOX_H
        kind = "diamond" if has_branch else "box"
        nodes.append({"id": "m%d" % i, "kind": kind, "x": L_MAIN_X, "y": y,
                      "w": L_MAIN_W, "h": h, "text": text, "sz": L_SZ_MAIN})
        if prev:
            edges.append({"from": prev, "to": "m%d" % i, "style": "v"})
        prev = "m%d" % i
        br = st.get("branch")
        if br:
            is_err = br.get("kind") == "err"
            nodes.append({"id": "b%d" % i, "kind": "short", "err": is_err,
                          "x": L_MAIN_X + L_MAIN_W + L_BR_GAP,
                          "y": y + h // 2 - L_BOX_H // 2,
                          "w": L_BR_W, "h": L_BOX_H,
                          "text": br.get("text", ""), "sz": L_SZ_BR})
            edges.append({"from": "m%d" % i, "to": "b%d" % i, "style": "el-right",
                          "label": br.get("label", ""),
                          "label_x": L_MAIN_X + L_MAIN_W + 110000,
                          "label_y": y + h // 2 - 80000,
                          "label_color": ERR_COL if is_err else BR_COL})
    return {"title": flow.get("title", "流程图"), "nodes": nodes, "edges": edges}


def node_style(n):
    """按 kind 返回 (fill, color)。高级模式可显式覆盖。"""
    kind = n.get("kind") or n.get("type") or "box"
    if kind == "diamond":
        return (n.get("fill", DIA_FILL), n.get("color", DIA_COL))
    if kind == "short":
        if n.get("err"):
            return (n.get("fill", ERR_FILL), n.get("color", ERR_COL))
        return (n.get("fill", BR_FILL), n.get("color", BR_COL))
    return (n.get("fill", MAIN_FILL), n.get("color", MAIN_COL))


def build(slide_xml, out_path):
    files = {
        "[Content_Types].xml": CT,
        "_rels/.rels": ROOT_RELS,
        "docProps/core.xml": CORE,
        "docProps/app.xml": APP,
        "ppt/presentation.xml": PRES,
        "ppt/_rels/presentation.xml.rels": PRES_RELS,
        "ppt/slides/slide1.xml": slide_xml,
        "ppt/slides/_rels/slide1.xml.rels": SLIDE_RELS,
        "ppt/slideMasters/slideMaster1.xml": MASTER,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": MASTER_RELS,
        "ppt/slideLayouts/slideLayout1.xml": LAYOUT,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": LAYOUT_RELS,
        "ppt/theme/theme1.xml": THEME,
    }
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b.encode("utf-8"))
    print("已生成: %s" % out_path)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("json", help="流程 JSON 文件（含 steps=语义模式；含 nodes=高级模式）")
    ap.add_argument("--out", default="流程图.pptx")
    a = ap.parse_args(argv)
    with open(a.json, encoding="utf-8") as f:
        raw = json.load(f)

    if "steps" in raw and "nodes" not in raw:
        flow = auto_layout(raw)
    else:
        flow = raw

    body = []
    sid = 1
    body.append(box_sp(sid, "hd", 0, 0, SLIDE_W, 850000, flow.get("title", "流程图"),
                       "1F3864", "FFFFFF", sz=2600, bold=True))
    sid += 1
    pos = {}
    for n in flow.get("nodes", []):
        pos[n["id"]] = n
        fill, color = node_style(n)
        if n.get("kind") == "diamond" or n.get("type") == "diamond":
            body.append(diamond_sp(sid, n["id"], n["x"], n["y"], n["w"], n["h"],
                                   n["text"], fill, color, n.get("sz", 1500)))
        else:
            body.append(box_sp(sid, n["id"], n["x"], n["y"], n["w"], n["h"],
                               n["text"], fill, color, n.get("sz", 1500)))
        sid += 1

    for e in flow.get("edges", []):
        na, nb = pos[e["from"]], pos[e["to"]]
        color = e.get("color", "1F3864")
        style = e.get("style", "v")
        if style == "v":
            body.append(_v(sid, na, nb, color)); sid += 1
        elif style == "el":
            for s in _elbow(sid, na, nb, color):
                body.append(s); sid += 1
        elif style == "el-left":
            for s in _elbow_left(sid, na, nb, color):
                body.append(s); sid += 1
        elif style == "h":
            x1 = na["x"] + na["w"]; y1 = na["y"] + na["h"] // 2
            x2 = nb["x"]; y2 = nb["y"] + nb["h"] // 2
            body.append(hline(sid, x1, y1, x2, y2, color=color, arrow=True)); sid += 1
        elif style == "el-right":
            x1 = na["x"] + na["w"]; y1 = na["y"] + na["h"] // 2
            x2 = nb["x"]; y2 = nb["y"] + nb["h"] // 2
            body.append(hline(sid, x1, y1, x2, y1, color=color, arrow=True)); sid += 1
        if e.get("label"):
            body.append(label_sp(sid, e.get("label_x", 0), e.get("label_y", 0),
                                 e["label"], e.get("label_color", "C00000"))); sid += 1

    slide = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '%s</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    ) % (NS_A, NS_R, NS_P, "".join(body))
    build(slide, a.out)


# ---------- 静态资源 ----------
X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
CT = (X + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
      '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
      '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
      '</Types>')
ROOT_RELS = (X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
             '</Relationships>')
PRES = (X + '<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
        '<p:sldSz cx="%d" cy="%d"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'
        % (NS_A, NS_R, NS_P, SLIDE_W, SLIDE_H))
PRES_RELS = (X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>'
             '</Relationships>')
SLIDE_RELS = (X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
             '</Relationships>')
MASTER = (X + '<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld><p:spTree>'
          '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
          '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
          '</p:spTree></p:cSld>'
          '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
          '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>'
          % (NS_A, NS_R, NS_P))
MASTER_RELS = (X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
              '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
              '</Relationships>')
LAYOUT = (X + '<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" type="blank" preserve="1">'
          '<p:cSld name="Blank"><p:spTree>'
          '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
          '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
          '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
          % (NS_A, NS_R, NS_P))
LAYOUT_RELS = (X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
              '</Relationships>')
CORE = (X + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
        ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
        ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>流程图</dc:title><dc:creator>lark-training-ppt-generator</dc:creator></cp:coreProperties>')
APP = (X + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
       ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
       '<Application>lark-training-ppt-generator</Application></Properties>')
THEME = (X + '<a:theme xmlns:a="%s" name="Office"><a:themeElements>'
         '<a:clrScheme name="Office">'
         '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
         '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
         '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
         '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
         '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
         '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
         '</a:clrScheme>'
         '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:majorFont>'
         '<a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="Office">'
         '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
         '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
         '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
         '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
         '</a:fmtScheme></a:themeElements></a:theme>' % NS_A)


if __name__ == "__main__":
    sys.exit(main())
