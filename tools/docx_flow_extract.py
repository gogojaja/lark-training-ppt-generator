#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_flow_extract.py - 读取 Word 文档中的流程图并转换为 PPT 流程图

从 .docx 中识别/读取三类流程图载体，并复用现有 CSV → PPT 流水线生成幻灯片：

  1. SmartArt（word/diagrams/*.xml）  —— 精确解析节点层级，无需 OCR
  2. 矢量图形（DrawingML <a:sp> + <a:cxnSp> 连接线）—— 精确解析节点与连接拓扑
  3. 位图图片（截图/识图）           —— 本地离线 OCR（Windows 内置 WinRT OCR，无需联网/安装）
                                        + 行聚类 + 排序，输出【草稿】CSV

输出与 flowchart-skill 相同的「全参数 CSV」（green 预设），并调用:
    生成脚本/csv_to_flowchart.py 生成 PPTX。
纯 Python 标准库 + 内置 PowerShell OCR 辅助脚本，无第三方依赖。

用法:
    py -3 tools/docx_flow_extract.py --scan 文档.docx                  # 只盘点载体形式
    py -3 tools/docx_flow_extract.py --smartart 文档.docx [--out 前缀]
    py -3 tools/docx_flow_extract.py --shapes  文档.docx [--out 前缀]
    py -3 tools/docx_flow_extract.py --image  文档.docx [--media image111.png | --rid rId122] [--out 前缀]
    py -3 tools/docx_flow_extract.py --auto   文档.docx [--out 前缀]  # 自动选择最优载体
    # 通用选项: --preset green --skip-ppt --json-only --ocr-lang zh-Hans-CN
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS_DIR = os.path.join(REPO_ROOT, "tools")
GENERATE_DIR = os.path.join(REPO_ROOT, "生成脚本")
DEFAULT_OUT = os.path.join(REPO_ROOT, "生成产物", "流程图")
OCR_HELPER = os.path.join(TOOLS_DIR, "win_ocr.ps1")
CSV_CONVERTER = os.path.join(GENERATE_DIR, "csv_to_flowchart.py")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}

JUDGE_KEYWORDS = ("是否", "判断", "确认", "校验", "通过", "审核", "核实",
                  "有效", "过期", "正确", "错误", "需要", "可否", "验证")

UI_CHROME = re.compile(
    r"(今天|周[一二三四五六日]|星期|年|月|日|秒|分钟|流水号|机构号|设备编号|"
    r"呼叫|排队|队列|叫号|签到|签退|登录|退出|主页|首页|返回|退出登录|角色|权限)")


# ---------------- docx 内部解析 ----------------

def open_docx(docx_path):
    try:
        return zipfile.ZipFile(docx_path)
    except zipfile.BadZipFile:
        raise SystemExit("无法读取 docs：%s（不是有效的 Word 文件）" % docx_path)


def read_xml_bytes(z, name):
    try:
        return z.read(name)
    except KeyError:
        return None


def _text_of_pt(pt_elem):
    """SmartArt 节点文本：取 <dgm:t>，排重嵌套子节点内的文本"""
    texts = []
    stack = [pt_elem]
    while stack:
        e = stack.pop(0)
        if e.tag == Q("dgm:t"):
            texts.append(e.text or "")
            continue
        for child in e:
            if child.tag == Q("dgm:pt"):
                continue
            stack.append(child)
    return "".join(texts).strip()


def Q(name):
    return "{%s}%s" % (NS[name.split(":")[0]], name.split(":")[1])


def scan_representations(docx_path):
    """盘点 docx 中的流程图载体：SmartArt / 矢量形状 / 图片"""
    z = open_docx(docx_path)
    try:
        names = z.namelist()
        doc = (z.read("word/document.xml") or b"").decode("utf-8", errors="replace")
        rels = (z.read("word/_rels/document.xml.rels") or b"").decode("utf-8", errors="replace")
        relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))

        diagrams = sorted(n for n in names
                          if n.startswith("word/diagrams/") and n.endswith(".xml"))
        cnxs = doc.count("a:cxnSp")
        vml = doc.count("<w:pict")

        images = []
        for block in re.finditer(
                r"<wp:(?:inline|anchor)[^>]*>.*?</wp:(?:inline|anchor)>", doc, re.S):
            blk = block.group(0)
            m = re.search(r'r:embed="(rId\d+)"', blk)
            if not m:
                continue
            ext = re.search(r'<wp:extent cx="([\d.]+)" cy="([\d.]+)"', blk)
            w_cm = round(float(ext.group(1)) / 360000, 1) if ext else 0
            h_cm = round(float(ext.group(2)) / 360000, 1) if ext else 0
            images.append({
                "rid": m.group(1),
                "media": os.path.basename(relmap.get(m.group(1), "")),
                "w_cm": w_cm,
                "h_cm": h_cm,
            })
        return {"diagrams": diagrams, "connectors": cnxs, "vml": vml, "images": images}
    finally:
        z.close()


# ---------------- 1. SmartArt ----------------

def extract_smartart(docx_path):
    z = open_docx(docx_path)
    try:
        diagram_files = sorted(n for n in z.namelist()
                               if n.startswith("word/diagrams/data") and n.endswith(".xml"))
        flows = []
        for df in diagram_files:
            xml = z.read(df)
            root = ET.fromstring(xml)
            ptLst = root.find("./dgm:ptLst", NS)
            if ptLst is None:
                continue
            # 收集全部 pt 元素（含嵌套），保留父子关系
            nodes = []          # (modelId, text, type)
            children = {}       # modelId -> [child modelId]
            all_pts = list(ptLst.iter(Q("dgm:pt")))
            for pt in all_pts:
                mid = pt.get("modelId")
                typ = pt.get("type", "node")
                text = _text_of_pt(pt)
                # 嵌套子节点
                sub_ids = []
                for sub in pt.findall(Q("dgm:pt")):
                    sm = sub.get("modelId")
                    if sm is not None:
                        sub_ids.append(sm)
                nodes.append((mid, text, typ))
                children.setdefault(mid, []).extend(sub_ids)
                # to/from 引用关系（扁平结构）
                if typ != "doc":
                    for to in pt.findall(Q("dgm:to")):
                        tid = to.get("pt")
                        if tid:
                            children.setdefault(mid, []).append(tid)

            roots = [m for m, _, typ in nodes if typ == "doc"]
            if not roots and nodes:
                roots = [nodes[0][0]]

            node_map = {m: (t, typ) for m, t, typ in nodes}

            # 层次序遍历得到节点顺序
            order = []
            seen = set()
            stack = list(roots)
            while stack:
                m = stack.pop(0)
                if m in seen or m not in node_map:
                    continue
                seen.add(m)
                order.append(m)
                for c in reversed(children.get(m, [])):
                    stack.insert(0, c)
            ordered = [node_map[m][0] for m in order if node_map[m][1] != "doc"]
            flows.append({"file": df, "nodes": ordered})
        return flows
    finally:
        z.close()


# ---------------- 2. DrawingML 矢量形状 ----------------

def extract_shapes(docx_path):
    z = open_docx(docx_path)
    try:
        doc = (z.read("word/document.xml") or b"").decode("utf-8", errors="replace")
    finally:
        z.close()

    shapes = []   # {id, x, y, text}
    connectors = []  # {from_id, to_id}
    num_re = re.compile(r'\bid="(\d+)"')
    pos_re = re.compile(r'<a:off x="([\d.-]+)" y="([\d.-]+)"')

    for m in re.finditer(r"<a:sp>.*?</a:sp>", doc, re.S):
        block = m.group(0)
        idm = num_re.search(block)
        texts = re.findall(r"<a:t[^>]*>([^<]*)</a:t>", block)
        text = "".join(texts).strip()
        ppm = pos_re.search(block)
        shapes.append({
            "id": idm.group(1) if idm else len(shapes),
            "x": int(ppm.group(1)) if ppm else 0,
            "y": int(ppm.group(2)) if ppm else 0,
            "text": text,
        })

    for m in re.finditer(r"<a:cxnSp>.*?</a:cxnSp>", doc, re.S):
        block = m.group(0)
        st = re.search(r'<a:stCxn id="(\d+)"', block)
        en = re.search(r'<a:endCxn id="(\d+)"', block)
        if st and en:
            connectors.append((st.group(1), en.group(1)))

    # 排序：有连接线则拓扑（从无入边的节点出发），否则按 (y, x)
    if connectors:
        by_id = {}
        for s in shapes:
            by_id.setdefault(str(s["id"]), []).append(s)
        indeg = {}
        adj = {}
        for a, b in connectors:
            indeg[b] = indeg.get(b, 0) + 1
            adj.setdefault(a, []).append(b)
        starts = [s for s in shapes if indeg.get(str(s["id"]), 0) == 0]
        order_ids = []
        for s in sorted(starts, key=lambda s: (s["y"], s["x"])):
            stack = [str(s["id"])]
            while stack:
                cur = stack.pop(0)
                if cur in order_ids:
                    continue
                order_ids.append(cur)
                for nxt in sorted(adj.get(cur, []), key=lambda i2: by_id.get(i2, [{}])[0].get("y", 0)):
                    stack.insert(0, nxt)
        id_set = set(order_ids)
        # 补漏
        for s in shapes:
            if str(s["id"]) not in id_set:
                order_ids.append(str(s["id"]))
        idx = {i: n for n, i in enumerate(order_ids)}
        ordered = sorted(shapes, key=lambda s: idx.get(str(s["id"]), 10 ** 9))
    else:
        ordered = sorted(shapes, key=lambda s: (s["y"], s["x"]))

    return [s["text"] for s in ordered if s["text"]]


# ---------------- 3. 位图图片（OCR） ----------------

def extract_image_file(docx_path, media_target, workdir):
    """抽取 docx 内 media 文件到 workdir，返回本地路径"""
    z = open_docx(docx_path)
    try:
        rels = (z.read("word/_rels/document.xml.rels") or b"").decode("utf-8", errors="replace")
        relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
        if media_target.startswith("rId"):
            target = "word/" + relmap.get(media_target, "")
        else:
            target = "word/media/" + media_target
        data = z.read(target)
        local = os.path.join(workdir, os.path.basename(target))
        with open(local, "wb") as f:
            f.write(data)
        return local
    finally:
        z.close()


def run_ocr(image_path, out_json, ocr_lang):
    cmd = [
        "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", OCR_HELPER,
        "-ImagePath", image_path,
        "-OutJson", out_json,
        "-Lang", ocr_lang,
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")
    if proc.returncode != 0:
        raise SystemExit("OCR 失败：\n" + (proc.stderr or proc.stdout))
    with open(out_json, encoding="utf-8-sig") as f:
        return json.load(f)


def group_lines(words):
    """OCR words → 文本行（含 bbox）"""
    ws = sorted(words, key=lambda w: (w["y"], w["x"]))
    lines = []
    for w in ws:
        cy0, ce0 = w["y"], w["y"] + w["h"]
        for line in lines:
            if max(cy0, line["y"]) < min(ce0, line["ye"]):
                line["words"].append(w)
                line["y"] = min(cy0, line["y"])
                line["ye"] = max(ce0, line["ye"])
                break
        else:
            lines.append({"y": cy0, "ye": ce0, "words": [w]})
    for ln in lines:
        ln["words"].sort(key=lambda w: w["x"])
        ln["text"] = "".join(w["text"] for w in ln["words"]).strip()
        ln["x"] = min(w["x"] for w in ln["words"])
        ln["xc"] = max(w["x"] + w["w"] for w in ln["words"])
        ln["yc"] = (ln["y"] + ln["ye"]) / 2
    lines.sort(key=lambda l: (l["y"], l["x"]))
    return lines


def cluster_boxes(lines, image_h):
    """文本行 → 流程框（按纵向间隔 + 横向重叠聚类）"""
    if not lines:
        return []
    avg_h = sum(l["ye"] - l["y"] for l in lines) / len(lines)
    gap = max(avg_h * 1.8, 12)
    boxes = []
    for ln in lines:
        for box in boxes:
            if abs(ln["yc"] - box["yc"]) < gap and (ln["x"] < box["xc"] and ln["xc"] > box["x"]):
                box["lines"].append(ln)
                break
        else:
            boxes.append({"lines": [ln], "y": ln["y"], "ye": ln["ye"],
                          "x": ln["x"], "xc": ln["xc"], "yc": ln["yc"]})
    for b in boxes:
        b["y"] = min(l["y"] for l in b["lines"])
        b["ye"] = max(l["ye"] for l in b["lines"])
        b["x"] = min(l["x"] for l in b["lines"])
        b["xc"] = max(l["xc"] for l in b["lines"])
        b["yc"] = (b["y"] + b["ye"]) / 2
        b["text"] = " ".join(l["text"] for l in sorted(b["lines"], key=lambda l: l["y"]))
    boxes.sort(key=lambda b: (b["yc"], b["x"]))
    return boxes


def is_ui_chrome(text):
    if not text:
        return True
    if re.search(r"\d{1,2}[:：]\d{2}", text):   # 时间戳
        return True
    if len(text) <= 1 and not re.search(r"[\u4e00-\u9fa5]", text):
        return True
    if UI_CHROME.search(text):
        return True
    return False


def ocr_to_nodes(ocr_json):
    """OCR 结果 → 顺序节点（rect/diamond 尽力识别）"""
    img_h = ocr_json.get("imageHeight", 1000)
    lines = group_lines(ocr_json.get("words", []))
    boxes = cluster_boxes(lines, img_h)
    nodes = []
    seen = set()
    for b in boxes:
        text = re.sub(r"\s+", "", b["text"])
        if is_ui_chrome(text) or text in seen:
            continue
        seen.add(text)
        nodes.append({"text": text, "shape": "diamond"
                      if any(k in text for k in JUDGE_KEYWORDS) else "rect"})
    return nodes


# ---------------- CSV + PPT 生成 ----------------

PRESET_COLORS = {
    "green": {"main": ("C6EFCE", "006100"), "diamond": ("FFF2CC", "7F6000"),
              "branch": ("DDEBF7", "1F3864"), "error": ("FCE4EC", "C00000"),
              "title_bg": "1F3864"},
    "blue": {"main": ("D6E4F0", "1F3864"), "diamond": ("FFF2CC", "7F6000"),
             "branch": ("E2EFDA", "375623"), "error": ("FCE4EC", "C00000"),
             "title_bg": "1F3864"},
}


def write_flow_csv(nodes, title, out_csv, preset="green"):
    colors = PRESET_COLORS.get(preset, PRESET_COLORS["green"])
    rows = ["type,key,value,desc,seq,node_type,content,shape,width_cm,height_cm,bg_color,text_color,branch_to,branch_label,branch_kind"]
    cfg = [
        ("title", title, "流程图标题"),
        ("preset", preset, "配色预设"),
        ("no_connectors", "false", "禁用连接线"),
        ("step_gap_cm", "1.1", "纵向间隔"),
        ("box_width_cm", "5.0", "矩形宽"),
        ("box_height_cm", "0.7", "矩形高"),
        ("diamond_width_cm", "4.8", "菱形宽"),
        ("diamond_height_cm", "1.1", "菱形高"),
        ("title_bg", colors["title_bg"], "标题背景"),
        ("title_text", "FFFFFF", "标题文字色"),
    ]
    for k, v, desc in cfg:
        rows.append("config,%s,%s,%s,,,,,,,,,,," % (k, v, desc))
    seq = 0
    for nd in nodes:
        seq += 1
        is_dia = nd["shape"] == "diamond"
        bg, tc = colors["diamond"] if is_dia else colors["main"]
        w = 4.8 if is_dia else 5.0
        h = 1.1 if is_dia else 0.7
        rows.append(",,,,%d,main,%s,%s,%s,%s,%s,%s,,," % (
            seq, nd["text"][:20], nd["shape"], w, h, bg, tc))
    with open(out_csv, "w", encoding="utf-8-sig", newline="") as f:
        f.write("\n".join(rows) + "\n")
    return out_csv


def make_ppt(csv_path, out_pptx):
    cmd = [sys.executable, CSV_CONVERTER, csv_path, "--out", out_pptx]
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", env=env)
    if proc.returncode != 0:
        raise SystemExit("csv_to_flowchart.py 失败：\n" + (proc.stderr or proc.stdout))
    return out_pptx


# ---------------- CLI ----------------

def build_output_paths(name, out_dir):
    base = re.sub(r"[\\/:*?\"<>|\s]+", "_", name)
    return (os.path.join(out_dir, base + ".csv"),
            os.path.join(out_dir, base + ".pptx"))


def emit(nodes, title, out_dir, preset, skip_ppt, json_only):
    if not nodes:
        print("未提取到任何流程节点。")
        return
    json_path = os.path.join(out_dir, title + "_wordflow.json")
    os.makedirs(out_dir, exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({"title": title, "nodes": nodes}, f, ensure_ascii=False, indent=2)
    print("节点 JSON: %s（%d 个节点）" % (json_path, len(nodes)))
    for i in range(0, len(nodes), 30):
        for nd in nodes[i:i + 30]:
            print("  %-4s %s" % ("◇" if nd["shape"] == "diamond" else "□", nd["text"]))
        if i + 30 < len(nodes):
            print("  ...（续）")

    if json_only:
        return
    csv_path, pptx_path = build_output_paths(title + "_wordflow", out_dir)
    write_flow_csv(nodes, title, csv_path, preset)
    print("CSV 节点表: %s" % csv_path)
    if skip_ppt:
        return
    make_ppt(csv_path, pptx_path)
    print("PPT 流程图: %s" % pptx_path)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", help="输入 Word 文档 (.docx)")
    ap.add_argument("--scan", action="store_true", help="只盘点载体")
    ap.add_argument("--smartart", action="store_true", help="提取 SmartArt")
    ap.add_argument("--shapes", action="store_true", help="提取矢量形状")
    ap.add_argument("--image", action="store_true", help="OCR 提取图片")
    ap.add_argument("--auto", action="store_true", help="自动选择载体（SmartArt>形状>图片）")
    ap.add_argument("--media", help="图片媒体文件名，如 image111.png")
    ap.add_argument("--rid", help="图片 rId，如 rId122")
    ap.add_argument("--out", dest="out_prefix", help="输出文件名前缀")
    ap.add_argument("--out-dir", default=DEFAULT_OUT, help="输出目录")
    ap.add_argument("--preset", default="green")
    ap.add_argument("--skip-ppt", action="store_true", help="只生成 CSV/JSON")
    ap.add_argument("--json-only", action="store_true", help="只输出 JSON")
    ap.add_argument("--ocr-lang", default="zh-Hans-CN", help="OCR 语言标签")
    args = ap.parse_args()

    if not os.path.isfile(args.docx):
        raise SystemExit("文件不存在: %s" % args.docx)
    base = os.path.splitext(os.path.basename(args.docx))[0]

    if args.scan:
        info = scan_representations(args.docx)
        print("清单: %s" % args.docx)
        print("  SmartArt 图:     %d" % len(info["diagrams"]))
        print("  矢量连接线:      %d" % info["connectors"])
        print("  VML 图形:        %d" % info["vml"])
        print("  插入图片:        %d" % len(info["images"]))
        for im in info["images"][:20]:
            print("      %-8s %-16s %4.1fcm x %4.1fcm"
                  % (im["rid"], im["media"], im["w_cm"], im["h_cm"]))
        return

    mode = None
    for m in ("smartart", "shapes", "image", "auto"):
        if getattr(args, m):
            mode = m
            break
    if not mode:
        ap.print_help()
        raise SystemExit("\n请指定模式：--smartart / --shapes / --image / --auto / --scan")

    if mode in ("smartart", "auto"):
        flows = extract_smartart(args.docx)
        if mode == "smartart" or (mode == "auto" and flows):
            if not flows:
                print("未检测到 SmartArt。")
            for i, f in enumerate(flows):
                title = "%s_SmartArt%d" % (base, i)
                emit(f["nodes"], title, args.out_dir,
                     args.preset, args.skip_ppt, args.json_only)
            return

    if mode in ("shapes", "auto"):
        nodes = extract_shapes(args.docx)
        if nodes:
            title = base + "_形状"
            emit(nodes, title, args.out_dir,
                 args.preset, args.skip_ppt, args.json_only)
            return

    if mode in ("image", "auto"):
        info = scan_representations(args.docx)
        target = args.media or args.rid
        cands = info["images"]
        if not cands:
            print("未检测到插入图片。")
            return
        if target:
            picks = [im for im in cands if im["media"] == target or im["rid"] == target]
        else:
            picks = [im for im in cands if im["w_cm"] >= 6]
            if not picks:
                picks = cands
        workdir = tempfile.mkdtemp(prefix="wordflow_")
        try:
            for im in picks:
                img = extract_image_file(args.docx, im["rid"], workdir)
                ocr_json = os.path.join(workdir, "ocr.json")
                print("OCR: %s (%s, %.1fcm x %.1fcm) ..." % (
                    im["media"], im["rid"], im["w_cm"], im["h_cm"]), flush=True)
                data = run_ocr(img, ocr_json, args.ocr_lang)
                nodes = ocr_to_nodes(data)
                title = base + "_" + os.path.splitext(im["media"])[0]
                emit(nodes, title, args.out_dir,
                     args.preset, args.skip_ppt, args.json_only)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)
        return

    print("未提取到任何内容。可尝试 --scan 查看载体。")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)