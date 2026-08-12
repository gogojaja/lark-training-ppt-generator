#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""generate_page.py — 基于框架的端到端页面生成器

将 style-brief 风格配方 + vertical_rhythm 垂直节奏 + 角色字号驱动，
渲染为真实 PPTX 页面。用于闭环验证设计框架。

== 依赖 ==
    pip install python-pptx

== 用法 ==
    py -3 generate_page.py --recipe ../skills/style-brief-skill/recipes/professional-blue.json \
        --out 生成产物/demo_业务要点.pptx

    py -3 generate_page.py --content my_content.json --out out.pptx

== content JSON 结构 ==
{
  "context": "业务要点 · 个人综合签约",
  "claim":   "签约费每月2元，扣费失败系统25日自动补扣",
  "cards": [
    {"title": "计费规则", "body": "系统每月15日自动扣款，每户每月2元/月签约费。"},
    {"title": "失败重试", "body": "扣费失败于25日再次自动扣款，无需柜员干预。"},
    {"title": "授权要求", "body": "业务代码036101，需业务主管授权办理。"}
  ],
  "meaning": "签约费统一由系统托收，柜员无需手工收费。",
  "source":  "来源：个人综合签约操作手册 v1.2"
}
"""
import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.util import Cm, Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# 让本文件可被直接运行，同时允许被 import
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "skills", "ppt-framework"))
from vertical_rhythm import VerticalRhythm  # noqa: E402

SLIDE_W_CM = 33.867
SLIDE_H_CM = 19.05

# 角色 -> (字号pt, 加粗, 颜色键, 行高)
ROLE_MAP = {
    "sectionLabel": {"size": 12, "bold": False, "color": "secondary", "lh": 1.3},
    "claim":        {"size": 32, "bold": True,  "color": "primary",   "lh": 1.2},
    "body":         {"size": 14, "bold": False, "color": "text",      "lh": 1.5},
    "annotation":   {"size": 12, "bold": False, "color": "text",      "lh": 1.4},
    "source":       {"size": 10, "bold": False, "color": "muted",     "lh": 1.3},
}


def hex2rgb(hexstr):
    hexstr = hexstr.lstrip("#")
    return RGBColor(int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))


class PageBuilder:
    def __init__(self, recipe):
        self.recipe = recipe
        pal = recipe.get("palette", {})
        self.colors = {
            "primary": pal.get("primary", {}).get("value", "1F3864"),
            "secondary": pal.get("secondary", {}).get("value", "2E75B6"),
            "accent": pal.get("accent", {}).get("value", "ED7D31"),
            "text": pal.get("text", {}).get("value", "333333"),
            "muted": "999999",
            "background": pal.get("background", {}).get("value", "FFFFFF"),
            "card": "F8FAFC",
            "cardBorder": "E2E8F0",
        }
        self.vr = VerticalRhythm()
        self.layout = self.vr.get_layout("standard")

    def _band_cm(self, name):
        b = self.layout.get_band(name)
        return b.start_cm, b.height_cm

    def _set_text(self, tf, text, role, align=PP_ALIGN.LEFT, color_override=None):
        cfg = ROLE_MAP[role]
        color = self.colors.get(color_override or cfg["color"], "333333")
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        f = run.font
        f.size = Pt(cfg["size"])
        f.bold = cfg["bold"]
        f.name = "Microsoft YaHei"
        f.color.rgb = hex2rgb(color)
        # 行高在 python-pptx 中通过段落 line_spacing 控制
        p.line_spacing = cfg["lh"]

    def build(self, content, out_path):
        prs = Presentation()
        prs.slide_width = Emu(int(SLIDE_W_CM * 360000))
        prs.slide_height = Emu(int(SLIDE_H_CM * 360000))
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

        # 背景
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = hex2rgb(self.colors["background"])

        # ---- 上下文条带（eyebrow）----
        y0, h0 = self._band_cm("context")
        tb = slide.shapes.add_textbox(Cm(2.5), Cm(y0 + 0.2), Cm(28), Cm(h0 - 0.2))
        self._set_text(tb.text_frame, content.get("context", ""), "sectionLabel")

        # ---- 主张条带（标题）----
        y1, h1 = self._band_cm("claim")
        tb = slide.shapes.add_textbox(Cm(2.5), Cm(y1), Cm(28), Cm(h1))
        tf = tb.text_frame
        tf.word_wrap = True
        self._set_text(tf, content.get("claim", ""), "claim")
        # 主张下方短分隔线（视觉停顿）
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Cm(2.5), Cm(y1 + h1 - 0.15), Cm(6), Cm(0.08))
        line.fill.solid()
        line.fill.fore_color.rgb = hex2rgb(self.colors["accent"])
        line.line.fill.background()

        # ---- 证据条带（卡片网格）----
        y2, h2 = self._band_cm("evidence")
        cards = content.get("cards", [])
        n = len(cards)
        if n == 1:
            cols = 1
        elif n <= 2:
            cols = 2
        else:
            cols = 3
        rows = (n + cols - 1) // cols
        margin_x = 2.5
        gap = 0.6
        total_w = SLIDE_W_CM - margin_x * 2
        card_w = (total_w - gap * (cols - 1)) / cols
        card_h = min((h2 - gap * (rows - 1)) / rows, 5.2)
        start_y = y2 + 0.6
        for i, card in enumerate(cards):
            r = i // cols
            c = i % cols
            x = margin_x + c * (card_w + gap)
            y = start_y + r * (card_h + gap)
            # 卡片面板
            panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                           Cm(x), Cm(y), Cm(card_w), Cm(card_h))
            panel.fill.solid()
            panel.fill.fore_color.rgb = hex2rgb(self.colors["card"])
            panel.line.color.rgb = hex2rgb(self.colors["cardBorder"])
            panel.line.width = Cm(0.03)
            panel.shadow.inherit = False
            # 卡片内容
            ctf = panel.text_frame
            ctf.word_wrap = True
            ctf.margin_left = Cm(0.4)
            ctf.margin_right = Cm(0.4)
            ctf.margin_top = Cm(0.3)
            ctf.vertical_anchor = MSO_ANCHOR.TOP
            # 标题（强调色小标题）
            pt = ctf.paragraphs[0]
            run = pt.add_run()
            run.text = card.get("title", "")
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.name = "Microsoft YaHei"
            run.font.color.rgb = hex2rgb(self.colors["primary"])
            # 正文
            pb = ctf.add_paragraph()
            run = pb.add_run()
            run.text = card.get("body", "")
            run.font.size = Pt(11)
            run.font.name = "Microsoft YaHei"
            run.font.color.rgb = hex2rgb(self.colors["text"])
            pb.line_spacing = 1.4
            pb.space_before = Cm(0.15)

        # ---- 含义条带（底部收尾 takeaway）----
        y3, h3 = self._band_cm("meaning")
        strip = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       Cm(2.5), Cm(y3 + 0.15),
                                       Cm(SLIDE_W_CM - 5), Cm(h3 - 0.3))
        strip.fill.solid()
        strip.fill.fore_color.rgb = hex2rgb(self.colors["primary"])
        strip.line.fill.background()
        strip.shadow.inherit = False
        stf = strip.text_frame
        stf.word_wrap = True
        stf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = stf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = "要点  |  " + content.get("meaning", "")
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.name = "Microsoft YaHei"
        run.font.color.rgb = hex2rgb("FFFFFF")

        # ---- 页脚安全区（来源 + 页码）----
        y4, h4 = self._band_cm("footer")
        ftb = slide.shapes.add_textbox(Cm(2.5), Cm(y4 + 0.1), Cm(SLIDE_W_CM - 5), Cm(h4 - 0.1))
        self._set_text(ftb.text_frame, content.get("source", "来源：内部培训材料"), "source")
        # 页码
        ptb = slide.shapes.add_textbox(Cm(SLIDE_W_CM - 5), Cm(y4 + 0.1), Cm(2.5), Cm(h4 - 0.1))
        self._set_text(ptb.text_frame, "01 / 01", "source", align=PP_ALIGN.RIGHT)

        prs.save(out_path)
        return out_path


def load_recipe(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def default_content():
    return {
        "context": "业务要点 · 个人综合签约",
        "claim": "签约费每月2元，扣费失败系统25日自动补扣",
        "cards": [
            {"title": "计费规则", "body": "系统每月15日自动扣款，每户每月2元/月签约费，统一托收。"},
            {"title": "失败重试", "body": "首次扣费失败，于25日再次自动扣款，无需柜员手工干预。"},
            {"title": "授权要求", "body": "业务代码036101，需业务主管授权方可办理签约与解约。"},
        ],
        "meaning": "签约费由系统统一托收，柜员无需手工收费，降低操作风险。",
        "source": "来源：个人综合签约操作手册 v1.2",
    }


def main():
    ap = argparse.ArgumentParser(description="框架驱动的 PPT 页面生成器")
    ap.add_argument("--recipe", default=os.path.join(
        SCRIPT_DIR, "..", "skills", "style-brief-skill", "recipes", "professional-blue.json"))
    ap.add_argument("--content", help="content JSON 路径（默认使用示例内容）")
    ap.add_argument("--out", default=os.path.join(SCRIPT_DIR, "..", "生成产物", "demo_业务要点.pptx"))
    args = ap.parse_args()

    recipe = load_recipe(args.recipe)
    if args.content:
        with open(args.content, encoding="utf-8") as f:
            content = json.load(f)
    else:
        content = default_content()

    builder = PageBuilder(recipe)
    out = builder.build(content, args.out)
    print("已生成:", os.path.abspath(out))


if __name__ == "__main__":
    main()
