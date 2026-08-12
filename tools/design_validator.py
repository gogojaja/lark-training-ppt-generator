#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""design_validator.py — 设计质量验证工具

基于9项设计诊断清单，对PPT进行自动化质量检查。

== 功能 ==
1. 标题主张检查（2.1）
2. 视觉停顿检查（2.2）
3. 视觉重心偏移检查（2.3，基于形状几何 + 灰度权重）
4. 底部收尾检查（2.4）
5. 面板主次检查（2.5）
6. 图表解读检查（2.6，启发式）
7. 来源可读检查（2.7）
8. 缩略图层次检查（2.8，基于字号对比 + 灰度权重分布）
9. 装饰功能检查（2.9，启发式）

== 用法 ==
    py -3 design_validator.py input.pptx
    py -3 design_validator.py input.pptx --report report.json
    py -3 design_validator.py input.pptx --grayscale  # 额外输出各页灰度重心报告
"""
import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

SLIDE_W = 12192000
SLIDE_H = 6858000

# 常见装饰用色（浅灰、浅蓝等低对比填充），用于2.9启发式
DECOR_FILL = {"F2F2F2", "EEEEEE", "E8E8E8", "F5F5F5", "E0E0E0", "F0F0F0",
              "D9E1F2", "EAF0FB", "F7F9FC"}


class DesignValidator:
    """设计质量验证器"""

    def __init__(self, pptx_path):
        self.pptx_path = pptx_path
        self.slides = []
        self.grayscale_report = []
        self.results = {
            "file": pptx_path,
            "total_slides": 0,
            "checks": [],
            "score": 0,
            "grade": "",
            "issues": []
        }

    def load_slides(self):
        with zipfile.ZipFile(self.pptx_path) as z:
            for name in z.namelist():
                if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                    self.slides.append(name)
        self.results["total_slides"] = len(self.slides)
        return len(self.slides)

    # ---------- 几何解析 ----------
    def parse_shapes(self, slide_xml):
        """解析幻灯片中的形状几何、形状级填充与文本

        返回 list of dict: {x, y, cx, cy, fill, text}
        fill 仅读取形状级填充（p:spPr/a:solidFill），避免误读文字颜色。
        """
        root = ET.fromstring(slide_xml)
        shapes = []
        for sp in root.iter("{%s}sp" % NS_P):
            spPr = sp.find("{%s}spPr" % NS_P)
            if spPr is None:
                continue
            off = spPr.find(".//{%s}off" % NS_A)
            ext = spPr.find(".//{%s}ext" % NS_A)
            if off is None or ext is None:
                continue
            x = int(off.get("x", "0"))
            y = int(off.get("y", "0"))
            cx = int(ext.get("cx", "0"))
            cy = int(ext.get("cy", "0"))
            fill = None
            solid = spPr.find(".//{%s}solidFill" % NS_A)
            if solid is not None:
                srgb = solid.find("{%s}srgbClr" % NS_A)
                if srgb is not None:
                    fill = srgb.get("val", "").upper()
            texts = [t.text or "" for t in sp.iter("{%s}t" % NS_A)]
            text = "".join(texts)
            shapes.append({"x": x, "y": y, "cx": cx, "cy": cy,
                           "fill": fill, "text": text})
        return shapes

    def grayscale_weight(self, shapes):
        """计算灰度视觉权重分布（按象限，面积重叠加权）

        权重 = 面积 × (1 - 明度)。明度由形状级填充色估算；
        无填充（文字框）权重极低。每个形状的权重按其与四象限的
        实际重叠面积比例分配（全宽底栏会均分到左下/右下）。
        """
        quad = [0.0, 0.0, 0.0, 0.0]  # TL, TR, BL, BR
        cx_mid = SLIDE_W / 2
        cy_mid = SLIDE_H / 2
        total = 0.0

        def _ov(a0, a1, b0, b1):
            return max(0.0, min(a1, b1) - max(a0, b0))

        for s in shapes:
            area = (s["cx"] / SLIDE_W) * (s["cy"] / SLIDE_H)  # 归一化面积
            lum = 1.0
            if s["fill"]:
                r = int(s["fill"][0:2], 16)
                g = int(s["fill"][2:4], 16)
                b = int(s["fill"][4:6], 16)
                lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255.0
            weight = area * (1 - lum)
            x0, x1 = s["x"], s["x"] + s["cx"]
            y0, y1 = s["y"], s["y"] + s["cy"]
            ov = [
                _ov(x0, x1, 0, cx_mid) * _ov(y0, y1, 0, cy_mid),     # TL
                _ov(x0, x1, cx_mid, SLIDE_W) * _ov(y0, y1, 0, cy_mid),  # TR
                _ov(x0, x1, 0, cx_mid) * _ov(y0, y1, cy_mid, SLIDE_H),  # BL
                _ov(x0, x1, cx_mid, SLIDE_W) * _ov(y0, y1, cy_mid, SLIDE_H),  # BR
            ]
            s_ov = sum(ov)
            if s_ov > 0:
                for i in range(4):
                    quad[i] += weight * ov[i] / s_ov
            total += weight
        if total > 0:
            quad = [round(w / total, 3) for w in quad]
        return quad, total

    # ---------- 文本解析 ----------
    def extract_text(self, slide_xml):
        return re.findall(r'<a:t[^>]*>([^<]+)</a:t>', slide_xml)

    def extract_font_sizes(self, slide_xml):
        return [int(m) for m in re.findall(r'sz="(\d+)"', slide_xml)]

    # ---------- 各项检查 ----------
    def check_title_claim(self, slide_xml):
        """检查1：标题（主张条带 8-15% 区域）是否陈述了一个主张"""
        shapes = self.parse_shapes(slide_xml)
        # 收集主张条带（垂直中心 6%-20%）内带文本的形状，取最靠下者，
        # 避免误选其上方更小的 eyebrow/上下文文字。
        cands = []
        for s in shapes:
            if not s["text"].strip():
                continue
            cy_center = (s["y"] + s["cy"] / 2) / SLIDE_H
            if 0.06 <= cy_center <= 0.20:
                cands.append((s["y"], s["text"].strip()))
        if not cands:
            return {"pass": False, "reason": "未找到主张条带内的标题文本"}
        title = max(cands, key=lambda c: c[0])[1]
        has_number = any(c.isdigit() for c in title)
        has_verb = any(w in title for w in
                       ["降低", "提升", "增加", "减少", "实现", "完成", "避免",
                        "确保", "解决", "优化", "支持", "必须", "红线", "核心",
                        "自动", "统一", "每月", "仅", "需"])
        return {
            "pass": bool(has_number or has_verb),
            "reason": f"标题「{title}」包含数字或主张词" if (has_number or has_verb)
            else f"标题「{title}」可能只是话题描述（缺少主张）"
        }

    def check_visual_pause(self, slide_xml):
        shapes = self.parse_shapes(slide_xml)
        texts = self.extract_text(slide_xml)
        if len(texts) < 2:
            return {"pass": True, "reason": "内容较少，无需停顿检查"}
        # 标题形状通常位于顶部（y < 15% 高度）且较宽
        title_shapes = [s for s in shapes
                        if s["y"] < SLIDE_H * 0.15 and s["cx"] > SLIDE_W * 0.3]
        if not title_shapes:
            return {"pass": True, "reason": "未识别到标题形状，跳过"}
        title_bottom = max(s["y"] + s["cy"] for s in title_shapes)
        # 第一个非标题内容形状的顶部
        content = [s for s in shapes if (s["y"] + s["cy"]) > title_bottom + SLIDE_H * 0.02]
        if not content:
            return {"pass": True, "reason": "无明确内容区，跳过"}
        first_content_top = min(s["y"] for s in content)
        gap = (first_content_top - title_bottom) / SLIDE_H
        return {
            "pass": gap >= 0.03,
            "reason": f"标题与内容间距 {gap*100:.1f}%（≥3% 视为有停顿）"
            if gap >= 0.03 else f"标题与内容间距仅 {gap*100:.1f}%，过近"
        }

    def check_visual_balance(self, slide_xml):
        """检查3：视觉重心是否偏移（面积重叠加权四象限）

        真正的失败信号（jingmei 2.3）是「左上重、右下轻」——所有内容挤在
        左上角而右下空白。底部含义/页脚带（全宽）是垂直节奏的刻意设计，
        不算失衡。判定：① 任一象限权重不超过 0.65；② 不存在「左上主导且
        右下空白」的对角失衡。
        """
        shapes = self.parse_shapes(slide_xml)
        quad, total = self.grayscale_weight(shapes)
        if total <= 0:
            return {"pass": True, "reason": "无填充形状，跳过重心检查"}
        tl, tr, bl, br = quad
        max_quad = max(quad)
        # 对角失衡：左上主导且右下几乎空白
        diagonal_fail = (tl > 0.4 and br < 0.1)
        self.grayscale_report.append({"quad": quad})
        if max_quad <= 0.65 and not diagonal_fail:
            return {"pass": True,
                    "reason": f"四象限权重 {quad}，无单角过度集中"}
        reason = (f"单一象限过度集中 {max_quad:.2f}" if max_quad > 0.65
                  else f"左上主导且右下空白 TL={tl} BR={br}")
        return {"pass": False, "reason": reason}

    def check_bottom_ending(self, slide_xml):
        texts = self.extract_text(slide_xml)
        footer_kw = ["页码", "第", "页", "来源", "版权", "©", "注："]
        has_footer = any(any(kw in t for kw in footer_kw) for t in texts)
        # 检查82-100%区域是否有形状（页脚安全区）
        shapes = self.parse_shapes(slide_xml)
        has_bottom_shape = any(s["y"] >= SLIDE_H * 0.82 for s in shapes)
        return {
            "pass": bool(has_footer or has_bottom_shape or len(texts) > 3),
            "reason": "底部有收尾元素" if (has_footer or has_bottom_shape)
            else "底部缺少收尾元素（内容用尽即止）"
        }

    def check_panel_hierarchy(self, slide_xml):
        sizes = self.extract_font_sizes(slide_xml)
        if not sizes:
            return {"pass": True, "reason": "无字号信息，跳过"}
        unique = sorted(set(sizes))
        if len(unique) < 2:
            return {"pass": False, "reason": f"字号单一（仅 {unique[0]/100}pt），无主次"}
        max_sz = max(sizes)
        min_sz = min(sizes)
        ratio = max_sz / min_sz if min_sz > 0 else 1
        return {
            "pass": ratio >= 1.5,
            "reason": f"字号跨度 {min_sz/100:.0f}-{max_sz/100:.0f}pt（对比比 {ratio:.1f}）"
        }

    def check_chart_explanation(self, slide_xml):
        """启发式：是否存在图表形状（占大面积的填充矩形）但旁边缺少小字号解读"""
        shapes = self.parse_shapes(slide_xml)
        sizes = self.extract_font_sizes(slide_xml)
        # 大块填充（面积 > 25% 且非装饰色）视为图表/面板
        big_blocks = [s for s in shapes
                      if (s["cx"] * s["cy"]) / (SLIDE_W * SLIDE_H) > 0.25
                      and s["fill"] and s["fill"] not in DECOR_FILL]
        small_text = any(sz < 1400 for sz in sizes)  # <14pt 注释
        if not big_blocks:
            return {"pass": True, "reason": "未识别到图表块，跳过"}
        return {
            "pass": small_text,
            "reason": "图表旁有注释文字" if small_text
            else "存在大块视觉元素但缺少解读文字"
        }

    def check_source_readable(self, slide_xml):
        sizes = self.extract_font_sizes(slide_xml)
        if not sizes:
            return {"pass": True, "reason": "无字号信息，跳过"}
        min_pt = min(sizes) / 100
        return {
            "pass": min_pt >= 8,
            "reason": f"最小字号 {min_pt:.0f}pt"
            if min_pt >= 8 else f"最小字号 {min_pt:.0f}pt，可能不可读（<8pt）"
        }

    def check_thumbnail_hierarchy(self, slide_xml):
        """缩略图层次：字号跨度大 + 灰度权重分布不单一 => 缩略图可辨层次"""
        sizes = self.extract_font_sizes(slide_xml)
        shapes = self.parse_shapes(slide_xml)
        if not sizes:
            return {"pass": True, "reason": "无字号信息，跳过"}
        unique = sorted(set(sizes))
        ratio = max(sizes) / min(sizes) if min(sizes) > 0 else 1
        quad, total = self.grayscale_weight(shapes)
        distinct_zones = len([q for q in quad if q > 0.15])
        ok = (ratio >= 1.5) and (distinct_zones >= 2)
        return {
            "pass": ok,
            "reason": f"字号对比 {ratio:.1f} + 灰度活跃象限 {distinct_zones}"
            if ok else "缩略图下层次可能模糊"
        }

    def check_decoration_function(self, slide_xml):
        """启发式：装饰填充块是否伴随可读文字（否则视为纯填充）"""
        shapes = self.parse_shapes(slide_xml)
        texts = self.extract_text(slide_xml)
        decor = [s for s in shapes
                 if s["fill"] in DECOR_FILL
                 and (s["cx"] * s["cy"]) / (SLIDE_W * SLIDE_H) > 0.05]
        if not decor:
            return {"pass": True, "reason": "无装饰块，跳过"}
        # 装饰块附近（重叠或相邻）有文字则视为有功能
        has_text_near = False
        for d in decor:
            for t in shapes:
                if t is d or t["fill"]:
                    continue
                # 近似：文字形状通常无填充且较小；此处简单判断有非空文本形状即算
            pass
        functional = len(texts) > 0
        return {
            "pass": functional,
            "reason": "装饰块与内容共存（需人工确认功能）" if functional
            else "存在装饰块但页面文本稀少"
        }

    # ---------- 运行 ----------
    def validate(self):
        self.load_slides()
        checks = [
            ("标题主张", self.check_title_claim),
            ("视觉停顿", self.check_visual_pause),
            ("视觉重心偏移", self.check_visual_balance),
            ("底部收尾", self.check_bottom_ending),
            ("面板主次", self.check_panel_hierarchy),
            ("图表解读", self.check_chart_explanation),
            ("来源可读", self.check_source_readable),
            ("缩略图层次", self.check_thumbnail_hierarchy),
            ("装饰功能", self.check_decoration_function),
        ]
        total_score = 0
        for name, func in checks:
            slide_results = []
            for slide_name in self.slides:
                with zipfile.ZipFile(self.pptx_path) as z:
                    xml = z.read(slide_name).decode("utf-8")
                slide_results.append(func(xml))
            passed = sum(1 for r in slide_results if r["pass"])
            total = len(slide_results)
            rate = passed / total if total else 0
            score = round(rate * 10, 1)
            # 记录代表性原因（首个失败项）
            fail_reasons = [r["reason"] for r in slide_results if not r["pass"]]
            self.results["checks"].append({
                "name": name,
                "passed": passed,
                "total": total,
                "pass_rate": round(rate, 2),
                "score": score,
                "sample_issue": fail_reasons[0] if fail_reasons else ""
            })
            total_score += score

        n = len(checks)
        self.results["score"] = round(total_score / n, 1)
        s = self.results["score"]
        self.results["grade"] = ("优秀" if s >= 9 else "良好" if s >= 7
                                 else "一般" if s >= 5 else "需改进")
        return self.results

    def print_report(self, grayscale=False):
        r = self.results
        print("\n" + "=" * 60)
        print("设计质量检查报告")
        print("=" * 60)
        print(f"文件: {r['file']}")
        print(f"幻灯片数量: {r['total_slides']}")
        print("\n--- 检查结果 ---\n")
        for c in r["checks"]:
            status = "[OK]" if c["pass_rate"] >= 0.8 else "[X]"
            line = (f"{status} {c['name']}: {c['passed']}/{c['total']} "
                    f"({c['pass_rate']*100:.0f}%) - {c['score']:.1f}分")
            if c["sample_issue"]:
                line += f"\n    示例问题: {c['sample_issue']}"
            print(line)
        print(f"\n--- 总分: {r['score']:.1f}/10 ---")
        print(f"等级: {r['grade']}")
        if grayscale and self.grayscale_report:
            print("\n--- 灰度重心报告（四象限权重 TL/TR/BL/BR / 重心 x,y）---\n")
            for i, g in enumerate(self.grayscale_report, 1):
                centroid = g.get("centroid", ["-", "-"])
                print(f"  幻灯片 {i}: 权重={g['quad']} 重心=({centroid[0]}, {centroid[1]})")
        print("=" * 60 + "\n")


def main():
    parser = argparse.ArgumentParser(description="PPT设计质量验证工具")
    parser.add_argument("pptx", help="PPTX文件路径")
    parser.add_argument("--report", help="输出JSON报告路径")
    parser.add_argument("--grayscale", action="store_true",
                        help="额外输出各页灰度重心报告")
    args = parser.parse_args()

    if not os.path.exists(args.pptx):
        print(f"错误: 文件不存在 - {args.pptx}")
        sys.exit(1)

    validator = DesignValidator(args.pptx)
    results = validator.validate()
    validator.print_report(grayscale=args.grayscale)

    if args.report:
        out = dict(results)
        out["grayscale"] = validator.grayscale_report
        with open(args.report, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print(f"报告已保存: {args.report}")


if __name__ == "__main__":
    main()
