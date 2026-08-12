#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""vertical_rhythm.py — 垂直节奏布局引擎

基于 jingmei-ppt 方法论，实现五条带垂直节奏布局。

== 功能 ==
1. 计算五条带的位置和高度
2. 提供布局验证
3. 生成布局建议

== 用法 ==
    from vertical_rhythm import VerticalRhythm, SlideLayout
    
    vr = VerticalRhythm()
    layout = vr.get_layout("standard")
    print(layout)
"""

# 画布尺寸（EMU）
SLIDE_W = 12192000  # 33.867 cm
SLIDE_H = 6858000   # 19.05 cm

# 五条带比例
BAND_RATIOS = {
    "context": 0.08,    # 上下文条带 0-8%
    "claim": 0.07,      # 主张条带 8-15%
    "evidence": 0.67,   # 证据区 15-82%
    "meaning": 0.10,    # 含义条带 82-92%
    "footer": 0.08,     # 页脚安全区 92-100%
}

# 间距（EMU）
GAPS = {
    "context_claim": 36200,    # 8px
    "claim_evidence": 54300,   # 12px
    "evidence_meaning": 36200, # 8px
    "meaning_footer": 36200,   # 8px
}

# 页面类型变体
PAGE_VARIANTS = {
    "standard": {
        "name": "标准内容页",
        "bands": ["context", "claim", "evidence", "meaning", "footer"],
    },
    "cover": {
        "name": "封面页",
        "bands": ["context", "claim", "evidence", "meaning", "footer"],
        "overrides": {
            "context": {"height_ratio": 0.0},
            "claim": {"height_ratio": 0.0},
            "evidence": {"height_ratio": 0.82},  # 0-82%
            "meaning": {"height_ratio": 0.10},   # 82-92%
            "footer": {"height_ratio": 0.08},    # 92-100%
        },
    },
    "toc": {
        "name": "目录页",
        "bands": ["context", "claim", "evidence", "meaning", "footer"],
        "overrides": {
            "claim": {"height_ratio": 0.0},
            "evidence": {"height_ratio": 0.84},  # 8-92%
            "meaning": {"height_ratio": 0.0},
            "footer": {"height_ratio": 0.08},    # 92-100%
        },
    },
    "section": {
        "name": "章节首页",
        "bands": ["context", "claim", "evidence", "meaning", "footer"],
        "overrides": {
            "context": {"height_ratio": 0.0},
            "claim": {"height_ratio": 0.22},     # 0-22%
            "evidence": {"height_ratio": 0.70},  # 22-92%
            "meaning": {"height_ratio": 0.0},
            "footer": {"height_ratio": 0.08},    # 92-100%
        },
    },
    "data密集": {
        "name": "数据密集页",
        "bands": ["context", "claim", "evidence", "meaning", "footer"],
    },
}


class Band:
    """条带类"""
    
    def __init__(self, name, start_ratio, height_ratio, content_types=None):
        self.name = name
        self.start_ratio = start_ratio
        self.height_ratio = height_ratio
        self.content_types = content_types or []
        
        # 计算实际位置（EMU）
        self.start_emu = int(SLIDE_H * start_ratio)
        self.height_emu = int(SLIDE_H * height_ratio)
        self.end_emu = self.start_emu + self.height_emu
        
        # 计算实际位置（cm）
        self.start_cm = self.start_emu / 360000
        self.height_cm = self.height_emu / 360000
        self.end_cm = self.end_emu / 360000
    
    def __repr__(self):
        return (f"Band({self.name}: {self.start_ratio*100:.0f}-{(self.start_ratio+self.height_ratio)*100:.0f}%, "
                f"height={self.height_cm:.2f}cm)")


class SlideLayout:
    """幻灯片布局类"""
    
    def __init__(self, variant="standard"):
        self.variant = variant
        self.variant_info = PAGE_VARIANTS.get(variant, PAGE_VARIANTS["standard"])
        self.bands = self._calculate_bands()
    
    def _calculate_bands(self):
        """计算五条带"""
        overrides = self.variant_info.get("overrides", {})
        
        bands = []
        current_ratio = 0.0
        
        for band_name in ["context", "claim", "evidence", "meaning", "footer"]:
            # 获取比例（支持覆盖）
            if band_name in overrides and "height_ratio" in overrides[band_name]:
                height_ratio = overrides[band_name]["height_ratio"]
            else:
                height_ratio = BAND_RATIOS[band_name]
            
            # 创建条带
            band = Band(band_name, current_ratio, height_ratio)
            bands.append(band)
            
            current_ratio += height_ratio
        
        return bands
    
    def get_band(self, name):
        """获取指定条带"""
        for band in self.bands:
            if band.name == name:
                return band
        return None
    
    def validate(self):
        """验证布局"""
        errors = []
        
        # 检查总比例
        total_ratio = sum(b.height_ratio for b in self.bands)
        if abs(total_ratio - 1.0) > 0.01:
            errors.append(f"总比例不为100%: {total_ratio*100:.1f}%")
        
        # 检查主张条带
        claim_band = self.get_band("claim")
        if claim_band and claim_band.height_ratio == 0:
            if self.variant not in ["cover", "toc"]:
                errors.append("主张条带高度为0（非封面/目录页）")
        
        # 检查底部收尾
        footer_band = self.get_band("footer")
        if footer_band and footer_band.height_ratio == 0:
            errors.append("页脚安全区高度为0，缺少底部收尾")
        
        return errors
    
    def to_dict(self):
        """转换为字典"""
        return {
            "variant": self.variant,
            "variant_name": self.variant_info["name"],
            "bands": [
                {
                    "name": b.name,
                    "start_ratio": b.start_ratio,
                    "height_ratio": b.height_ratio,
                    "start_cm": round(b.start_cm, 2),
                    "height_cm": round(b.height_cm, 2),
                    "end_cm": round(b.end_cm, 2),
                }
                for b in self.bands
            ],
        }
    
    def __repr__(self):
        return f"SlideLayout({self.variant}: {self.variant_info['name']})"


class VerticalRhythm:
    """垂直节奏引擎"""
    
    def __init__(self):
        self.slide_width = SLIDE_W
        self.slide_height = SLIDE_H
    
    def get_layout(self, variant="standard"):
        """获取布局"""
        return SlideLayout(variant)
    
    def get_available_variants(self):
        """获取可用的布局变体"""
        return [
            {"id": k, "name": v["name"]}
            for k, v in PAGE_VARIANTS.items()
        ]
    
    def calculate_position(self, band_name, content_height_cm, vertical_align="top"):
        """计算内容在条带中的位置
        
        Args:
            band_name: 条带名称
            content_height_cm: 内容高度（cm）
            vertical_align: 垂直对齐方式（top/center/bottom）
        
        Returns:
            dict: {y: 位置, height: 高度}
        """
        layout = SlideLayout("standard")
        band = layout.get_band(band_name)
        
        if not band:
            return None
        
        band_height_cm = band.height_cm
        
        if content_height_cm > band_height_cm:
            # 内容超出条带高度
            y = band.start_emu
            height = band.height_emu
        else:
            # 根据对齐方式计算位置
            if vertical_align == "center":
                y = band.start_emu + int((band.height_emu - content_height_cm * 360000) / 2)
            elif vertical_align == "bottom":
                y = band.end_emu - int(content_height_cm * 360000)
            else:  # top
                y = band.start_emu
            
            height = int(content_height_cm * 360000)
        
        return {
            "y": y,
            "height": height,
            "band": band_name,
            "overflow": content_height_cm > band_height_cm,
        }
    
    def validate_slide(self, elements):
        """验证幻灯片元素是否符合垂直节奏
        
        Args:
            elements: 元素列表 [{"y": emu, "height": emu, "type": str}, ...]
        
        Returns:
            list: 验证错误列表
        """
        errors = []
        
        for elem in elements:
            y = elem.get("y", 0)
            height = elem.get("height", 0)
            elem_type = elem.get("type", "unknown")
            
            # 计算元素占据的比例
            start_ratio = y / SLIDE_H
            end_ratio = (y + height) / SLIDE_H
            
            # 检查是否超出证据区
            if end_ratio > 0.82 and elem_type not in ["footer", "meaning"]:
                errors.append(f"{elem_type} 超出证据区（结束于 {end_ratio*100:.1f}%）")
            
            # 检查是否与页脚重叠
            if start_ratio < 0.92 and end_ratio > 0.92:
                if elem_type not in ["footer"]:
                    errors.append(f"{elem_type} 与页脚安全区重叠")
        
        return errors


# 快捷函数
def get_layout(variant="standard"):
    """获取布局（快捷函数）"""
    vr = VerticalRhythm()
    return vr.get_layout(variant)


def validate_layout(variant="standard"):
    """验证布局（快捷函数）"""
    layout = get_layout(variant)
    return layout.validate()


def print_layout_info(variant="standard"):
    """打印布局信息（快捷函数）"""
    layout = get_layout(variant)
    print(f"\n=== {layout.variant_info['name']} ===")
    for band in layout.bands:
        print(f"  {band.name:10s}: {band.start_ratio*100:5.1f}% - {(band.start_ratio+band.height_ratio)*100:5.1f}% "
              f"(height={band.height_cm:.2f}cm)")
    print()


if __name__ == "__main__":
    # 测试
    print("=== 垂直节奏布局引擎 ===\n")
    
    # 打印所有布局变体
    vr = VerticalRhythm()
    for variant in vr.get_available_variants():
        print_layout_info(variant["id"])
    
    # 验证布局
    print("=== 布局验证 ===")
    for variant in ["standard", "cover", "toc", "section"]:
        errors = validate_layout(variant)
        if errors:
            print(f"\n{variant}:")
            for err in errors:
                print(f"  - {err}")
        else:
            print(f"\n{variant}: OK")
