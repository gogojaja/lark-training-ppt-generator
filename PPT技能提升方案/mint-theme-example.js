/**
 * 浅绿色主题 PPT 示例（三页）
 * 基于设计令牌 + 组件化架构
 * 
 * 运行前安装依赖：
 *   npm install pptxgenjs
 * 
 * 运行：
 *   node mint-theme-example.js
 * 
 * 页面内容：
 *   1. 封面页
 *   2. 目录页
 *   3. 职能流程图（泳道图）
 */

const pptxgen = require("pptxgenjs");
const { color: C, typography: T, size: S, spacing: SP, roleColors: RC } = require("./theme.js");

// ============================================================
// 角色色板（泳道用）
// ============================================================
const ROLES = [
  { name: "客户",     color: RC.customer },
  { name: "厅堂人员", color: RC.hall },
  { name: "柜员",     color: RC.teller },
  { name: "系统",     color: RC.system },
  { name: "授权人员", color: RC.auth },
];

// ============================================================
// 基础组件库
// ============================================================

/** 统一页眉：页码徽章 + 标题 + 副标题 + 分割线 */
function addHeader(slide, pres, num, title, subtitle) {
  slide.background = { color: C.bg };
  // 顶部装饰线
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: S.slideW, h: S.headerBarH, fill: { color: C.primary },
  });
  // 页码徽章
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: S.margin, y: 0.18, w: S.pageBadgeW, h: S.pageBadgeH,
    fill: { color: C.primary }, rectRadius: 0.05,
  });
  slide.addText(String(num).padStart(2, "0"), {
    x: S.margin, y: 0.18, w: S.pageBadgeW, h: S.pageBadgeH,
    fontSize: T.pageNum, fontFace: T.fontTitle, color: C.darkInk, bold: true,
    align: "center", valign: "middle",
  });
  // 主标题
  slide.addText(title, {
    x: S.margin + 0.52, y: S.claimBandY, w: S.contentW - 1.0, h: S.claimBandH,
    fontSize: T.claim, fontFace: T.fontTitle, color: C.ink, bold: true,
    valign: "middle", fit: "shrink",
  });
  // 副标题
  if (subtitle) {
    slide.addText(subtitle, {
      x: S.margin + 0.52, y: 0.55, w: S.contentW - 1.0, h: 0.24,
      fontSize: 10, fontFace: T.fontBody, color: C.muted,
      valign: "middle", fit: "shrink",
    });
  }
  // 分割线
  slide.addShape(pres.shapes.RECTANGLE, {
    x: S.margin, y: S.dividerY || 0.82, w: S.contentW, h: S.headerDividerH,
    fill: { color: C.primary },
  });
}

/** 统一页脚 */
function addFooter(slide, text) {
  slide.addText(text, {
    x: S.margin, y: S.footerY, w: S.contentW, h: S.footerH,
    fontSize: T.source, fontFace: T.fontBody, color: C.subtle,
    align: "center",
  });
}

/** 信息卡片：左侧装饰条 + 标题 + 内容列表 */
function addInfoCard(slide, pres, x, y, w, h, title, items, opts = {}) {
  const border = opts.border || C.secondary;
  const bg = opts.bg || C.surface;
  const fs = opts.fontSize || T.body;

  // 卡片底
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: bg }, rectRadius: S.cardRadius,
    line: { color: border, width: S.cardBorderW },
  });
  // 左侧装饰条
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y: y + 0.04, w: S.cardAccentW, h: h - 0.08, fill: { color: border },
  });
  // 标题
  if (title) {
    slide.addText(title, {
      x: x + S.cardInnerPad, y: y + 0.06, w: w - S.cardInnerPad - 0.1, h: 0.24,
      fontSize: 11, fontFace: T.fontTitle, color: border, bold: true,
    });
  }
  // 内容
  const bodyY = title ? y + 0.3 : y + 0.08;
  const bodyH = title ? h - 0.36 : h - 0.14;
  const arr = items.map(item => {
    if (typeof item === "string") {
      return { text: item, options: { bullet: true, breakLine: true, paraSpaceAfter: 1 } };
    }
    return { text: item.text, options: { ...item.options, breakLine: true, paraSpaceAfter: 1 } };
  });
  slide.addText(arr, {
    x: x + S.cardInnerPad, y: bodyY, w: w - S.cardInnerPad - 0.1, h: bodyH,
    fontSize: fs, fontFace: T.fontBody, color: C.text,
    valign: "top", fit: "shrink",
  });
}

// ============================================================
// 幻灯片 1：封面页
// ============================================================
function slideCover(pres, data) {
  const s = pres.addSlide();
  s.background = { color: C.dark };

  // 左侧装饰条
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.08, h: S.slideH, fill: { color: C.accent },
  });

  // 顶部细线
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.4, w: 1.5, h: 0.02, fill: { color: C.primaryLight },
  });

  // 主标题
  s.addText(data.title, {
    x: 0.8, y: 1.6, w: 8.4, h: 0.9,
    fontSize: T.deckTitle, fontFace: T.fontTitle, color: C.darkInk, bold: true,
    align: "left", valign: "middle",
  });

  // 副标题
  s.addText(data.subtitle, {
    x: 0.8, y: 2.55, w: 8.4, h: 0.5,
    fontSize: T.deckSubtitle, fontFace: T.fontTitle, color: C.primaryLight,
    align: "left", valign: "middle",
  });

  // 装饰线
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 3.2, w: 2.5, h: 0.03, fill: { color: C.accent },
  });

  // 描述行
  s.addText(data.description, {
    x: 0.8, y: 3.4, w: 8.4, h: 0.35,
    fontSize: 13, fontFace: T.fontBody, color: C.darkMuted,
    align: "left", valign: "middle",
  });

  // 元信息行
  s.addText(data.meta, {
    x: 0.8, y: 4.3, w: 8.4, h: 0.3,
    fontSize: 11, fontFace: T.fontBody, color: C.darkSubtle,
    align: "left", valign: "middle",
  });

  s.addNotes(data.notes || "");
}

// ============================================================
// 幻灯片 2：目录页
// ============================================================
function slideToc(pres, data) {
  const s = pres.addSlide();
  addHeader(s, pres, 2, data.title, data.subtitle);

  const items = data.items;
  const n = items.length;

  const gap = 0.08;
  const cardW = (S.contentW - gap * (n - 1)) / n;
  const cardH = 3.4;
  const startX = S.margin;
  const startY = S.evidenceY + 0.1;

  items.forEach((it, i) => {
    const x = startX + i * (cardW + gap);

    // 卡片底
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: startY, w: cardW, h: cardH,
      fill: { color: C.surface }, rectRadius: S.cardRadius,
      line: { color: C.secondary, width: S.cardBorderW },
    });
    // 顶部装饰条
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: startY, w: cardW, h: 0.05, fill: { color: C.primary },
    });
    // 序号
    s.addText(it.num, {
      x: x + 0.05, y: startY + 0.2, w: cardW - 0.1, h: 0.45,
      fontSize: 24, fontFace: T.fontTitle, color: C.primary, bold: true,
      align: "center",
    });
    // 标题
    s.addText(it.title, {
      x: x + 0.05, y: startY + 0.7, w: cardW - 0.1, h: 0.55,
      fontSize: 10, fontFace: T.fontTitle, color: C.ink, bold: true,
      align: "center", valign: "top", fit: "shrink",
    });
    // 描述
    s.addText(it.desc, {
      x: x + 0.05, y: startY + 1.35, w: cardW - 0.1, h: 1.9,
      fontSize: T.annotation, fontFace: T.fontBody, color: C.muted,
      align: "center", valign: "top", fit: "shrink",
    });
  });

  addFooter(s, data.footerText);
}

// ============================================================
// 幻灯片 3：职能流程图（泳道图）
// ============================================================
function slideSwimlane(pres, data) {
  const s = pres.addSlide();
  addHeader(s, pres, 3, data.title, data.subtitle);

  const lanes = data.lanes;
  const nodes = data.nodes;
  const arrows = data.arrows;

  const laneStartY = S.evidenceY;
  const laneH = 0.62;
  const labelW = 0.95;
  const procX = S.margin + labelW;
  const procW = S.slideW - S.margin - procX;

  // 绘制泳道
  lanes.forEach((ln, i) => {
    const y = laneStartY + i * laneH;
    // 流程区背景（交替色）
    s.addShape(pres.shapes.RECTANGLE, {
      x: procX, y, w: procW, h: laneH,
      fill: { color: i % 2 === 0 ? C.surface : C.bg },
    });
    // 泳道标签
    s.addShape(pres.shapes.RECTANGLE, {
      x: S.margin, y, w: labelW, h: laneH,
      fill: { color: ln.color },
    });
    s.addText(ln.name, {
      x: S.margin, y, w: labelW, h: laneH,
      fontSize: 10, fontFace: T.fontTitle, color: C.darkInk, bold: true,
      align: "center", valign: "middle",
    });
    // 分隔线
    if (i < lanes.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: S.margin, y: y + laneH, w: S.slideW - 2 * S.margin, h: 0,
        line: { color: C.hairline, width: 0.5 },
      });
    }
  });

  // 绘制节点
  const nodeW = 1.05, nodeH = 0.45;
  nodes.forEach(nd => {
    const nx = procX + nd.x;
    const ny = laneStartY + nd.lane * laneH + (laneH - nodeH) / 2;
    const fillColor = lanes[nd.lane].color;

    if (nd.dashed) {
      s.addShape(pres.shapes.OVAL, {
        x: nx, y: ny, w: nodeW, h: nodeH,
        fill: { color: fillColor },
        line: { color: C.risk, width: 1.5, dashType: "dash" },
      });
    } else {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: nx, y: ny, w: nodeW, h: nodeH,
        fill: { color: fillColor }, rectRadius: 0.04,
        line: { color: fillColor, width: 1 },
      });
    }

    s.addText(nd.label, {
      x: nx, y: ny, w: nodeW, h: nodeH,
      fontSize: 7, fontFace: T.fontBody, color: C.darkInk, bold: true,
      align: "center", valign: "middle", fit: "shrink",
    });
  });

  // 绘制箭头（折线）
  arrows.forEach(a => {
    const x1 = procX + a.from.x + nodeW;
    const y1 = laneStartY + a.from.lane * laneH + laneH / 2;
    const x2 = procX + a.to.x;
    const y2 = laneStartY + a.to.lane * laneH + laneH / 2;
    const midX = (x1 + x2) / 2;

    // 水平段1
    s.addShape(pres.shapes.LINE, {
      x: x1, y: y1, w: midX - x1, h: 0,
      line: { color: C.accent, width: 1.5 },
    });
    // 垂直段
    s.addShape(pres.shapes.LINE, {
      x: midX, y: Math.min(y1, y2), w: 0, h: Math.abs(y2 - y1),
      line: { color: C.accent, width: 1.5 },
    });
    // 水平段2（带箭头）
    s.addShape(pres.shapes.LINE, {
      x: midX, y: y2, w: x2 - midX, h: 0,
      line: { color: C.accent, width: 1.5, endArrowType: "triangle" },
    });
  });

  // 图例
  const legY = 4.4;
  s.addText("图例：", {
    x: S.margin, y: legY, w: 0.5, h: 0.25,
    fontSize: 9, fontFace: T.fontTitle, color: C.ink, bold: true,
    valign: "middle",
  });
  lanes.forEach((ln, i) => {
    const lx = S.margin + 0.55 + i * 1.0;
    s.addShape(pres.shapes.RECTANGLE, {
      x: lx, y: legY + 0.04, w: 0.18, h: 0.18,
      fill: { color: ln.color },
    });
    s.addText(ln.name, {
      x: lx + 0.22, y: legY, w: 0.7, h: 0.25,
      fontSize: 9, fontFace: T.fontBody, color: C.text, valign: "middle",
    });
  });
  // 条件节点图例
  const dashLx = S.margin + 0.55 + lanes.length * 1.0 + 0.5;
  s.addShape(pres.shapes.OVAL, {
    x: dashLx, y: legY + 0.04, w: 0.18, h: 0.18,
    fill: { color: C.risk },
    line: { color: C.risk, width: 1.5, dashType: "dash" },
  });
  s.addText("条件触发节点", {
    x: dashLx + 0.22, y: legY, w: 1.5, h: 0.25,
    fontSize: 9, fontFace: T.fontBody, color: C.text, valign: "middle",
  });

  addFooter(s, data.footerText);
}

// ============================================================
// 组装：生成 PPT
// ============================================================
function buildPPT() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Training Team";
  pres.title = "个人客户信息维护 · 操作培训手册";

  // ---- 数据：封面 ----
  slideCover(pres, {
    title: "个人客户信息维护",
    subtitle: "操作培训手册",
    description: "柜面 · 智能柜员机 · 移动Pad  全渠道操作培训",
    meta: "交易代码 030401  |  权限：业务柜员",
    notes: "【开场白】各位同事大家好，今天我们来学习个人客户信息维护的操作流程...",
  });

  // ---- 数据：目录 ----
  slideToc(pres, {
    title: "培训内容概览",
    subtitle: "目录",
    footerText: "个人客户信息维护 · 操作培训手册",
    items: [
      { num: "01", title: "场景说明", desc: "场景定义、支持渠道与业务范围" },
      { num: "02", title: "业务办理总览", desc: "完整操作流程，7个关键环节" },
      { num: "03", title: "职能流程图", desc: "跨角色泳道图，清晰展现协作关系" },
      { num: "04", title: "柜面操作详解", desc: "8步操作流程，含代理办理规则" },
      { num: "05", title: "智能柜员机操作", desc: "6步自助操作流程" },
      { num: "06", title: "移动Pad操作", desc: "7步移动办理流程" },
      { num: "07", title: "业务规则与字段", desc: "核心规则、证件类型、凭证管理" },
    ],
  });

  // ---- 数据：职能流程图 ----
  slideSwimlane(pres, {
    title: "职能流程图",
    subtitle: "按角色分色泳道图 · 客户/厅堂人员/柜员/系统/授权人员",
    footerText: "个人客户信息维护 · 操作培训手册",
    lanes: ROLES,
    nodes: [
      { label: "客户到达\n网点", lane: 0, x: 0.1 },
      { label: "身份识别\n与分流", lane: 1, x: 1.25 },
      { label: "联网核查\n人脸识别", lane: 2, x: 2.4 },
      { label: "三要素比对\n信息维护", lane: 3, x: 3.55 },
      { label: "手机号核查\n验证码验证", lane: 3, x: 4.7 },
      { label: "智能授权\n(税收居民\n变化)", lane: 4, x: 5.85, dashed: true },
      { label: "签名确认\n打印评价\n归档", lane: 2, x: 7.0 },
    ],
    arrows: [
      { from: { lane: 0, x: 0.1 }, to: { lane: 1, x: 1.25 } },
      { from: { lane: 1, x: 1.25 }, to: { lane: 2, x: 2.4 } },
      { from: { lane: 2, x: 2.4 }, to: { lane: 3, x: 3.55 } },
      { from: { lane: 3, x: 3.55 }, to: { lane: 3, x: 4.7 } },
      { from: { lane: 3, x: 4.7 }, to: { lane: 4, x: 5.85 } },
      { from: { lane: 4, x: 5.85 }, to: { lane: 2, x: 7.0 } },
    ],
  });

  // 保存
  pres.writeFile({ fileName: "浅绿色主题示例.pptx" })
    .then(fn => console.log("✅ 生成成功:", fn))
    .catch(err => console.error("❌ 生成失败:", err));
}

buildPPT();
