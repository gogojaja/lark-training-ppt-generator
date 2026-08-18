const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.author = "AI Assistant";
pres.title = "个人客户信息维护 · 操作培训手册";
pres.subject = "操作培训";

// ============================================================
// DIMENSIONS
// ============================================================
pres.layout = "LAYOUT_16x9";
const W = 10, H = 5.625;
const M = 0.5;
const CW = W - 2 * M; // content width = 9

// ============================================================
// COLORS — Formal Work Report (Blue + Amber Accent)
// ============================================================
const C = {
  primary: "27AE60",      // 柔和中绿（主色/标题/表头）
  secondary: "5DADE2",    // 柔和浅蓝（辅色）
  dark: "1E5631",         // 深绿（封面/深色块背景）
  accent: "F4D03F",       // 暖黄强调色
  bg: "F5FFF7",           // 极浅绿背景
  white: "FFFFFF",
  text: "2C3E2F",         // 深灰绿主文本
  textLight: "85998A",    // 浅灰绿次要文本
  success: "58D68D",      // 浅绿
  warning: "F39C12",      // 暖橙黄（保证可读性）
  danger: "EC7063",       // 柔和浅红
  cardBg: "FFFFFF",
  // 角色颜色（泳道图）
  roleCustomer: "2ECC71", // 绿
  roleHall: "F4D03F",     // 黄
  roleTeller: "3498DB",   // 蓝
  roleSystem: "85C1E9",   // 浅蓝
  roleAuth: "EC7063",     // 红
  rolePurple: "AF7AC5",   // 浅紫
  noteBg: "FEF9E7",
  noteBorder: "F4D03F",
  errorBg: "FDF2F0",
  errorBorder: "EC7063",
  infoBg: "EBF8FB",
  infoBorder: "5DADE2",
};

const TF = "Microsoft YaHei"; // title font
const BF = "Microsoft YaHei"; // body font

// ============================================================
// HELPER: slide header
// ============================================================
function header(slide, num, title, subtitle) {
  slide.background = { color: C.bg };
  // top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.06, fill: { color: C.primary },
  });
  // page number badge
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 0.18, w: 0.42, h: 0.38,
    fill: { color: C.primary }, rectRadius: 0.05,
  });
  slide.addText(String(num).padStart(2, '0'), {
    x: M, y: 0.18, w: 0.42, h: 0.38,
    fontSize: 13, fontFace: TF, color: C.white, bold: true,
    align: "center", valign: "middle",
  });
  // title
  slide.addText(title, {
    x: M + 0.52, y: 0.15, w: CW - 1.0, h: 0.42,
    fontSize: 20, fontFace: TF, color: C.dark, bold: true,
    valign: "middle", fit: "shrink",
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: M + 0.52, y: 0.52, w: CW - 1.0, h: 0.24,
      fontSize: 10, fontFace: BF, color: C.textLight,
      valign: "middle", fit: "shrink",
    });
  }
  // divider line
  slide.addShape(pres.shapes.RECTANGLE, {
    x: M, y: 0.82, w: CW, h: 0.015, fill: { color: C.primary },
  });
}

// ============================================================
// HELPER: section label
// ============================================================
function sectionLabel(slide, label) {
  slide.addText(label, {
    x: M, y: 0.92, w: CW, h: 0.22,
    fontSize: 9, fontFace: BF, color: C.accent, bold: true,
  });
}

// ============================================================
// HELPER: notice box (注意事项)
// ============================================================
function noticeBox(slide, x, y, w, h, items) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: C.noteBg }, rectRadius: 0.06,
    line: { color: C.noteBorder, width: 1 },
  });
  slide.addText("⚠ 注意事项", {
    x: x + 0.12, y: y + 0.04, w: w - 0.24, h: 0.22,
    fontSize: 10, fontFace: TF, color: C.accent, bold: true,
  });
  const txtH = h - 0.3;
  const arr = items.map((t, i) => ({
    text: `(${i + 1}) ${t}`,
    options: { breakLine: true, paraSpaceAfter: 2 },
  }));
  slide.addText(arr, {
    x: x + 0.12, y: y + 0.26, w: w - 0.24, h: txtH,
    fontSize: 8, fontFace: BF, color: C.text,
    valign: "top", fit: "shrink",
  });
}

// ============================================================
// HELPER: error box (易错点)
// ============================================================
function errorBox(slide, x, y, w, h, items) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: C.errorBg }, rectRadius: 0.06,
    line: { color: C.errorBorder, width: 1 },
  });
  slide.addText("✕ 易错点提示", {
    x: x + 0.12, y: y + 0.04, w: w - 0.24, h: 0.22,
    fontSize: 10, fontFace: TF, color: C.danger, bold: true,
  });
  const txtH = h - 0.3;
  const arr = items.map((t, i) => ({
    text: `• ${t}`,
    options: { breakLine: true, paraSpaceAfter: 2 },
  }));
  slide.addText(arr, {
    x: x + 0.12, y: y + 0.26, w: w - 0.24, h: txtH,
    fontSize: 8, fontFace: BF, color: C.text,
    valign: "top", fit: "shrink",
  });
}

// ============================================================
// HELPER: info card
// ============================================================
function infoCard(slide, x, y, w, h, title, lines, opts = {}) {
  const bg = opts.bg || C.cardBg;
  const border = opts.border || C.secondary;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: bg }, rectRadius: 0.05,
    line: { color: border, width: 1 },
  });
  // left accent
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y: y + 0.04, w: 0.05, h: h - 0.08, fill: { color: border },
  });
  if (title) {
    slide.addText(title, {
      x: x + 0.14, y: y + 0.06, w: w - 0.24, h: 0.24,
      fontSize: 11, fontFace: TF, color: border, bold: true,
    });
  }
  const bodyY = title ? y + 0.3 : y + 0.08;
  const bodyH = title ? h - 0.36 : h - 0.14;
  const arr = lines.map(l => {
    if (typeof l === "string") return { text: l, options: { bullet: true, breakLine: true, paraSpaceAfter: 1 } };
    return { text: l.text, options: { ...l.options, breakLine: true, paraSpaceAfter: 1 } };
  });
  slide.addText(arr, {
    x: x + 0.14, y: bodyY, w: w - 0.24, h: bodyH,
    fontSize: opts.fontSize || 9, fontFace: BF, color: C.text,
    valign: "top", fit: "shrink",
  });
}

// ============================================================
// HELPER: step badge
// ============================================================
function stepBadge(slide, x, y, num, label) {
  slide.addShape(pres.shapes.OVAL, {
    x, y, w: 0.32, h: 0.32, fill: { color: C.primary },
  });
  slide.addText(String(num), {
    x, y, w: 0.32, h: 0.32,
    fontSize: 12, fontFace: TF, color: C.white, bold: true,
    align: "center", valign: "middle",
  });
  if (label) {
    slide.addText(label, {
      x: x + 0.38, y: y - 0.02, w: 5, h: 0.36,
      fontSize: 12, fontFace: TF, color: C.dark, bold: true,
      valign: "middle", fit: "shrink",
    });
  }
}

// ============================================================
// HELPER: footer
// ============================================================
function footer(slide) {
  slide.addText("个人客户信息维护 · 操作培训手册", {
    x: M, y: H - 0.3, w: CW, h: 0.22,
    fontSize: 7, fontFace: BF, color: C.textLight,
    align: "center",
  });
}

// ============================================================
// SLIDE 1: COVER
// ============================================================
{
  let s = pres.addSlide();
  s.background = { color: C.dark };
  // decorative left bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.08, h: H, fill: { color: C.accent },
  });
  // top thin line
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.4, w: 1.5, h: 0.02, fill: { color: C.accent },
  });
  s.addText("个人客户信息维护", {
    x: 0.8, y: 1.6, w: 8.4, h: 0.9,
    fontSize: 42, fontFace: TF, color: C.white, bold: true,
    align: "left", valign: "middle",
  });
  s.addText("操作培训手册", {
    x: 0.8, y: 2.55, w: 8.4, h: 0.5,
    fontSize: 22, fontFace: TF, color: "82E0AA",
    align: "left", valign: "middle",
  });
  // accent line
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 3.2, w: 2.5, h: 0.03, fill: { color: C.accent },
  });
  s.addText("柜面 · 智能柜员机 · 移动Pad  全渠道操作培训", {
    x: 0.8, y: 3.4, w: 8.4, h: 0.35,
    fontSize: 13, fontFace: BF, color: "D5F5E3",
    align: "left", valign: "middle",
  });
  s.addText("交易代码 030401  |  权限：业务柜员", {
    x: 0.8, y: 4.3, w: 8.4, h: 0.3,
    fontSize: 11, fontFace: BF, color: "82E0AA",
    align: "left", valign: "middle",
  });
  s.addNotes("【开场白】各位同事大家好，今天我们来学习个人客户信息维护的操作流程。本培训面向柜面及厅堂业务人员，目标是让大家掌握030401交易的完整操作，涵盖柜面、智能柜员机、移动Pad三个渠道。培训时长约40分钟，请大家关注每一步的关键字段和注意事项。");
}

// ============================================================
// SLIDE 2: TABLE OF CONTENTS
// ============================================================
{
  let s = pres.addSlide();
  header(s, 2, "培训内容概览", "目录");

  const items = [
    { num: "01", title: "场景说明", desc: "场景定义、支持渠道与业务范围" },
    { num: "02", title: "业务办理总览", desc: "完整操作流程，7个关键环节" },
    { num: "03", title: "职能流程图", desc: "跨角色泳道图，清晰展现协作关系" },
    { num: "04", title: "柜面操作详解", desc: "8步操作流程，含代理办理规则" },
    { num: "05", title: "智能柜员机操作", desc: "6步自助操作流程" },
    { num: "06", title: "移动Pad操作", desc: "7步移动办理流程" },
    { num: "07", title: "业务规则与字段", desc: "核心规则、证件类型、凭证管理" },
  ];

  const cardW = 1.18, cardH = 3.4, gap = 0.08;
  let startX = (W - (cardW * 7 + gap * 6)) / 2;
  items.forEach((it, i) => {
    let x = startX + i * (cardW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.15, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.05,
      line: { color: C.secondary, width: 1 },
    });
    // top accent
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.15, w: cardW, h: 0.05, fill: { color: C.primary },
    });
    s.addText(it.num, {
      x: x + 0.05, y: 1.35, w: cardW - 0.1, h: 0.45,
      fontSize: 24, fontFace: TF, color: C.primary, bold: true,
      align: "center",
    });
    s.addText(it.title, {
      x: x + 0.05, y: 1.85, w: cardW - 0.1, h: 0.55,
      fontSize: 10, fontFace: TF, color: C.dark, bold: true,
      align: "center", valign: "top", fit: "shrink",
    });
    s.addText(it.desc, {
      x: x + 0.05, y: 2.5, w: cardW - 0.1, h: 1.9,
      fontSize: 7.5, fontFace: BF, color: C.textLight,
      align: "center", valign: "top", fit: "shrink",
    });
  });
  s.addNotes("【章节过渡】本次培训分为七大模块：场景说明、业务办理总览、职能流程图、柜面操作详解、智能柜员机操作、移动Pad操作、业务规则与字段。我们先从场景说明开始，了解个人客户信息维护的业务背景。");
  footer(s);
}

// ============================================================
// SLIDE 3: SCENE DESCRIPTION
// ============================================================
{
  let s = pres.addSlide();
  header(s, 3, "场景说明", "个人客户信息维护场景定义与业务范围");

  // Left card: scene definition
  infoCard(s, M, 1.05, 4.3, 2.2, "场景定义", [
    "为客户提供客户信息维护功能",
    "各场景可通过客户信息检查跳转至本场景",
    "支持个人客户信息的新建及存量正常客户的信息修改",
    "支持通过柜面、智能柜员机、移动Pad进行业务办理",
    "柜面支持代理维护信息",
  ], { border: C.primary, fontSize: 9 });

  // Right card: supported channels
  infoCard(s, M + 4.5, 1.05, 4.5, 2.2, "支持渠道", [
    { text: "柜面渠道", options: { bold: true, color: C.primary } },
    "客户持有效身份证件办理，支持代理",
    { text: "智能柜员机", options: { bold: true, color: C.secondary } },
    "客户自助办理，需厅堂人员授权",
    { text: "移动Pad", options: { bold: true, color: C.success } },
    "厅堂服务人员手持Pad上门或移动办理",
  ], { border: C.secondary, fontSize: 9 });

  // Bottom: key info bar
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 3.45, w: CW, h: 1.4,
    fill: { color: C.dark }, rectRadius: 0.06,
  });
  s.addText("核心要点", {
    x: M + 0.2, y: 3.55, w: 2, h: 0.28,
    fontSize: 11, fontFace: TF, color: C.accent, bold: true,
  });
  const pts = [
    "信息维护 = 新建客户信息 + 存量客户信息修改",
    "三要素验证：证件号码 + 姓名 + 证件类型",
    "身份核实 = 联网核查 + 人脸识别 + 证件影像留存",
    "九要素信息：姓名、性别、国籍、职业、地址、联系方式、证件类型、证件号码、证件有效期",
  ];
  s.addText(pts.map(p => ({ text: p, options: { bullet: true, breakLine: true, paraSpaceAfter: 3 } })), {
    x: M + 0.2, y: 3.85, w: CW - 0.4, h: 0.9,
    fontSize: 9, fontFace: BF, color: C.white,
    valign: "top", fit: "shrink",
  });
  s.addNotes("【背景说明】个人客户信息维护是一项基础业务，很多其他业务场景（如开户、签约）都可以通过客户信息检查跳转到本场景。因此掌握好这个交易非常重要。\n【关键概念】三要素验证是新建客户时的唯一性校验——证件号码、姓名、证件类型三者组合必须唯一。九要素是反洗钱要求的客户身份基本信息最低标准。");
  footer(s);
}

// ============================================================
// SLIDE 4: BUSINESS PROCESS OVERVIEW
// ============================================================
{
  let s = pres.addSlide();
  header(s, 4, "业务办理流程概述", "操作介绍 · 7个关键环节");

  const steps = [
    { n: 1, t: "客户到达与身份识别", d: "客户到达网点后，由取号机、厅堂服务人员手持Pad对客户身份进行识别，结合网点忙闲和客户情况进行分流或营销。" },
    { n: 2, t: "身份核实与确认", d: "通过联网核查、人脸识别等方式核实客户身份。人脸识别不通过时，由上一级管理人员进行现场审核。" },
    { n: 3, t: "信息比对与判断", d: "系统通过有效证件提取的姓名、证件类型、证件号码与存量客户比对。比对不一致引导新建，一致引导修改。" },
    { n: 4, t: "联系方式验证", d: "若联系方式或主手机号修改，需先核查手机号码，返回结果后进行短信验证码验证。结果仅作提示，不拦截办理。" },
    { n: 5, t: "电子签名确认", d: "系统回显客户声明、风险提示及电子凭证。客户确认后进行电子签名（正楷或指纹），签名需柜员审核。" },
    { n: 6, t: "回单打印与评价", d: "提供纸质回单、电子回单打印服务，客户对本次服务进行评价。" },
    { n: 7, t: "凭证质检与归档", d: "业务办理结束后，系统对凭证进行质检并自动归档处理。" },
  ];

  const cardW = 1.2, cardH = 3.5, gap = 0.07;
  let startX = (W - (cardW * 7 + gap * 6)) / 2;
  steps.forEach((st, i) => {
    let x = startX + i * (cardW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.1, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.05,
      line: { color: C.secondary, width: 1 },
    });
    s.addShape(pres.shapes.OVAL, {
      x: x + cardW / 2 - 0.22, y: 1.22, w: 0.44, h: 0.44,
      fill: { color: C.primary },
    });
    s.addText(String(st.n), {
      x: x + cardW / 2 - 0.22, y: 1.22, w: 0.44, h: 0.44,
      fontSize: 16, fontFace: TF, color: C.white, bold: true,
      align: "center", valign: "middle",
    });
    s.addText(st.t, {
      x: x + 0.06, y: 1.75, w: cardW - 0.12, h: 0.55,
      fontSize: 9, fontFace: TF, color: C.dark, bold: true,
      align: "center", valign: "top", fit: "shrink",
    });
    s.addText(st.d, {
      x: x + 0.06, y: 2.35, w: cardW - 0.12, h: 2.1,
      fontSize: 7.5, fontFace: BF, color: C.text,
      valign: "top", fit: "shrink",
    });
  });
  // arrows
  for (let i = 0; i < 6; i++) {
    let ax = startX + (i + 1) * cardW + i * gap + gap / 2 - 0.09;
    s.addText("→", {
      x: ax, y: 2.7, w: 0.25, h: 0.25,
      fontSize: 12, fontFace: TF, color: C.accent, bold: true,
      align: "center", valign: "middle",
    });
  }
  s.addNotes("【讲解节奏】本页是业务全流程总览，7个环节串起了从客户进门到业务结束的完整路径。建议用时3分钟快速过一遍，不展开细节。\n【关键环节】第3步信息比对是分支点——决定走新建还是修改路径。第4步联系方式验证是容易忽略的环节，手机号变更必须先核查再验证码。\n【互动提问】大家想一想，人脸识别不通过时应该怎么处理？");
  footer(s);
}

// ============================================================
// SLIDE 5: SWIMLANE DIAGRAM
// ============================================================
{
  let s = pres.addSlide();
  header(s, 5, "职能流程图", "按角色分色泳道图 · 客户/厅堂人员/柜员/系统/授权人员");

  const lanes = [
    { name: "客户", color: C.roleCustomer },
    { name: "厅堂人员", color: C.roleHall },
    { name: "柜员", color: C.roleTeller },
    { name: "系统", color: C.roleSystem },
    { name: "授权人员", color: C.roleAuth },
  ];

  const laneStartY = 1.05;
  const laneH = 0.62;
  const labelW = 0.95;
  const procX = M + labelW;
  const procW = W - M - procX;

  // Draw lanes
  lanes.forEach((ln, i) => {
    let y = laneStartY + i * laneH;
    // process area background
    s.addShape(pres.shapes.RECTANGLE, {
      x: procX, y, w: W - M - procX, h: laneH,
      fill: { color: i % 2 === 0 ? C.white : C.bg },
    });
    // lane label
    s.addShape(pres.shapes.RECTANGLE, {
      x: M, y, w: labelW, h: laneH,
      fill: { color: ln.color },
    });
    s.addText(ln.name, {
      x: M, y, w: labelW, h: laneH,
      fontSize: 10, fontFace: TF, color: C.white, bold: true,
      align: "center", valign: "middle",
    });
    // separator line
    if (i < lanes.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: M, y: y + laneH, w: W - 2 * M, h: 0,
        line: { color: C.textLight, width: 0.5 },
      });
    }
  });

  // Process nodes
  const nodes = [
    { label: "客户到达\n网点", lane: 0, x: 0.1 },
    { label: "身份识别\n与分流", lane: 1, x: 1.25 },
    { label: "联网核查\n人脸识别", lane: 2, x: 2.4 },
    { label: "三要素比对\n信息维护", lane: 3, x: 3.55 },
    { label: "手机号核查\n验证码验证", lane: 3, x: 4.7, dashed: false },
    { label: "智能授权\n(税收居民\n变化)", lane: 4, x: 5.85, dashed: true },
    { label: "签名确认\n打印评价\n归档", lane: 2, x: 7.0 },
  ];

  const nodeW = 1.05, nodeH = 0.45;
  nodes.forEach(nd => {
    let nx = procX + nd.x;
    let ny = laneStartY + nd.lane * laneH + (laneH - nodeH) / 2;
    let shapeType = nd.dashed ? pres.shapes.OVAL : pres.shapes.ROUNDED_RECTANGLE;
    s.addShape(shapeType, {
      x: nx, y: ny, w: nodeW, h: nodeH,
      fill: { color: lanes[nd.lane].color },
      line: nd.dashed ? { color: C.danger, width: 1.5, dashType: "dash" } : { color: lanes[nd.lane].color, width: 1 },
      rectRadius: 0.04,
    });
    s.addText(nd.label, {
      x: nx, y: ny, w: nodeW, h: nodeH,
      fontSize: 7, fontFace: BF, color: C.white, bold: true,
      align: "center", valign: "middle", fit: "shrink",
    });
  });

  // Arrows between nodes
  const arrowData = [
    { from: { lane: 0, x: 0.1 }, to: { lane: 1, x: 1.25 } },
    { from: { lane: 1, x: 1.25 }, to: { lane: 2, x: 2.4 } },
    { from: { lane: 2, x: 2.4 }, to: { lane: 3, x: 3.55 } },
    { from: { lane: 3, x: 3.55 }, to: { lane: 3, x: 4.7 } },
    { from: { lane: 3, x: 4.7 }, to: { lane: 4, x: 5.85 } },
    { from: { lane: 4, x: 5.85 }, to: { lane: 2, x: 7.0 } },
  ];
  arrowData.forEach(a => {
    let x1 = procX + a.from.x + nodeW;
    let y1 = laneStartY + a.from.lane * laneH + laneH / 2;
    let x2 = procX + a.to.x;
    let y2 = laneStartY + a.to.lane * laneH + laneH / 2;
    let midX = (x1 + x2) / 2;
    s.addShape(pres.shapes.LINE, {
      x: x1, y: y1, w: midX - x1, h: 0,
      line: { color: C.accent, width: 1.5, endArrowType: "none" },
    });
    s.addShape(pres.shapes.LINE, {
      x: midX, y: Math.min(y1, y2), w: 0, h: Math.abs(y2 - y1),
      line: { color: C.accent, width: 1.5, endArrowType: "none" },
    });
    s.addShape(pres.shapes.LINE, {
      x: midX, y: y2, w: x2 - midX, h: 0,
      line: { color: C.accent, width: 1.5, endArrowType: "triangle" },
    });
  });

  // Legend
  const legY = 4.4;
  s.addText("图例：", {
    x: M, y: legY, w: 0.5, h: 0.25,
    fontSize: 9, fontFace: TF, color: C.dark, bold: true,
    valign: "middle",
  });
  const legends = [
    { color: C.roleCustomer, label: "客户" },
    { color: C.roleHall, label: "厅堂人员" },
    { color: C.roleTeller, label: "柜员" },
    { color: C.roleSystem, label: "系统" },
    { color: C.roleAuth, label: "授权人员" },
  ];
  legends.forEach((lg, i) => {
    let lx = M + 0.55 + i * 1.0;
    s.addShape(pres.shapes.RECTANGLE, {
      x: lx, y: legY + 0.04, w: 0.18, h: 0.18,
      fill: { color: lg.color },
    });
    s.addText(lg.label, {
      x: lx + 0.22, y: legY, w: 0.7, h: 0.25,
      fontSize: 9, fontFace: BF, color: C.text, valign: "middle",
    });
  });
  // dashed node legend
  s.addShape(pres.shapes.OVAL, {
    x: M + 5.7, y: legY + 0.04, w: 0.18, h: 0.18,
    fill: { color: C.roleAuth },
    line: { color: C.danger, width: 1.5, dashType: "dash" },
  });
  s.addText("条件触发节点", {
    x: M + 5.95, y: legY, w: 1.5, h: 0.25,
    fontSize: 9, fontFace: BF, color: C.text, valign: "middle",
  });
  s.addNotes("【角色交接要点】重点关注跨泳道交接：①客户→厅堂人员：身份识别分流；②厅堂人员→柜员：引导至柜台；③系统→授权人员：税收居民标识变化触发授权。\n【条件节点】虚线椭圆表示条件触发——税收居民标识发生变化时才需要授权，不是每笔都走。授权不通过，交易终止；授权通过，柜员继续后续服务。");
  footer(s);
}

// ============================================================
// SLIDE 6: COUNTER OPERATION OVERVIEW
// ============================================================
{
  let s = pres.addSlide();
  header(s, 6, "柜面操作总览", "交易代码 · 权限 · 8步操作流程");

  // Transaction code & permission bar
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 1.0, w: CW, h: 0.55,
    fill: { color: C.dark }, rectRadius: 0.05,
  });
  s.addText([
    { text: "交易代码：", options: { fontSize: 12, color: "82E0AA" } },
    { text: "030401 个人客户信息维护", options: { fontSize: 12, color: C.white, bold: true } },
    { text: "    |    ", options: { fontSize: 12, color: "7D8C8D" } },
    { text: "交易权限：", options: { fontSize: 12, color: "82E0AA" } },
    { text: "业务柜员", options: { fontSize: 12, color: C.white, bold: true } },
  ], {
    x: M + 0.2, y: 1.0, w: CW - 0.4, h: 0.55,
    fontFace: TF, align: "center", valign: "middle",
  });

  // 8 step cards
  const stepData = [
    { n: 1, t: "进入交易", d: "输入交易代码\n030401或场景名称" },
    { n: 2, t: "证件识别", d: "选择证件类型\n读取+人脸识别" },
    { n: 3, t: "代理选择", d: "判断是否代理\n代理需核实身份" },
    { n: 4, t: "信息查看", d: "查看要素信息\n确认无误下一步" },
    { n: 5, t: "详细填写", d: "详细信息/税收信息\n现居地址填写" },
    { n: 6, t: "交易确认", d: "柜外清签名/指纹\n柜员审核确认" },
    { n: 7, t: "凭证打印", d: "纸质/电子回单\n打印" },
    { n: 8, t: "客户评价", d: "非常满意/满意\n不满意/未评价" },
  ];

  const cardW = 1.03, cardH = 2.6, gap = 0.08;
  let startX = (W - (cardW * 8 + gap * 7)) / 2;
  stepData.forEach((st, i) => {
    let x = startX + i * (cardW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.75, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.05,
      line: { color: C.secondary, width: 1 },
    });
    // top accent
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.75, w: cardW, h: 0.05, fill: { color: C.primary },
    });
    // number
    s.addShape(pres.shapes.OVAL, {
      x: x + cardW / 2 - 0.18, y: 1.88, w: 0.36, h: 0.36,
      fill: { color: C.primary },
    });
    s.addText(String(st.n), {
      x: x + cardW / 2 - 0.18, y: 1.88, w: 0.36, h: 0.36,
      fontSize: 13, fontFace: TF, color: C.white, bold: true,
      align: "center", valign: "middle",
    });
    s.addText(st.t, {
      x: x + 0.04, y: 2.3, w: cardW - 0.08, h: 0.3,
      fontSize: 9, fontFace: TF, color: C.dark, bold: true,
      align: "center", valign: "middle", fit: "shrink",
    });
    s.addText(st.d, {
      x: x + 0.04, y: 2.65, w: cardW - 0.08, h: 1.5,
      fontSize: 7.5, fontFace: BF, color: C.text,
      align: "center", valign: "top", fit: "shrink",
    });
  });

  // arrow
  s.addText("→ → → → → → → →", {
    x: startX, y: 4.45, w: cardW * 8 + gap * 7, h: 0.22,
    fontSize: 10, fontFace: TF, color: C.accent, bold: true,
    align: "center", valign: "middle",
  });
  s.addNotes("【讲解节奏】本页是柜面8步操作的总览，建议用时2分钟快速过一遍。后续每一步或每两步会有独立页面详细讲解。\n【关键记忆】交易代码030401是必须记住的。与开户不同，信息维护多了一步「代理选择」——因为柜面支持代理办理。\n【互动提问】大家猜猜，代理办理有什么限制条件？");
  footer(s);
}

// ============================================================
// SLIDE 7: COUNTER STEPS 1-3
// ============================================================
{
  let s = pres.addSlide();
  header(s, 7, "柜面步骤 1-3：进入交易 · 证件识别 · 代理选择", "操作说明 · 关键字段 · 注意事项");

  // Step 1
  stepBadge(s, M, 1.0, 1, "进入交易界面");
  infoCard(s, M, 1.38, CW, 0.55, null, [
    "柜员成功登录后，输入交易代码\"030401\"或场景名称进入\"个人客户信息维护\"场景界面",
  ], { border: C.primary, fontSize: 10 });

  // Step 2
  stepBadge(s, M, 2.05, 2, "证件识别与人脸识别");
  infoCard(s, M, 2.43, 5.3, 1.2, "操作说明", [
    "根据客户提供有效身份证件类型选择对应业务办理的证件类型",
    "将有效身份证件放在证件读取区域（或手工输入）",
    "客户身份核实无误后点击\"下一步\"进入代理办理选择界面",
  ], { border: C.secondary, fontSize: 9 });

  // Step 2 notice
  noticeBox(s, M + 5.5, 2.43, 3.5, 1.2, [
    "证件类型为户口簿发起新建或修改时，证件有效到期日系统自动计算并填充，置灰且不允许修改",
  ]);

  // Step 3
  stepBadge(s, M, 3.75, 3, "代理办理选择");
  infoCard(s, M, 4.13, 5.3, 1.0, "操作说明", [
    "若为代理选择\"是\"，进入代理人客户身份核实界面，核实后进行人脸识别，通过后点击\"下一步\"",
    "若本人办理选择\"否\"，进入人脸识别界面，通过后点击\"下一步\"",
  ], { border: C.secondary, fontSize: 9 });

  // Step 3 notice
  noticeBox(s, M + 5.5, 4.13, 3.5, 1.0, [
    "支持16周岁以下（不含）代理，需上传相关证明资料",
    "不支持16周岁以上（含）代理",
    "支持16周岁以上（含）特殊人群代理，需上传相关证明资料",
  ]);

  s.addNotes("【讲解节奏】步骤1简单带过（30秒），步骤2和步骤3是重点，建议各2分钟。\n【易混淆点】户口簿的证件有效期是系统自动计算的，不需要手动输入，而且是置灰的不能改。\n【代理规则】代理办理有严格的年龄限制：16岁以下可以代理，16岁以上一般不支持，只有特殊人群（如重病、行动不便）才能代理且需上传证明资料。");
  footer(s);
}

// ============================================================
// SLIDE 8: COUNTER STEPS 4-6
// ============================================================
{
  let s = pres.addSlide();
  header(s, 8, "柜面步骤 4-6：信息维护 · 详细填写 · 交易确认", "操作说明 · 关键字段 · 注意事项");

  // Step 4
  stepBadge(s, M, 1.0, 4, "客户信息维护界面");
  infoCard(s, M, 1.38, CW, 0.55, null, [
    "进入客户信息维护界面，查看要素信息。要素信息无误后，点击\"下一步\"。",
  ], { border: C.primary, fontSize: 10 });

  // Step 5
  stepBadge(s, M, 2.05, 5, "详细信息填写");
  infoCard(s, M, 2.43, 5.3, 1.3, "填写内容", [
    { text: "详细信息", options: { bold: true, color: C.primary } },
    { text: "税收信息", options: { bold: true, color: C.primary } },
    { text: "现居地址", options: { bold: true, color: C.primary } },
    "证件地址支持一键同步为现居地址",
    "输入正确验证码完成联系方式验证后，可点击同步至主手机号",
  ], { border: C.secondary, fontSize: 9 });

  // Step 5 key point
  noticeBox(s, M + 5.5, 2.43, 3.5, 1.3, [
    "自动提取信息（不可修改）：姓名、国籍、性别、证件类型、证件号码、证件有效起始日、证件有效到期日、签发机关、民族、出生日期",
    "需手动补充：职业、联系方式、现居地址、税收居民标识等",
  ]);

  // Step 6
  stepBadge(s, M, 3.85, 6, "交易确认与签名");
  infoCard(s, M, 4.23, CW, 0.85, "操作说明", [
    "进入交易确认界面，查看输入项无误后点击提交",
    "提醒客户由柜外清进行交易确认并完成正楷签名（需柜员审核确认）或指纹录入",
    "确认后提交后台进行业务处理",
  ], { border: C.primary, fontSize: 9 });

  s.addNotes("【讲解节奏】步骤4-6是信息维护的核心环节，建议用时4分钟。\n【关键概念】步骤5中有两类字段要区分清楚：一类是证件自动提取的，不可修改（共10项）；另一类是需要手动补充的。\n【易混淆点】证件地址和现居地址是两个不同的字段，证件地址从证件读取，现居地址需要客户提供。但有个便捷功能——证件地址可以一键同步为现居地址。\n【签名要求】签名必须正楷，不符合要求的或者指纹不清晰的，需要求客户重新签字确认。");
  footer(s);
}

// ============================================================
// SLIDE 9: COUNTER STEPS 7-8
// ============================================================
{
  let s = pres.addSlide();
  header(s, 9, "柜面步骤 7-8：凭证打印与客户评价", "操作说明 · 凭证类型 · 评价选项");

  // Step 7
  stepBadge(s, M, 1.0, 7, "凭证及回单打印");
  infoCard(s, M, 1.38, 4.3, 1.6, "操作说明", [
    "查看交易结果并完成相关凭证打印",
    "若客户需要纸质回单，则进行纸质回单打印",
    "若客户不需要纸质回单，可通过手机银行扫描二维码查看业务办理信息",
  ], { border: C.primary, fontSize: 9 });

  infoCard(s, M + 4.5, 1.38, 4.5, 1.6, "回单类型", [
    { text: "纸质回单", options: { bold: true, color: C.primary } },
    "打印物理凭证交客户保管",
    { text: "电子回单", options: { bold: true, color: C.secondary } },
    "手机银行扫描二维码查看业务办理信息",
    "环保便捷，随时查阅",
  ], { border: C.secondary, fontSize: 9 });

  // Step 8
  stepBadge(s, M, 3.15, 8, "客户评价");
  infoCard(s, M, 3.53, CW, 0.9, "操作说明与评价选项", [
    "由客户在柜外清对本次交易进行评价",
    "可选评价：非常满意 / 满意 / 不满意",
    "不评价默认为\"客户未评价\"",
  ], { border: C.primary, fontSize: 10 });

  // Bottom note
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.6, w: CW, h: 0.4,
    fill: { color: C.infoBg }, rectRadius: 0.04,
    line: { color: C.infoBorder, width: 1 },
  });
  s.addText("ℹ 业务办理结束后，系统对凭证进行质检并进行自动归档处理", {
    x: M + 0.15, y: 4.6, w: CW - 0.3, h: 0.4,
    fontSize: 9, fontFace: TF, color: C.secondary, bold: true,
    valign: "middle",
  });
  s.addNotes("【讲解节奏】步骤7-8合并讲解，建议用时2分钟。\n【服务提示】推荐客户使用电子回单，既环保又方便随时查阅。但如果客户习惯纸质回单，还是要按客户需求打印。\n【注意事项】系统自动质检归档是后台自动完成的，柜员不需要手动操作。但如果质检不通过，系统会有提示，需要按提示处理。");
  footer(s);
}

// ============================================================
// SLIDE 10: SMART TELLER MACHINE OPERATION
// ============================================================
{
  let s = pres.addSlide();
  header(s, 10, "智能柜员机操作流程", "6步自助操作流程 · 客户自助 + 厅堂授权");

  const steps = [
    { n: 1, t: "选择服务类型", d: "客户自助选择服务类型，点击\"自助办理\"后进入下一界面" },
    { n: 2, t: "进入个人信息维护", d: "点击\"信息管理\"，进入界面后点击\"个人信息维护\"进入交易场景" },
    { n: 3, t: "身份证读取+人脸识别", d: "将二代身份证插入指定区域，成功读取后进行人脸识别，完成后进入维护界面" },
    { n: 4, t: "交易确认与签名", d: "阅读《个人税收居民文件》并点击\"同意\"，查看无误后点击提交，客户自主正楷签名或指纹签名后点击\"完成\"" },
    { n: 5, t: "厅堂人员授权", d: "厅堂服务人员查看交易信息及客户签名情况，确认无误后输入柜员号，选择密码或指纹登录后点击\"授权\"" },
    { n: 6, t: "交易结果与打印", d: "查看交易结果并完成相应凭证及回单的打印。需要纸质凭条点击\"打印凭条\"，也可扫码查看电子回单" },
  ];

  const cardW = 1.35, cardH = 3.5, gap = 0.1;
  let startX = (W - (cardW * 6 + gap * 5)) / 2;
  steps.forEach((st, i) => {
    let x = startX + i * (cardW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.1, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.05,
      line: { color: C.success, width: 1 },
    });
    s.addShape(pres.shapes.OVAL, {
      x: x + cardW / 2 - 0.22, y: 1.22, w: 0.44, h: 0.44,
      fill: { color: C.success },
    });
    s.addText(String(st.n), {
      x: x + cardW / 2 - 0.22, y: 1.22, w: 0.44, h: 0.44,
      fontSize: 16, fontFace: TF, color: C.white, bold: true,
      align: "center", valign: "middle",
    });
    s.addText(st.t, {
      x: x + 0.08, y: 1.75, w: cardW - 0.16, h: 0.5,
      fontSize: 9.5, fontFace: TF, color: C.dark, bold: true,
      align: "center", valign: "top", fit: "shrink",
    });
    s.addText(st.d, {
      x: x + 0.08, y: 2.3, w: cardW - 0.16, h: 2.1,
      fontSize: 8, fontFace: BF, color: C.text,
      valign: "top", fit: "shrink",
    });
  });

  // Key feature bar
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.75, w: CW, h: 0.45,
    fill: { color: C.dark }, rectRadius: 0.04,
  });
  s.addText("特点：客户自助办理为主 · 厅堂人员仅需最终授权 · 支持二代身份证等多种证件", {
    x: M + 0.15, y: 4.75, w: CW - 0.3, h: 0.45,
    fontSize: 9, fontFace: TF, color: C.white, bold: false,
    align: "center", valign: "middle",
  });
  s.addNotes("【讲解节奏】智能柜员机6步流程，建议用时3分钟。\n【与柜面区别】智能柜员机是客户自助操作，厅堂人员只在最后一步授权。柜面是柜员全程操作。\n【注意点】智能柜员机需要客户阅读《个人税收居民文件》并点击同意，这一步是客户自助完成的，柜员不直接参与，但授权时要确认客户已同意。\n【易错点】厅堂人员授权需要先输入柜员号，再选择密码或指纹登录，然后才能点击授权。顺序不能错。");
  footer(s);
}

// ============================================================
// SLIDE 11: MOBILE PAD OPERATION
// ============================================================
{
  let s = pres.addSlide();
  header(s, 11, "移动Pad操作流程", "7步移动办理流程 · 厅堂人员手持Pad上门或移动办理");

  const steps = [
    { n: 1, t: "登录并选择场景", d: "成功登录后，选择\"个人信息维护\"，进入交易界面" },
    { n: 2, t: "证件读取", d: "选择业务办理的证件类型（如二代身份证），将客户身份证放置在Pad背夹证件读取区域，进行证件信息读取" },
    { n: 3, t: "人脸识别", d: "证件读取成功后，进行人脸识别。完成人脸识别后，点击\"下一步\"" },
    { n: 4, t: "信息填写", d: "进入客户信息维护界面，根据客户情况完成相关字段信息的填写，完成后点击\"下一步\"" },
    { n: 5, t: "交易确认", d: "进入交易确认界面，查看交易信息，确认无误后点击\"提交\"" },
    { n: 6, t: "签名确认", d: "根据提示，客户自主进行正楷签名或指纹签名后，点击\"下一步\"。若有误可点击\"重新输入\"" },
    { n: 7, t: "结果与打印", d: "查看交易结果，通过蓝牙打印机完成相应凭证及回单的打印。也可扫码查看电子回单" },
  ];

  const cardW = 1.15, cardH = 3.5, gap = 0.09;
  let startX = (W - (cardW * 7 + gap * 6)) / 2;
  steps.forEach((st, i) => {
    let x = startX + i * (cardW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.1, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.05,
      line: { color: C.warning, width: 1 },
    });
    s.addShape(pres.shapes.OVAL, {
      x: x + cardW / 2 - 0.2, y: 1.22, w: 0.4, h: 0.4,
      fill: { color: C.warning },
    });
    s.addText(String(st.n), {
      x: x + cardW / 2 - 0.2, y: 1.22, w: 0.4, h: 0.4,
      fontSize: 14, fontFace: TF, color: C.white, bold: true,
      align: "center", valign: "middle",
    });
    s.addText(st.t, {
      x: x + 0.06, y: 1.7, w: cardW - 0.12, h: 0.5,
      fontSize: 9, fontFace: TF, color: C.dark, bold: true,
      align: "center", valign: "top", fit: "shrink",
    });
    s.addText(st.d, {
      x: x + 0.06, y: 2.25, w: cardW - 0.12, h: 2.2,
      fontSize: 7.5, fontFace: BF, color: C.text,
      valign: "top", fit: "shrink",
    });
  });

  // Key feature bar
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.75, w: CW, h: 0.45,
    fill: { color: C.dark }, rectRadius: 0.04,
  });
  s.addText("特点：移动灵活 · 支持上门服务 · 蓝牙打印 · 厅堂服务人员全程操作", {
    x: M + 0.15, y: 4.75, w: CW - 0.3, h: 0.45,
    fontSize: 9, fontFace: TF, color: C.white, bold: false,
    align: "center", valign: "middle",
  });
  s.addNotes("【讲解节奏】移动Pad 7步流程，建议用时3分钟。\n【与其他渠道区别】移动Pad最大的特点是灵活，可以上门服务，适合行动不便的客户。使用蓝牙打印机，现场就能出凭证。\n【操作主体】移动Pad是厅堂服务人员全程操作，客户只需要签字确认。这和智能柜员机的客户自助不一样。\n【易错点】Pad的证件读取是通过背夹读取的，要注意放置位置正确。蓝牙打印机要提前配对好，避免现场出问题。");
  footer(s);
}

// ============================================================
// SLIDE 12: THREE CHANNELS COMPARISON TABLE
// ============================================================
{
  let s = pres.addSlide();
  header(s, 12, "三渠道对比", "柜面 · 智能柜员机 · 移动Pad 操作对比");

  const rows = [
    [
      { text: "对比维度", options: { bold: true, color: C.white, fill: { color: C.primary } } },
      { text: "柜面", options: { bold: true, color: C.white, fill: { color: C.primary } } },
      { text: "智能柜员机", options: { bold: true, color: C.white, fill: { color: C.primary } } },
      { text: "移动Pad", options: { bold: true, color: C.white, fill: { color: C.primary } } },
    ],
    ["操作主体", "柜员操作", "客户自助", "厅堂服务人员操作"],
    ["支持证件类型", "111-居民身份证、112-临时身份证、121-户口簿等多种", "二代身份证、港澳居民来往内地通行证、台湾居民来往大陆通行证、外国人永久居留证", "二代身份证、港澳居民来往内地通行证、台湾居民来往大陆通行证、外国人永久居留证"],
    ["代理办理", "支持（有条件）", "不支持", "不支持"],
    ["授权方式", "集中授权/智能授权", "厅堂人员现场授权", "智能授权"],
    ["凭证打印", "柜面打印机", "自助打印机", "蓝牙打印机"],
    ["适用场景", "全功能、复杂业务、代理业务", "标准化业务、客户自助", "上门服务、移动办理"],
    ["签名方式", "柜外清正楷/指纹", "设备正楷/指纹", "Pad正楷/指纹"],
  ];

  const colW = [1.5, 2.5, 2.5, 2.5];
  s.addTable(rows, {
    x: M, y: 1.05, w: CW,
    colW,
    border: { type: "solid", pt: 1, color: "BDC3C7" },
    rowH: 0.45,
    fontSize: 8, fontFace: BF, color: C.text,
    align: "center", valign: "middle",
    autoPage: false,
  });

  // Bottom note
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.7, w: CW, h: 0.5,
    fill: { color: C.noteBg }, rectRadius: 0.04,
    line: { color: C.noteBorder, width: 1 },
  });
  s.addText("💡 选择建议：简单信息修改优先引导智能柜员机；复杂业务或代理办理走柜面；行动不便客户可安排移动Pad上门", {
    x: M + 0.15, y: 4.7, w: CW - 0.3, h: 0.5,
    fontSize: 9, fontFace: TF, color: C.accent, bold: true,
    align: "center", valign: "middle",
  });
  s.addNotes("【讲解要点】三渠道对比表帮助大家快速理解各渠道的差异和适用场景。建议用时3分钟，重点讲证件类型和适用场景的区别。\n【易混淆点】柜面支持的证件类型最多，包括户口簿等；智能柜员机和移动Pad只支持4种带芯片的证件。\n【分流建议】简单的信息修改（如改手机号、地址）优先引导智能柜员机，减轻柜面压力。代理业务、户口簿等特殊证件必须走柜面。");
  footer(s);
}

// ============================================================
// SLIDE 13: BUSINESS RULES (1)
// ============================================================
{
  let s = pres.addSlide();
  header(s, 13, "业务规则（一）", "渠道证件 · 名单管理 · 联网核查 · 授权规则");

  const rules = [
    {
      title: "渠道与证件支持", color: C.primary, icon: "📋",
      items: [
        "支持柜面、智能柜员机、移动PAD三种渠道办理",
        "柜面支持多种有效身份证件（居民身份证、户口簿、港澳通行证、台湾通行证、中国护照、外国人永久居留证、外国人护照等）",
        "移动Pad、智能柜员机支持：二代身份证、港澳居民来往内地通行证、台湾居民来往大陆通行证、外国人永久居留证",
      ],
    },
    {
      title: "名单管理", color: C.danger, icon: "🚫",
      items: [
        "在名单的个人客户，不允许办理个人信息维护业务",
        "系统自动校验，命中名单直接拒绝交易",
      ],
    },
    {
      title: "联网核查规则", color: C.warning, icon: "🔍",
      items: [
        "姓名、身份证号码和照片存在一项或多项不一致的，作为疑义身份信息进一步核实",
        "\"公民身份号码存在，但与姓名不匹配\"、\"公民身份号码不存在\"不通过",
        "客户身份证联网核查无照片，则拒绝交易",
        "\"公民身份号码与姓名一致，且存在照片\"，核查通过，继续交易",
      ],
    },
    {
      title: "授权规则", color: C.secondary, icon: "🔐",
      items: [
        "智能授权（柜面、移动PAD）：税收居民标识发生变化需要授权",
        "集中授权审核未通过，交易终止",
        "集中授权审核通过，交易返回，柜员继续完成后续服务",
      ],
    },
  ];

  const cardW = 4.35, cardH = 1.7, gapX = 0.3, gapY = 0.15;
  rules.forEach((r, i) => {
    let col = i % 2, row = Math.floor(i / 2);
    let x = M + col * (cardW + gapX);
    let y = 1.05 + row * (cardH + gapY);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.05,
      line: { color: r.color, width: 1.5 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: cardH, fill: { color: r.color },
    });
    s.addText(`${r.icon} ${r.title}`, {
      x: x + 0.14, y: y + 0.06, w: cardW - 0.2, h: 0.24,
      fontSize: 10, fontFace: TF, color: r.color, bold: true,
    });
    s.addText(r.items.map(it => ({ text: it, options: { bullet: true, breakLine: true, paraSpaceAfter: 1 } })), {
      x: x + 0.14, y: y + 0.32, w: cardW - 0.24, h: cardH - 0.38,
      fontSize: 8, fontFace: BF, color: C.text, valign: "top", fit: "shrink",
    });
  });
  s.addNotes("【讲解节奏】业务规则第一部分，4个规则卡片，建议用时4分钟。\n【重点强调】联网核查规则要逐条讲清楚，哪些情况通过，哪些不通过。特别是\"无照片拒绝交易\"这一条很关键。\n【授权触发】税收居民标识变化是智能授权的触发条件。不是所有信息修改都需要授权，只有税收居民标识变了才触发。\n【名单管理】名单客户是直接拒绝的，没有商量余地，系统自动校验。");
  footer(s);
}

// ============================================================
// SLIDE 14: BUSINESS RULES (2)
// ============================================================
{
  let s = pres.addSlide();
  header(s, 14, "业务规则（二）", "身份核实 · 手机号 · 九要素 · 验证码 · 其他重要规则");

  const rules = [
    {
      title: "身份核实要求", color: C.primary, icon: "✅",
      items: [
        "客户身份需进行联网核查、人脸识别",
        "需留存身份证件影像资料",
        "人脸识别不通过时，由上一级管理人员现场审核",
      ],
    },
    {
      title: "手机号管理", color: C.secondary, icon: "📱",
      items: [
        "联系电话与个人身份证件号码一一对应关系",
        "同一手机号支持最多绑定三个客户",
        "客户已有主手机号，可一键同步为留存联系手机号，无需验证短信验证码",
        "手机号核查三种结果：属于客户、不属于客户、查无数据（仅提示不拦截）",
      ],
    },
    {
      title: "九要素信息", color: C.success, icon: "📝",
      items: [
        "自然人客户\"身份基本信息\"至少包含九要素",
        "姓名、性别、国籍、职业",
        "住所地或工作单位地址、联系方式",
        "证件类型、证件号码、证件有效期",
      ],
    },
    {
      title: "验证码规则", color: C.warning, icon: "🔢",
      items: [
        "允许客户最多获取5次验证码，间隔60秒",
        "每次验证码验证错误次数不得超过5次（60秒内）",
        "第5次错误报错提示",
      ],
    },
    {
      title: "新建客户三要素验证", color: C.rolePurple, icon: "🔑",
      items: [
        "通过\"证件号码\"、\"姓名\"和\"证件类型\"三要素进行唯一性验证",
        "新建客户（无银行账户情况下），无需签署《个人税收居民身份声明文件》",
        "客户已有银行账户，信息修改时需补充签署",
      ],
    },
    {
      title: "证件到期提醒", color: C.danger, icon: "⏰",
      items: [
        "预留证件到期前一个月，短信提醒客户更新身份信息",
        "到期后三个月未更新则进行账户止付",
        "职业选择\"Y0000\"时，视同客户信息不完整，系统提示转到信息维护",
      ],
    },
  ];

  const cardW = 2.85, cardH = 1.6, gapX = 0.2, gapY = 0.15;
  rules.forEach((r, i) => {
    let col = i % 3, row = Math.floor(i / 3);
    let x = M + col * (cardW + gapX);
    let y = 1.05 + row * (cardH + gapY);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.05,
      line: { color: r.color, width: 1.2 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.05, h: cardH, fill: { color: r.color },
    });
    s.addText(`${r.icon} ${r.title}`, {
      x: x + 0.12, y: y + 0.05, w: cardW - 0.18, h: 0.22,
      fontSize: 9, fontFace: TF, color: r.color, bold: true,
    });
    s.addText(r.items.map(it => ({ text: it, options: { bullet: true, breakLine: true, paraSpaceAfter: 1 } })), {
      x: x + 0.12, y: y + 0.28, w: cardW - 0.2, h: cardH - 0.34,
      fontSize: 7.5, fontFace: BF, color: C.text, valign: "top", fit: "shrink",
    });
  });

  // Bottom red line
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.6, w: CW, h: 0.5,
    fill: { color: C.dark }, rectRadius: 0.04,
  });
  s.addText("红线提醒：证件到期3个月未更新 → 账户止付  |  验证码5次错误 → 锁定  |  名单客户 → 拒绝交易", {
    x: M + 0.15, y: 4.6, w: CW - 0.3, h: 0.5,
    fontSize: 9, fontFace: TF, color: C.accent, bold: true,
    align: "center", valign: "middle",
  });
  s.addNotes("【讲解节奏】业务规则第二部分，6个规则卡片，建议用时5分钟。\n【重点强调】九要素是反洗钱的硬性要求，必须齐全。手机号最多绑定3个客户，这是人行规定。\n【红线提醒】证件到期3个月未更新会导致账户止付，这个影响很大，一定要提醒客户及时更新。验证码5次错误后会锁定，60秒后才能重试。\n【新建vs修改】新建客户和修改客户的税收居民声明要求不同——新建且无账户不需要签，已有账户修改时需要补签。");
  footer(s);
}

// ============================================================
// SLIDE 15: FIELD INPUT RULES
// ============================================================
{
  let s = pres.addSlide();
  header(s, 15, "字段输入规则", "证件类型选择 · 身份证号码输入规范");

  // Left: document types table
  s.addText("证件类型代码表（部分）", {
    x: M, y: 1.0, w: 4.3, h: 0.3,
    fontSize: 12, fontFace: TF, color: C.primary, bold: true,
  });

  const docRows = [
    [
      { text: "代码", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
      { text: "证件类型", options: { bold: true, color: C.white, fill: { color: C.primary } } },
    ],
    ["111", "居民身份证"],
    ["112", "临时身份证"],
    ["121", "户口簿"],
    ["131", "中国护照"],
    ["132", "外国护照"],
    ["141", "士兵证"],
    ["142", "军官证"],
    ["161", "港澳居民来往内地通行证"],
    ["163", "台湾居民来往大陆通行证"],
    ["171", "外国人永久居留证"],
    ["181", "驾驶证"],
  ];
  s.addTable(docRows, {
    x: M, y: 1.35, w: 4.3,
    colW: [1, 3.3],
    border: { type: "solid", pt: 1, color: "BDC3C7" },
    rowH: 0.3,
    fontSize: 8, fontFace: BF, color: C.text,
    valign: "middle",
    autoPage: false,
  });

  // Right: input rules
  infoCard(s, M + 4.7, 1.0, 4.3, 1.5, "身份证号码输入规范", [
    "当证件类型选择居民身份证、临时身份证和户口簿时适用",
    "证件号码需输入18位数字身份证号",
    "仅最后一位可出现\"X\"",
    "不能出现空格、标点符号等不规范字符",
  ], { border: C.danger, fontSize: 9 });

  // Auto-fill fields
  infoCard(s, M + 4.7, 2.65, 4.3, 1.7, "自动提取字段（不可修改）", [
    "姓名、国籍、性别",
    "证件类型、证件号码",
    "证件有效起始日、证件有效到期日",
    "签发机关、民族、出生日期",
    "共 10 项信息由证件自动读取",
  ], { border: C.success, fontSize: 9 });

  // Bottom: full document type note
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.5, w: CW, h: 0.6,
    fill: { color: C.infoBg }, rectRadius: 0.04,
    line: { color: C.infoBorder, width: 1 },
  });
  s.addText([
    { text: "完整证件类型：", options: { bold: true, color: C.secondary } },
    { text: "111居民身份证、112临时身份证、121户口簿、131中国护照、132外国护照、141士兵证、142军官证、143文职干部证、144军官退休证、151武警士兵证、152警官证、153武警干部证、154武警军官退休证、155武警文职干部退休证、161港澳居民来往内地通行证、162港澳居民身份证、163台湾居民来往大陆通行证、164身份证、165港澳台居民居住证、166便民出境通行证、171外国人永久居留证、172外国人居留证、181驾驶证、182学员证、183退休证、184工作证、185执行公务证、186回乡证、187教师资格证、191个人其他证件、192解放军干部证", options: { color: C.text } },
  ], {
    x: M + 0.12, y: 4.5, w: CW - 0.24, h: 0.6,
    fontSize: 7, fontFace: BF,
    valign: "middle", fit: "shrink",
  });
  s.addNotes("【讲解节奏】字段输入规则，建议用时3分钟。\n【重点记忆】18位身份证号，只有最后一位可以是X。不能有空格和标点。\n【自动提取字段】证件读取后有10项信息是自动填充的，而且不能修改。柜员需要核对这些信息是否与证件一致。\n【完整清单】证件类型非常多（30+种），不需要全部记住，常用的重点掌握即可。遇到不常见的可以查代码表。");
  footer(s);
}

// ============================================================
// SLIDE 16: VOUCHER MANAGEMENT
// ============================================================
{
  let s = pres.addSlide();
  header(s, 16, "凭证管理", "凭证内容 · 生成规则 · 归档流程");

  // Left card: voucher content
  infoCard(s, M, 1.05, 4.3, 2.0, "客户确认凭证内容", [
    { text: "客户需确认的信息：", options: { bold: true, color: C.primary } },
    "证件号码、客户姓名",
    "住所地或者工作单位地址",
    "有效期限、证件类型",
    "性别、国籍、职业",
    "联系方式、税收居民标志",
    { text: "确认完成后进行签字", options: { bold: true, color: C.accent } },
  ], { border: C.primary, fontSize: 9 });

  // Right card: voucher generation
  infoCard(s, M + 4.5, 1.05, 4.5, 2.0, "凭证生成规则", [
    { text: "对外生成：客户回单", options: { bold: true, color: C.secondary } },
    "纸质回单：柜面/自助设备打印",
    "电子回单：手机银行扫码查看",
    { text: "对内生成：电子凭证", options: { bold: true, color: C.success } },
    "系统自动生成，内部留存",
    "包含完整业务信息和签名影像",
  ], { border: C.secondary, fontSize: 9 });

  // Bottom: archiving process
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 3.25, w: CW, h: 1.8,
    fill: { color: C.dark }, rectRadius: 0.06,
  });
  s.addText("电子档案归档流程", {
    x: M + 0.2, y: 3.35, w: CW - 0.4, h: 0.3,
    fontSize: 12, fontFace: TF, color: C.accent, bold: true,
  });

  const archiveSteps = [
    { n: 1, t: "业务凭证", d: "业务办理产生的所有凭证" },
    { n: 2, t: "客户影像", d: "人脸识别、签名等影像资料" },
    { n: 3, t: "身份证影像", d: "证件读取留存的影像" },
    { n: 4, t: "联网核查", d: "联网核查结果记录" },
    { n: 5, t: "质检归档", d: "系统质检 → 自动归档" },
  ];

  const stepW = 1.6, stepH = 1.0, gap = 0.1;
  let startX = M + 0.2 + (CW - 0.4 - (stepW * 5 + gap * 4)) / 2;
  archiveSteps.forEach((st, i) => {
    let x = startX + i * (stepW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 3.75, w: stepW, h: stepH,
      fill: { color: C.white }, rectRadius: 0.04,
    });
    s.addText(`${st.n}. ${st.t}`, {
      x: x + 0.05, y: 3.8, w: stepW - 0.1, h: 0.28,
      fontSize: 9, fontFace: TF, color: C.primary, bold: true,
      align: "center",
    });
    s.addText(st.d, {
      x: x + 0.05, y: 4.1, w: stepW - 0.1, h: 0.6,
      fontSize: 7.5, fontFace: BF, color: C.text,
      align: "center", valign: "top", fit: "shrink",
    });
    if (i < 4) {
      s.addText("→", {
        x: x + stepW - 0.02, y: 4.05, w: gap + 0.04, h: 0.3,
        fontSize: 12, fontFace: TF, color: C.accent, bold: true,
        align: "center", valign: "middle",
      });
    }
  });
  s.addNotes("【讲解节奏】凭证管理，建议用时2分钟。\n【凭证内容】客户需要确认的信息包括10项，确认后签字。柜员要注意客户签字是否符合要求。\n【归档流程】业务办理结束后，系统通过无纸化档案系统将四类资料（业务凭证、客户影像、身份证影像、联网核查）进行电子档案归档。这个过程是自动的，柜员不需要手动操作。\n【质检环节】系统会先进行质检，质检通过后才归档。如果质检不通过，系统会有提示。");
  footer(s);
}

// ============================================================
// SLIDE 17: QUICK REFERENCE TABLE
// ============================================================
{
  let s = pres.addSlide();
  header(s, 17, "操作速查表", "柜面8步操作快速参考");

  const rows = [
    [
      { text: "步骤", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
      { text: "操作内容", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
      { text: "关键要点", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
      { text: "注意事项", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
    ],
    ["1", "输入交易代码030401进入场景", "交易代码或场景名称", "—"],
    ["2", "选择证件类型+读取+人脸识别", "多种证件类型支持", "户口簿有效期自动计算填充"],
    ["3", "代理办理选择+代理人核实", "是/否选择", "16岁以下可代理；16岁以上仅特殊人群可代理"],
    ["4", "进入客户信息维护界面查看", "核对要素信息", "10项自动提取字段不可修改"],
    ["5", "详细信息/税收/地址填写", "详细信息+税收+现居地址", "证件地址可一键同步；手机号需验证码"],
    ["6", "交易确认+签名/指纹", "柜外清正楷签名", "签名需柜员审核；不符合需重签"],
    ["7", "凭证及回单打印", "纸质/电子回单", "电子回单可手机银行扫码查看"],
    ["8", "客户评价", "3级评价", "不评价默认\"未评价\"；系统自动归档"],
  ];

  const colW = [0.5, 2.5, 2.3, 3.7];
  s.addTable(rows, {
    x: M, y: 1.05, w: CW, h: 3.2,
    colW,
    border: { type: "solid", pt: 1, color: "BDC3C7" },
    rowH: 0.4,
    fontSize: 8, fontFace: BF, color: C.text,
    valign: "middle",
    autoPage: false,
  });

  // Key rules summary
  infoCard(s, M, 4.0, 4.3, 0.9, "核心规则速记", [
    "三要素验证：证件号码+姓名+证件类型",
    "九要素信息：反洗钱最低要求",
    "手机号绑定：最多3个客户",
  ], { border: C.primary, fontSize: 8 });

  infoCard(s, M + 4.5, 4.0, 4.5, 0.9, "红线提醒", [
    "名单客户 → 直接拒绝交易",
    "证件到期3个月未更新 → 账户止付",
    "验证码5次错误 → 锁定",
  ], { border: C.danger, fontSize: 8 });

  s.addNotes("【复习用途】这页是操作速查表，可以打印出来贴在工位上，日常操作随时查阅。\n【复习提问】现在我们做几个自测题：\n1. 交易代码是多少？\n2. 柜面操作有几步？\n3. 代理办理的年龄限制是什么？\n4. 自动提取的字段有多少项？\n5. 什么情况下触发智能授权？\n6. 验证码最多能错几次？");
  footer(s);
}

// ============================================================
// SLIDE 18: CLOSING
// ============================================================
{
  let s = pres.addSlide();
  s.background = { color: C.dark };
  // decorative left bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.08, h: H, fill: { color: C.accent },
  });

  s.addText("培训总结", {
    x: 0.8, y: 1.2, w: 8.4, h: 0.7,
    fontSize: 36, fontFace: TF, color: C.white, bold: true,
    align: "left", valign: "middle",
  });

  // accent line
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 2.0, w: 2, h: 0.03, fill: { color: C.accent },
  });

  const summaryPoints = [
    "掌握 030401 交易代码及三渠道操作流程",
    "牢记九要素信息要求及三要素验证规则",
    "严格执行联网核查、人脸识别、签名审核规范",
    "关注证件到期提醒，及时通知客户更新信息",
    "遇到不确定情况，及时请教主管或查阅操作手册",
  ];

  s.addText(summaryPoints.map(p => ({ text: p, options: { bullet: true, breakLine: true, paraSpaceAfter: 8 } })), {
    x: 0.8, y: 2.3, w: 8.4, h: 2.0,
    fontSize: 14, fontFace: BF, color: "D5F5E3",
    valign: "top",
  });

  s.addText("交易代码 030401  |  权限：业务柜员  |  个人客户信息维护 · 操作培训手册", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.3,
    fontSize: 11, fontFace: BF, color: "82E0AA",
    align: "left", valign: "middle",
  });
  s.addNotes("【培训总结】感谢大家参加本次培训。希望大家通过今天的学习，能够熟练掌握个人客户信息维护的操作流程，严格遵守各项业务规则。\n【行动号召】培训结束后，请大家在实际操作中严格遵守注意事项与红线要求。遇到不确定的情况，及时请教老员工或主管。\n【后续支持】有任何问题可以随时联系培训组，也可以查阅完整的操作手册。祝大家工作顺利！");
}

// ============================================================
// SAVE
// ============================================================
pres.writeFile({ fileName: "个人客户信息维护_操作培训手册.pptx" })
  .then(fn => console.log("Generated: " + fn))
  .catch(err => console.error("Error:", err));
