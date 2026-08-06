const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.author = "AI Assistant";
pres.title = "综合个人开户 · 柜面操作培训";
pres.subject = "操作培训";

// ============================================================
// DIMENSIONS
// ============================================================
pres.layout = "LAYOUT_16x9";
const W = 10, H = 5.625;
const M = 0.5;
const CW = W - 2 * M; // content width = 9

// ============================================================
// COLORS — Ocean Gradient + Orange Accent
// ============================================================
const C = {
  primary: "065A82",
  secondary: "1C7293",
  dark: "21295C",
  accent: "E86A33",
  bg: "F5F7FA",
  white: "FFFFFF",
  text: "1F2937",
  textLight: "6B7280",
  success: "2D9C5E",
  warning: "E8A838",
  danger: "DC2626",
  cardBg: "FFFFFF",
  roleCustomer: "2D9C5E",
  roleHall: "E8A838",
  roleTeller: "065A82",
  roleSystem: "1C7293",
  roleAuth: "E86A33",
  noteBg: "FFF7ED",
  noteBorder: "E86A33",
  errorBg: "FEF2F2",
  errorBorder: "DC2626",
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
    x: M, y: 0.18, w: 0.38, h: 0.38,
    fill: { color: C.primary }, rectRadius: 0.05,
  });
  slide.addText(String(num), {
    x: M, y: 0.18, w: 0.38, h: 0.38,
    fontSize: 14, fontFace: TF, color: C.white, bold: true,
    align: "center", valign: "middle",
  });
  // title
  slide.addText(title, {
    x: M + 0.48, y: 0.15, w: CW - 0.96, h: 0.42,
    fontSize: 20, fontFace: TF, color: C.dark, bold: true,
    valign: "middle", fit: "shrink",
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: M + 0.48, y: 0.52, w: CW - 0.96, h: 0.24,
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
      x: x + 0.38, y: y - 0.02, w: 4, h: 0.36,
      fontSize: 12, fontFace: TF, color: C.dark, bold: true,
      valign: "middle", fit: "shrink",
    });
  }
}

// ============================================================
// HELPER: footer
// ============================================================
function footer(slide) {
  slide.addText("综合个人开户 · 柜面操作培训", {
    x: M, y: H - 0.3, w: CW, h: 0.22,
    fontSize: 7, fontFace: BF, color: C.textLight,
    align: "center",
  });
}

// ============================================================
// HELPER: business process overview slide (Slide 5 & 6)
// ============================================================
function processOverviewSlide(pageNum, title, subtitle, steps) {
  let s = pres.addSlide();
  header(s, pageNum, title, subtitle);
  const cardW = 1.68, cardH = 3.3, gap = 0.1;
  let startX = (W - (cardW * 5 + gap * 4)) / 2;
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
      x: x + 0.08, y: 1.75, w: cardW - 0.16, h: 0.5,
      fontSize: 10, fontFace: TF, color: C.dark, bold: true,
      align: "center", valign: "top", fit: "shrink",
    });
    s.addText(st.d, {
      x: x + 0.08, y: 2.3, w: cardW - 0.16, h: 2.0,
      fontSize: 8, fontFace: BF, color: C.text,
      valign: "top", fit: "shrink",
    });
  });
  for (let i = 0; i < 4; i++) {
    let ax = startX + (i + 1) * cardW + i * gap + gap / 2 - 0.11;
    s.addText("→", {
      x: ax, y: 2.5, w: 0.3, h: 0.3,
      fontSize: 14, fontFace: TF, color: C.accent, bold: true,
      align: "center", valign: "middle",
    });
  }
  return s;
}

// ============================================================
// HELPER: media selection slide (Slide 11 & 12)
// ============================================================
function mediaSelectionSlide(pageNum, title, subtitle, cards) {
  let s = pres.addSlide();
  header(s, pageNum, title, subtitle);
  stepBadge(s, M, 1.0, 3, pageNum === 11 ? "选择开户介质（6种可选）" : "选择开户介质（续）");
  const cardW = 2.85, cardH = 3.5, gap = 0.22;
  cards.forEach((cd, i) => {
    let x = M + i * (cardW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.4, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.05,
      line: { color: cd.border, width: 1 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.4, w: cardW, h: 0.05, fill: { color: cd.border },
    });
    s.addText(cd.title, {
      x: x + 0.1, y: 1.48, w: cardW - 0.2, h: 0.25,
      fontSize: 11, fontFace: TF, color: cd.border, bold: true,
    });
    s.addText("操作步骤：", {
      x: x + 0.1, y: 1.76, w: cardW - 0.2, h: 0.18,
      fontSize: 8, fontFace: TF, color: C.dark, bold: true,
    });
    s.addText(cd.ops.map(o => ({ text: o, options: { bullet: true, breakLine: true, paraSpaceAfter: 1 } })), {
      x: x + 0.1, y: 1.95, w: cardW - 0.2, h: 0.85,
      fontSize: 8, fontFace: BF, color: C.text, valign: "top", fit: "shrink",
    });
    s.addText("关键字段：", {
      x: x + 0.1, y: 2.85, w: cardW - 0.2, h: 0.18,
      fontSize: 8, fontFace: TF, color: C.accent, bold: true,
    });
    s.addText(cd.fields.map(f => ({ text: f, options: { bullet: true, breakLine: true, paraSpaceAfter: 0 } })), {
      x: x + 0.1, y: 3.04, w: cardW - 0.2, h: 0.65,
      fontSize: 7, fontFace: BF, color: C.text, valign: "top", fit: "shrink",
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.05, y: 3.75, w: cardW - 0.1, h: 0.015, fill: { color: C.textLight },
    });
    s.addText(cd.note, {
      x: x + 0.1, y: 3.79, w: cardW - 0.2, h: 1.0,
      fontSize: 7, fontFace: BF, color: C.textLight, valign: "top", fit: "shrink",
    });
  });
  return s;
}

// ============================================================
// SLIDE 1: COVER
// ============================================================
{
  let s = pres.addSlide();
  s.background = { path: "images/bg-cover_16x9.jpg" };
  // dark overlay for text readability
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: H,
    fill: { color: C.dark, transparency: 35 },
  });
  s.addText("综合个人开户", {
    x: 0.8, y: 1.5, w: 8.4, h: 0.8,
    fontSize: 40, fontFace: TF, color: C.white, bold: true,
    align: "center", valign: "middle",
  });
  s.addText("柜面操作培训手册", {
    x: 0.8, y: 2.35, w: 8.4, h: 0.5,
    fontSize: 22, fontFace: TF, color: "B0D4E8",
    align: "center", valign: "middle",
  });
  // accent line
  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.5, y: 3.0, w: 3, h: 0.03, fill: { color: C.accent },
  });
  s.addText("面向柜面业务人员 · 操作培训用", {
    x: 0.8, y: 3.2, w: 8.4, h: 0.35,
    fontSize: 13, fontFace: BF, color: "D1D5DB",
    align: "center", valign: "middle",
  });
  s.addText("交易代码 030601  |  权限：业务柜员", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.3,
    fontSize: 11, fontFace: BF, color: "9CA3AF",
    align: "center", valign: "middle",
  });
  s.addNotes("【开场白】各位同事大家好，今天我们来学习综合个人开户的柜面操作流程。本培训面向柜面业务人员，目标是让大家掌握030601交易的完整操作。培训时长约45分钟，包含操作演示和注意事项讲解。请大家关注每一步的关键字段和红色注意事项。");
}

// ============================================================
// SLIDE 2: TABLE OF CONTENTS
// ============================================================
{
  let s = pres.addSlide();
  header(s, 2, "培训内容概览", "目录");
  const items = [
    { num: "01", title: "场景说明与渠道对比", desc: "场景定义、账户类型、5种开户渠道" },
    { num: "02", title: "业务办理流程概述", desc: "完整操作介绍，10个关键环节" },
    { num: "03", title: "柜面操作步骤详解", desc: "7步操作流程，含四要素与注意事项" },
    { num: "04", title: "特殊场景与红线清单", desc: "代理办理、未成年、大额授权等" },
    { num: "05", title: "操作速查表", desc: "7步骤快速参考" },
  ];
  const cardW = 1.68, cardH = 3.2, gap = 0.1;
  let startX = (W - (cardW * 5 + gap * 4)) / 2;
  items.forEach((it, i) => {
    let x = startX + i * (cardW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.2, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.06,
      line: { color: C.secondary, width: 1 },
    });
    // top accent
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.2, w: cardW, h: 0.06, fill: { color: C.primary },
    });
    s.addText(it.num, {
      x: x + 0.1, y: 1.4, w: cardW - 0.2, h: 0.5,
      fontSize: 28, fontFace: TF, color: C.primary, bold: true,
      align: "center",
    });
    s.addText(it.title, {
      x: x + 0.1, y: 2.0, w: cardW - 0.2, h: 0.6,
      fontSize: 12, fontFace: TF, color: C.dark, bold: true,
      align: "center", valign: "top", fit: "shrink",
    });
    s.addText(it.desc, {
      x: x + 0.1, y: 2.7, w: cardW - 0.2, h: 1.4,
      fontSize: 9, fontFace: BF, color: C.textLight,
      align: "center", valign: "top", fit: "shrink",
    });
  });
  s.addNotes("【章节过渡】本次培训分为五大模块：场景说明与渠道对比、业务办理流程概述、柜面操作步骤详解、特殊场景与红线清单、操作速查表。我们先从场景说明开始，了解综合开户的业务背景。");
  footer(s);
}

// ============================================================
// SLIDE 3: SCENE DESCRIPTION
// ============================================================
{
  let s = pres.addSlide();
  header(s, 3, "场景说明", "综合个人开户场景定义与账户类型");

  // Left card: scene definition
  infoCard(s, M, 1.05, 4.3, 2.0, "场景定义", [
    "客户到我社办理银行卡、存折、存单开户业务",
    "同时为客户提供多种签约服务开通",
    "客户可选择开立结算账户（Ⅰ、Ⅱ、Ⅲ类户）或储蓄账户",
    "我社为客户提供结算户的零余额开户",
  ], { border: C.primary, fontSize: 10 });

  // Right card: account types
  infoCard(s, M + 4.5, 1.05, 4.5, 2.0, "账户类型", [
    { text: "结算账户 — Ⅰ类户（全功能）", options: { bold: true, color: C.primary } },
    { text: "结算账户 — Ⅱ类户（限制功能）", options: { bold: true, color: C.primary } },
    { text: "结算账户 — Ⅲ类户（小额功能）", options: { bold: true, color: C.primary } },
    { text: "储蓄账户", options: { bold: true, color: C.secondary } },
    { text: "结算户支持零余额开户", options: { color: C.accent } },
  ], { border: C.secondary, fontSize: 10 });

  // Bottom: key info bar
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 3.3, w: CW, h: 1.5,
    fill: { color: C.dark }, rectRadius: 0.06,
  });
  s.addText("核心要点", {
    x: M + 0.2, y: 3.4, w: 2, h: 0.28,
    fontSize: 11, fontFace: TF, color: C.accent, bold: true,
  });
  const pts = [
    "综合开户 = 开户 + 签约服务一站式办理",
    "支持银行卡、存折、存单、大额存单、一本通、电子账户 6 种介质",
    "结算账户与储蓄账户均可开立，结算户支持零余额开户",
    "涉及客户身份核实、信息维护、介质绑定、密码设置、签约服务全流程",
  ];
  s.addText(pts.map(p => ({ text: p, options: { bullet: true, breakLine: true, paraSpaceAfter: 3 } })), {
    x: M + 0.2, y: 3.7, w: CW - 0.4, h: 1.0,
    fontSize: 10, fontFace: BF, color: C.white,
    valign: "top", fit: "shrink",
  });
  s.addNotes("【背景故事】上周有位客户来网点要求同时开立银行卡和存折，还要求开通短信通知。如果没有综合开户交易，柜员需要分别走3个交易，耗时约15分钟。使用030601综合开户后，一站式办理仅需5分钟。\n【实际案例】特别注意结算账户和储蓄账户的区别——结算户支持零余额开户，储蓄户必须存入资金。");
  footer(s);
}

// ============================================================
// SLIDE 4: CHANNEL COMPARISON
// ============================================================
{
  let s = pres.addSlide();
  header(s, 4, "开户渠道对比", "5种渠道的办理方式、介质与账户类型");

  const rows = [
    [
      { text: "渠道", options: { bold: true, color: C.white, fill: { color: C.primary } } },
      { text: "办理方式", options: { bold: true, color: C.white, fill: { color: C.primary } } },
      { text: "可办介质", options: { bold: true, color: C.white, fill: { color: C.primary } } },
      { text: "账户类型", options: { bold: true, color: C.white, fill: { color: C.primary } } },
    ],
    ["智能柜员机", "客户自助", "—", "I、II类户、储蓄账户"],
    ["移动PAD", "厅堂服务人员", "银行卡、存折", "I、II类户"],
    ["柜面", "柜员", "银行卡、存折、存单", "I、II类户、储蓄账户"],
    ["社保卡特殊卡种", "大堂经理打印", "社保卡", "开户后专用机打印"],
    ["手机银行", "客户自助", "电子账户", "II、III类户（需已办I类户）"],
  ];
  const colW = [1.6, 2.0, 2.5, 2.9];
  s.addTable(rows, {
    x: M, y: 1.1, w: CW, h: 2.6,
    colW,
    border: { type: "solid", pt: 1, color: C.textLight },
    rowH: 0.5,
    fontSize: 10, fontFace: BF, color: C.text,
    align: "center", valign: "middle",
    autoPage: false,
  });

  // Note below table
  noticeBox(s, M, 3.9, CW, 1.0, [
    "柜面渠道功能最全，支持全部6种介质和全部账户类型",
    "手机银行仅限已办理I类户的存量客户开立II、III类户电子账户",
    "社保卡特殊卡种在开户完成后需由大堂经理在专用打印信息机上进行打印",
  ]);
  s.addNotes("【讲解要点】5种渠道中，柜面功能最全，支持全部6种介质。手机银行仅限已办I类户的存量客户开立II、III类户。\n【易混淆点】社保卡特殊卡种不是在柜面直接完成，开户后还需大堂经理在专用打印信息机上打印。\n【互动提问】如果客户在手机银行开II类户但未办I类户，系统会怎么提示？");
  footer(s);
}

// ============================================================
// SLIDE 5: BUSINESS PROCESS OVERVIEW (Part 1)
// ============================================================
{
  const steps = [
    { n: 1, t: "客户到达与身份识别", d: "客户到达网点后，由取号机、厅堂服务人员手持Pad对客户身份进行识别，结合网点忙闲情况和客户情况进行分流或展开营销。" },
    { n: 2, t: "身份核实与确认", d: "通过联网核查、人脸识别等方式对客户身份进行核实和确认。人脸识别不通过时，由当前操作人员、辅助人员的上一级管理人员进行现场审核。" },
    { n: 3, t: "客户信息维护", d: "判断是否为新建客户。若为新客户或客户九要素信息缺失，则进行客户信息维护。" },
    { n: 4, t: "协议确认与介质选择", d: "客户经过开户协议阅读和确认后，选择需要开通的介质类型，再由客户勾选需要开的业务种类。" },
    { n: 5, t: "账户关联与密码设置", d: "系统自动关联账户种类（结算户/储蓄户）。储蓄户弹出资金存入界面，结算户不经过该界面。进行个人账户协议签订，系统自动绑定介质与账户，引导客户设置取款密码和查询密码。" },
  ];
  let s = processOverviewSlide(5, "业务办理流程概述（一）", "操作介绍 · 第1-5环节", steps);
  s.addNotes("【讲解节奏】本页讲解业务流程的前5个环节，建议用时3分钟。\n【易混淆点】注意第3步\u201C客户信息维护\u201D仅在新建客户或九要素信息缺失时触发，不要每次都做。\n【互动提问】大家可以想一想，储蓄户和结算户在账户关联环节有什么区别？");
  footer(s);
}

// ============================================================
// SLIDE 6: BUSINESS PROCESS OVERVIEW (Part 2)
// ============================================================
{
  const steps = [
    { n: 6, t: "签约服务选择", d: "结算账户开立后，客户选择签约服务：互联网数字银行、短信通知、非柜面转账限额。界面显示服务类型、名称、简述。客户勾选后回显签约信息、风险、事项，阅读协议后确认签约。" },
    { n: 7, t: "集中作业授权", d: "存入资金超过50万以上，或是代理办理，则触发集中作业授权人员进行授权。" },
    { n: 8, t: "电子签名确认", d: "系统回显客户声明、风险提示及综合个人开户业务信息电子凭证。客户确认后进行电子签名（支持手写正楷和指纹签名）。电子签名需经柜员审核，不符合要求需重新签字。" },
    { n: 9, t: "回单打印与评价", d: "根据客户需求提供纸质回单、电子回单打印服务，同时客户对本次服务进行评价。" },
    { n: 10, t: "凭证质检与归档", d: "业务办理结束后，系统对凭证进行质检并进行自动归档处理。" },
  ];
  let s = processOverviewSlide(6, "业务办理流程概述（二）", "操作介绍 · 第6-10环节", steps);
  s.addNotes("【讲解节奏】本页讲解后5个环节，建议用时3分钟，重点讲第7步集中授权。\n【易混淆点】第7步授权触发条件有两个：资金超50万 或 代理办理，两者满足其一即触发。\n【互动提问】电子签名不合规时该怎么处理？需要重新签字还是重新走流程？");
  footer(s);
}

// ============================================================
// SLIDE 7: SWIMLANE DIAGRAM
// ============================================================
{
  let s = pres.addSlide();
  header(s, 7, "职能流程图", "按角色分色泳道图 · 客户/厅堂人员/柜员/系统/授权人员");

  const lanes = [
    { name: "客户", color: C.roleCustomer },
    { name: "厅堂人员", color: C.roleHall },
    { name: "柜员", color: C.roleTeller },
    { name: "系统", color: C.roleSystem },
    { name: "授权人员", color: C.roleAuth },
  ];

  const laneStartY = 1.1;
  const laneH = 0.62;
  const labelW = 0.95;
  const procX = M + labelW;
  const procW = W - M - procX;

  // Draw lanes — label area and process area separated to avoid overlap
  lanes.forEach((ln, i) => {
    let y = laneStartY + i * laneH;
    // process area background (alternating, only process area)
    s.addShape(pres.shapes.RECTANGLE, {
      x: procX, y, w: W - M - procX, h: laneH,
      fill: { color: i % 2 === 0 ? C.white : C.bg },
    });
    // lane label (separate from process area)
    s.addShape(pres.shapes.RECTANGLE, {
      x: M, y, w: labelW, h: laneH,
      fill: { color: ln.color },
    });
    s.addText(ln.name, {
      x: M, y, w: labelW, h: laneH,
      fontSize: 10, fontFace: TF, color: C.white, bold: true,
      align: "center", valign: "middle",
    });
    // separator line between lanes
    if (i < lanes.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: M, y: y + laneH, w: W - 2 * M, h: 0,
        line: { color: C.textLight, width: 0.5 },
      });
    }
  });

  // Process nodes (6 key steps, merged last two to fit boundary)
  const nodes = [
    { label: "客户到达\n网点", lane: 0, x: 0.1 },
    { label: "身份识别\n与分流", lane: 1, x: 1.35 },
    { label: "联网核查\n人脸识别", lane: 2, x: 2.6 },
    { label: "信息维护\n选介质\n设密码", lane: 3, x: 3.85 },
    { label: "大额授权\n(>50万\n/代理)", lane: 4, x: 5.1, dashed: true },
    { label: "签名确认\n打印评价\n归档", lane: 2, x: 6.35 },
  ];

  const nodeW = 1.1, nodeH = 0.45;
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
    { from: { lane: 0, x: 0.1 }, to: { lane: 1, x: 1.35 } },
    { from: { lane: 1, x: 1.35 }, to: { lane: 2, x: 2.6 } },
    { from: { lane: 2, x: 2.6 }, to: { lane: 3, x: 3.85 } },
    { from: { lane: 3, x: 3.85 }, to: { lane: 4, x: 5.1 } },
    { from: { lane: 4, x: 5.1 }, to: { lane: 2, x: 6.35 } },
  ];
  arrowData.forEach(a => {
    let x1 = procX + a.from.x + nodeW;
    let y1 = laneStartY + a.from.lane * laneH + laneH / 2;
    let x2 = procX + a.to.x;
    let y2 = laneStartY + a.to.lane * laneH + laneH / 2;
    // horizontal then vertical L-shaped arrow
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
  const legY = 4.45;
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
    let lx = M + 0.55 + i * 1.1;
    s.addShape(pres.shapes.RECTANGLE, {
      x: lx, y: legY + 0.04, w: 0.18, h: 0.18,
      fill: { color: lg.color },
    });
    s.addText(lg.label, {
      x: lx + 0.22, y: legY, w: 0.8, h: 0.25,
      fontSize: 9, fontFace: BF, color: C.text, valign: "middle",
    });
  });
  // dashed node legend
  s.addShape(pres.shapes.OVAL, {
    x: M + 6.1, y: legY + 0.04, w: 0.18, h: 0.18,
    fill: { color: C.roleAuth },
    line: { color: C.danger, width: 1.5, dashType: "dash" },
  });
  s.addText("条件触发节点", {
    x: M + 6.35, y: legY, w: 1.5, h: 0.25,
    fontSize: 9, fontFace: BF, color: C.text, valign: "middle",
  });
  s.addNotes("【角色交接要点】重点关注3个跨泳道交接：①客户→厅堂人员：身份识别分流；②厅堂人员→柜员：引导至柜台；③柜员→授权人员：大额或代理触发授权。\n【跨部门协作】大额授权（>50万）需要集中作业授权人员在线审核，授权通过后柜员才能继续操作。如果授权人员不在线，客户需要等待。");
  footer(s);
}

// ============================================================
// SLIDE 8: COUNTER OPERATION OVERVIEW
// ============================================================
{
  let s = pres.addSlide();
  header(s, 8, "柜面操作总览", "交易代码 · 权限 · 7步操作流程");

  // Transaction code & permission bar
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 1.0, w: CW, h: 0.55,
    fill: { color: C.dark }, rectRadius: 0.05,
  });
  s.addText([
    { text: "交易代码：", options: { fontSize: 12, color: "B0D4E8" } },
    { text: "030601 个人综合开户", options: { fontSize: 12, color: C.white, bold: true } },
    { text: "    |    ", options: { fontSize: 12, color: "6B7280" } },
    { text: "交易权限：", options: { fontSize: 12, color: "B0D4E8" } },
    { text: "业务柜员", options: { fontSize: 12, color: C.white, bold: true } },
  ], {
    x: M + 0.2, y: 1.0, w: CW - 0.4, h: 0.55,
    fontFace: TF, align: "center", valign: "middle",
  });

  // 7 step cards
  const stepData = [
    { n: 1, t: "进入交易", d: "输入交易代码030601\n或场景名称" },
    { n: 2, t: "证件识别", d: "选择证件类型\n读取证件+人脸识别" },
    { n: 3, t: "选择介质", d: "6种介质：银行卡/存折\n存单/大额存单/一本通/电子账户" },
    { n: 4, t: "密码设置", d: "设置查询密码\n和取款密码" },
    { n: 5, t: "交易确认", d: "柜外清签名/指纹\n存折写磁" },
    { n: 6, t: "凭证打印", d: "纸质/电子回单\n打印" },
    { n: 7, t: "客户评价", d: "非常满意/满意\n不满意/未评价" },
  ];

  const cardW = 1.2, cardH = 2.5, gap = 0.09;
  let startX = (W - (cardW * 7 + gap * 6)) / 2;
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
      x: x + cardW / 2 - 0.2, y: 1.88, w: 0.4, h: 0.4,
      fill: { color: C.primary },
    });
    s.addText(String(st.n), {
      x: x + cardW / 2 - 0.2, y: 1.88, w: 0.4, h: 0.4,
      fontSize: 14, fontFace: TF, color: C.white, bold: true,
      align: "center", valign: "middle",
    });
    s.addText(st.t, {
      x: x + 0.05, y: 2.35, w: cardW - 0.1, h: 0.3,
      fontSize: 10, fontFace: TF, color: C.dark, bold: true,
      align: "center", valign: "middle", fit: "shrink",
    });
    s.addText(st.d, {
      x: x + 0.05, y: 2.68, w: cardW - 0.1, h: 1.4,
      fontSize: 8, fontFace: BF, color: C.text,
      align: "center", valign: "top", fit: "shrink",
    });
  });

  // arrow
  s.addText("→ → → → → → →", {
    x: startX, y: 4.35, w: cardW * 7 + gap * 6, h: 0.25,
    fontSize: 12, fontFace: TF, color: C.accent, bold: true,
    align: "center", valign: "middle",
  });
  s.addNotes("【讲解节奏】本页是7步操作的总览，建议用时2分钟快速过一遍，不展开细节。后续每一步都会有独立页面详细讲解。\n【关键记忆】交易代码030601是必须记住的，权限为业务柜员。7步流程：进入交易→证件识别→选介质→设密码→确认签名→打印→评价。");
  footer(s);
}

// ============================================================
// SLIDE 9: STEPS 1-2
// ============================================================
{
  let s = pres.addSlide();
  header(s, 9, "步骤 1-2：进入交易与证件识别", "操作说明 · 关键字段");

  // Step 1
  stepBadge(s, M, 1.0, 1, "进入交易界面");
  infoCard(s, M, 1.4, CW, 0.7, null, [
    "柜员成功登录后，输入交易代码“030601”或场景名称进入“个人综合开户”场景界面",
  ], { border: C.primary, fontSize: 10 });

  // Step 2
  stepBadge(s, M, 2.3, 2, "证件识别与人脸识别");
  infoCard(s, M, 2.7, 5.3, 1.8, "操作说明与关键字段", [
    { text: "操作说明：", options: { bold: true, color: C.primary } },
    "进入交易界面后，根据携带的有效证件选择对应的证件类型",
    "将有效身份证件放在证件读取区域，查看证件信息无误后点击“下一步”",
    "进行人脸识别，人脸识别通过后点击“下一步”",
    { text: "可选证件类型（4种）：", options: { bold: true, color: C.accent } },
    "二代身份证 / 港澳台居民居住证 / 外国人永久居留身份证 / 其他证件",
  ], { border: C.secondary, fontSize: 9 });

  // Error-prone box
  errorBox(s, M + 5.5, 2.7, 3.5, 1.8, [
    "证件类型选择错误会导致读取失败",
    "证件放置位置不正确影响信息读取",
    "人脸识别不通过时需上级管理人员现场审核",
    "务必核对证件信息无误后再点击“下一步”",
  ]);

  // Note: precautions on next slide
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.65, w: CW, h: 0.35,
    fill: { color: C.noteBg }, rectRadius: 0.04,
    line: { color: C.noteBorder, width: 1 },
  });
  s.addText("⚠ 步骤2注意事项较多（代理办理/未成年人/特殊客群），详见下一页", {
    x: M + 0.12, y: 4.65, w: CW - 0.24, h: 0.35,
    fontSize: 9, fontFace: TF, color: C.accent, bold: true,
    valign: "middle",
  });
  s.addNotes("【讲解节奏】步骤1简单带过（30秒），步骤2重点讲解（2分钟）。\n【易混淆点】4种证件类型要选对，选错会导致读取失败。人脸识别不通过时不要慌，由上一级管理人员现场审核即可。\n【互动提问】如果客户持有港澳台居民居住证，应该选哪种证件类型？");
  footer(s);
}

// ============================================================
// SLIDE 10: STEP 2 PRECAUTIONS — AGENT / MINORS / SPECIAL
// ============================================================
{
  let s = pres.addSlide();
  header(s, 10, "步骤2 注意事项：代理办理与特殊客群", "源文档原文逐条保留");

  // Scenario 1: Agent
  infoCard(s, M, 1.05, 4.3, 1.65, "① 代理人办理", [
    "本场景允许代理人代为办理，可通过柜面渠道进行",
    "需携带：授权委托书、代理人与被代理人的关系证明文件",
    "需携带：被代理人所在社区居委会（村民委员会）及以上组织或县级以上医院出具的特殊情况证明",
    "代理人办理时需对业务办理全过程进行双录",
    "留存代理人的人脸识别信息，与业务办理信息进行关联",
  ], { border: C.accent, fontSize: 8 });

  // Scenario 2: Minors
  infoCard(s, M + 4.5, 1.05, 4.5, 1.65, "② 未满16周岁人员开户", [
    "对于未满16周岁人员开户",
    "以监护人代理办理处理",
    "需提供监护人身份证明及监护关系证明",
  ], { border: C.warning, fontSize: 8 });

  // Scenario 3: Special circumstances
  infoCard(s, M, 2.85, CW, 1.65, "③ 特殊情况（重病/行动不便/无自理能力）", [
    "因身患重病、行动不便、无自理能力等无法自行前往银行的存款人办理个人开户，可采取上门服务方式办理",
    "也可由配偶、父母或成年子女凭以下材料代理办理：",
    "  • 合法的委托书",
    "  • 代理人与被代理人的关系证明文件",
    "  • 被代理人所在社区居委会（村民委员会）及以上组织或县级以上医院出具的特殊情况证明",
  ], { border: C.danger, fontSize: 8 });

  // Error-prone box
  errorBox(s, M, 4.55, CW, 0.7, [
    "代理办理必须双录并留存人脸识别信息，否则无法通过合规检查",
    "三种特殊情况均需提供关系证明 + 特殊情况证明，缺一不可",
  ]);
  s.addNotes("【合规风险强调】代理办理必须双录！这是合规底线，不双录将直接导致合规检查不通过。\n【真实案例】去年某网点代理办理未留存人脸识别信息，事后客户否认授权，因无双录记录无法举证，最终被通报批评。\n【处理话术】面对特殊客群（重病/行动不便），建议主动告知可提供上门服务，话术：\u201C您不方便到网点的话，我们可以安排上门服务，需要您提供委托书和关系证明。\u201D");
  footer(s);
}

// ============================================================
// SLIDE 11: STEP 3 PART 1 — Bank Card / Passbook / Certificate
// ============================================================
{
  const cards = [
    {
      title: "银行卡开户", border: C.primary,
      ops: ["选择\u201C银行卡开户\u201D", "选择卡种", "将新卡放置在IC卡读取区域进行读取"],
      fields: ["卡种", "（卡片信息自动读取）"],
      note: "卡片信息读取成功后，提醒客户阅读相关协议签字确认后点击\u201C下一步\u201D。完成后流程至步骤4。",
    },
    {
      title: "存折开户", border: C.secondary,
      ops: ["选择\u201C存折开户\u201D", "根据客户开户意愿填写各项信息", "系统提示进行现金\u201C冠字号码\u201D录入"],
      fields: ["业务种类", "通兑标志", "支取方式", "开户金额", "凭证号码", "转存方式", "利率类型", "期限"],
      note: "若开通一类户存折，需提醒客户阅读相关协议签字确认后点击\u201C下一步\u201D。完成后流程至步骤4。",
    },
    {
      title: "存单开户", border: C.warning,
      ops: ["选择\u201C存单开户\u201D", "根据客户开户意愿填写各项信息", "系统提示进行现金\u201C冠字号码\u201D录入"],
      fields: ["业务种类", "通兑标志", "支取方式", "开户金额", "凭证号码", "转存方式", "利率类型", "期限"],
      note: "填写完成后点击\u201C下一步\u201D。完成后流程至步骤4。",
    },
  ];
  let s = mediaSelectionSlide(11, "步骤3：选择开户介质（一）", "银行卡 / 存折 / 存单 · 操作说明与字段清单", cards);
  s.addNotes("【讲解节奏】本页讲前3种介质，每种约1分钟，建议共3分钟。\n【易混淆点】存折和存单的字段清单完全一致（8项），但完成后的操作不同——存折需协议签字确认，存单直接下一步。\n【互动提问】银行卡开户为什么字段最少？因为卡片信息是自动读取的。");
  footer(s);
}

// ============================================================
// SLIDE 12: STEP 3 PART 2 — Large CD / One-pass / Electronic
// ============================================================
{
  const cards = [
    {
      title: "大额存单开户", border: C.primary,
      ops: [
        "选择\u201C大额存单\u201D开户，进入产品类型选择界面",
        "勾选完成后点击\u201C下一步\u201D",
        "确认产品信息，选择\u201C科目来源\u201D",
        "选择\u201C存单形式\u201D（电子账户/纸质存单）",
      ],
      fields: [
        "科目来源：现金 → 输入凭证号码、开户金额",
        "科目来源：转账 → 介质读取识别",
        "存单形式：电子账户（不需凭证号码）",
        "存单形式：纸质存单（需输入凭证号码）",
      ],
      note: "选择存单形式后输入开户金额，点击下一步。完成后流程至步骤4。",
    },
    {
      title: "一本通开户", border: C.secondary,
      ops: [
        "选择\u201C一本通开户\u201D",
        "根据客户意愿填写开户各项信息",
        "填写完成后点击\u201C下一步\u201D",
      ],
      fields: ["通兑标志", "支取方式", "凭证号码"],
      note: "完成后流程至步骤4。",
    },
    {
      title: "电子账户开户", border: C.success,
      ops: [
        "选择\u201C电子账户开户\u201D",
        "根据客户意愿填写开户信息",
        "将IC卡放入读取区域点击\u201C读卡\u201D",
        "提醒客户阅读相关协议后签字确认",
      ],
      fields: ["卡种", "产品代码"],
      note: "成功读取卡片信息后，签字确认点击\u201C下一步\u201D。完成后流程至步骤4。",
    },
  ];
  let s = mediaSelectionSlide(12, "步骤3：选择开户介质（二）", "大额存单 / 一本通 / 电子账户 · 操作说明与字段清单", cards);
  s.addNotes("【讲解节奏】本页讲后3种介质，重点讲大额存单的科目来源分支，建议共4分钟。\n【易混淆点】大额存单的\u201C科目来源\u201D选择现金和转账，后续操作路径完全不同——现金直接输入凭证号码和金额，转账需要介质识别。\n【互动提问】电子账户开户为什么需要IC卡读取？因为要绑定实体卡片信息。");
  footer(s);
}

// ============================================================
// SLIDE 13: STEP 3 PRECAUTIONS
// ============================================================
{
  let s = pres.addSlide();
  header(s, 13, "步骤3 注意事项", "存单开户 · 储蓄账户 · 资金存入要求");

  // Main notice
  noticeBox(s, M, 1.1, CW, 1.0, [
    "若客户开立储蓄账户且介质为存单，则客户必须进行资金存入",
  ]);

  // Additional context cards
  infoCard(s, M, 2.35, 4.3, 1.8, "介质与账户类型关联", [
    { text: "结算账户：", options: { bold: true, color: C.primary } },
    "不经过开户资金存入界面",
    "系统自动关联，支持零余额开户",
    { text: "储蓄账户：", options: { bold: true, color: C.secondary } },
    "系统弹出储蓄账户开户资金存入界面",
    "必须进行资金存入（特别是存单介质）",
  ], { border: C.secondary, fontSize: 9 });

  infoCard(s, M + 4.5, 2.35, 4.5, 1.8, "6种介质完成后均进入步骤4", [
    "银行卡 → IC卡读取 + 协议签字 → 步骤4",
    "存折 → 填写8项字段 + 冠字号码 → 步骤4",
    "存单 → 填写8项字段 + 冠字号码 → 步骤4",
    "大额存单 → 选产品 + 科目来源 + 存单形式 → 步骤4",
    "一本通 → 填写3项字段 → 步骤4",
    "电子账户 → 填写2项字段 + IC卡读取 → 步骤4",
  ], { border: C.primary, fontSize: 8 });

  errorBox(s, M, 4.3, CW, 0.85, [
    "储蓄账户+存单组合容易遗漏资金存入步骤，导致开户失败",
    "存折和存单开户必须录入现金\u201C冠字号码\u201D，遗漏将影响合规检查",
    "大额存单的\u201C科目来源\u201D选择不同，后续操作路径不同（现金直接输入 vs 转账需介质识别）",
  ]);
  s.addNotes("【合规风险强调】储蓄账户+存单组合最容易遗漏资金存入步骤！如果漏做，开户流程无法完成。\n【真实案例】某新员工在开立储蓄存单时忘记资金存入，客户离开后发现账户未激活，投诉至网点主任。\n【易混淆点】6种介质完成后都进入步骤4，但存折和存单必须录入冠字号码，这是反洗钱要求。");
  footer(s);
}

// ============================================================
// SLIDE 14: STEP 4 — Password Setting
// ============================================================
{
  let s = pres.addSlide();
  header(s, 14, "步骤4：密码设置", "操作说明 · 关键字段 · 注意事项 · 易错点");

  stepBadge(s, M, 1.0, 4, "密码设置");

  // Operation description
  infoCard(s, M, 1.4, 4.3, 1.3, "操作说明", [
    "若选择密码支取方式，则进行“密码设置”",
    "提示客户分别输入“查询密码”和“取款密码”",
    "完成后点击“下一步”",
  ], { border: C.primary, fontSize: 10 });

  // Key fields
  infoCard(s, M + 4.5, 1.4, 4.5, 1.3, "关键字段", [
    { text: "查询密码 — 用于账户查询", options: { bold: true, color: C.primary } },
    { text: "取款密码 — 用于资金支取", options: { bold: true, color: C.accent } },
    "两种密码可相同也可不同",
    "设置时同步提示客户密码复杂程度强弱等级",
  ], { border: C.secondary, fontSize: 9 });

  // Notice box — full precautions
  noticeBox(s, M, 2.85, 5.2, 2.05, [
    "密码设置时必须输入两次，且两次输入必须一致",
    "两次密码输入不能超过30秒",
    "取款密码、交易密码必须为6位纯数字组成",
    "不能为空或连续数字简单密码（如123456、000000）",
    "设置密码时系统同步提示客户密码复杂程度强弱等级",
  ]);

  // Error box
  errorBox(s, M + 5.4, 2.85, 3.6, 2.05, [
    "两次输入不一致需重新设置，客户易产生焦躁情绪",
    "超时30秒未完成将导致设置失败",
    "简单密码被系统拒绝后需向客户解释原因",
    "查询密码与取款密码混淆是常见客诉点",
  ]);
  s.addNotes("【讲解节奏】密码设置是高频出错环节，建议用时3分钟，重点强调30秒限制和6位纯数字规则。\n【易混淆点】查询密码和取款密码可以相同也可以不同，但很多客户会混淆。建议口头提醒：\u201C查询密码用于查余额，取款密码用于取钱。\u201D\n【互动提问】如果客户两次输入不一致，系统会怎么处理？答：需要重新设置，不会锁定。");

  footer(s);
}

// ============================================================
// SLIDE 15: STEP 5 — Transaction Confirmation & Signature
// ============================================================
{
  let s = pres.addSlide();
  header(s, 15, "步骤5：交易确认与签名", "操作说明 · 关键字段 · 注意事项 · 易错点");

  stepBadge(s, M, 1.0, 5, "交易确认与电子签名");

  // Operation
  infoCard(s, M, 1.4, CW, 1.1, "操作说明", [
    "进入交易确认界面，查看输入项无误后点击提交",
    "提醒客户由柜外清进行交易确认并完成正楷签名（需柜员审核确认）",
    "无法进行签名时进行指纹录入",
    "查看交易结果。若为存折类凭证，根据系统提示进行写磁操作",
  ], { border: C.primary, fontSize: 9 });

  // Notice box
  noticeBox(s, M, 2.65, 5.2, 1.75, [
    "对于老年客群或无签字能力客群（手部残疾人员、手部受伤人员）进行指纹签名",
    "客户签名、指纹录入必须为本人或代理人",
    "签字必须使用正楷，书写工整、清晰可辨",
  ]);

  // Error box
  errorBox(s, M + 5.4, 2.65, 3.6, 1.75, [
    "签名不符合要求需重新签字，影响客户体验",
    "指纹信息不清晰需重新录入",
    "存折类凭证漏做写磁操作导致后续无法使用",
    "未柜员审核确认直接进入下一步将导致合规问题",
  ]);

  // Key fields
  infoCard(s, M, 4.5, CW, 0.6, "关键字段", [
    "交易确认界面：输入项核对 → 柜外清签名/指纹 → 交易结果 → 存折写磁（如适用）",
  ], { border: C.secondary, fontSize: 9 });
  s.addNotes("【讲解节奏】交易确认与签名建议用时2分钟。\n【易混淆点】签名必须正楷！草书签名会被柜员审核打回。老年客群或手部不便人员可用指纹替代。\n【关键提醒】存折类凭证别忘了写磁操作，漏做会导致存折后续无法使用，客户必须返回网点补做。");

  footer(s);
}

// ============================================================
// SLIDE 16: STEPS 6-7 — Voucher Printing & Evaluation
// ============================================================
{
  let s = pres.addSlide();
  header(s, 16, "步骤6-7：凭证打印与客户评价", "操作说明 · 关键字段 · 注意事项");

  // Step 6
  stepBadge(s, M, 1.0, 6, "凭证及回单打印");
  infoCard(s, M, 1.4, 4.3, 1.6, "操作说明", [
    "根据提示完成相应凭证及回单的打印",
    "若客户需要纸质回单，则进行纸质回单打印",
    "若客户不需要纸质回单，可通过手机银行扫描二维码查看业务办理信息",
  ], { border: C.primary, fontSize: 9 });

  infoCard(s, M + 4.5, 1.4, 4.5, 1.6, "关键字段", [
    { text: "纸质回单", options: { bold: true, color: C.primary } },
    "打印物理凭证交客户保管",
    { text: "电子回单", options: { bold: true, color: C.secondary } },
    "手机银行扫描二维码查看业务办理信息",
  ], { border: C.secondary, fontSize: 9 });

  // Step 7
  stepBadge(s, M, 3.15, 7, "客户评价");
  infoCard(s, M, 3.55, CW, 1.0, "操作说明与评价选项", [
    "由客户在柜外清对本次交易进行评价",
    "可选评价：非常满意 / 满意 / 不满意",
    "不评价默认为“客户未评价”",
  ], { border: C.primary, fontSize: 10 });

  // Note
  noticeBox(s, M, 4.7, CW, 0.45, [
    "业务办理结束后，系统对凭证进行质检并进行自动归档处理",
  ]);
  s.addNotes("【讲解节奏】步骤6-7合并讲解，建议用时2分钟。\n【易混淆点】电子回单不需要打印纸质凭证，客户可通过手机银行扫码查看。如果客户不习惯电子回单，仍可提供纸质回单。\n【注意事项】业务办理结束后系统自动质检归档，柜员无需手动操作。但如质检不通过，需按系统提示处理。");

  footer(s);
}

// ============================================================
// SLIDE 17: SPECIAL SCENARIOS & RED LINES
// ============================================================
{
  let s = pres.addSlide();
  header(s, 17, "特殊场景与红线清单", "集中展示所有特殊场景与注意事项");

  const scenarios = [
    {
      title: "代理办理", color: C.accent, icon: "⚠",
      items: [
        "需携带授权委托书 + 关系证明 + 特殊情况证明",
        "全过程双录并留存代理人的人脸识别信息",
        "人脸识别信息与业务办理信息关联",
      ],
    },
    {
      title: "未满16周岁", color: C.warning, icon: "⚠",
      items: [
        "以监护人代理办理处理",
        "需提供监护人身份及监护关系证明",
      ],
    },
    {
      title: "特殊客群上门服务", color: C.danger, icon: "⚠",
      items: [
        "重病/行动不便/无自理能力可上门服务",
        "或由配偶/父母/成年子女凭委托书+关系证明+特殊情况证明代理",
      ],
    },
    {
      title: "大额授权", color: C.primary, icon: "◆",
      items: [
        "存入资金超过50万触发集中作业授权",
        "代理办理也触发集中作业授权",
        "授权人员审核通过后方可继续",
      ],
    },
    {
      title: "老年/无签字能力客群", color: C.secondary, icon: "◆",
      items: [
        "手部残疾/手部受伤人员进行指纹签名",
        "签名必须正楷，书写工整清晰可辨",
        "电子签名需经柜员审核确认",
      ],
    },
    {
      title: "密码设置红线", color: C.dark, icon: "◆",
      items: [
        "6位纯数字，不能为空或连续简单密码",
        "两次输入一致，30秒内完成",
        "查询密码与取款密码分别设置",
      ],
    },
  ];

  const cardW = 2.85, cardH = 1.65, gapX = 0.22, gapY = 0.15;
  scenarios.forEach((sc, i) => {
    let col = i % 3, row = Math.floor(i / 3);
    let x = M + col * (cardW + gapX);
    let y = 1.05 + row * (cardH + gapY);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.white }, rectRadius: 0.05,
      line: { color: sc.color, width: 1.5 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: cardH, fill: { color: sc.color },
    });
    s.addText(`${sc.icon} ${sc.title}`, {
      x: x + 0.14, y: y + 0.06, w: cardW - 0.2, h: 0.24,
      fontSize: 10, fontFace: TF, color: sc.color, bold: true,
    });
    s.addText(sc.items.map(it => ({ text: it, options: { bullet: true, breakLine: true, paraSpaceAfter: 1 } })), {
      x: x + 0.14, y: y + 0.32, w: cardW - 0.24, h: cardH - 0.38,
      fontSize: 8, fontFace: BF, color: C.text, valign: "top", fit: "shrink",
    });
  });

  // Bottom red line banner
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.55, w: CW, h: 0.5,
    fill: { color: C.dark }, rectRadius: 0.04,
  });
  s.addText("红线提醒：代理办理必须双录 + 留存人脸识别  |  大额（>50万）必须授权  |  密码不得为简单连续数字", {
    x: M + 0.15, y: 4.55, w: CW - 0.3, h: 0.5,
    fontSize: 9, fontFace: TF, color: C.accent, bold: true,
    align: "center", valign: "middle",
  });
  s.addNotes("【红线强调】本页是全部红线清单的集中展示，建议用时4分钟逐条讲解。\n【处理话术】面对大额业务（>50万），提前告知客户：\u201C这笔业务需要授权人员审核，大约需要5-10分钟，请您稍候。\u201D避免客户焦躁。\n【真实案例】某网点代理办理未留存人脸识别信息，事后客户否认授权，因无双录记录无法举证，最终被通报批评。\n【互动提问】密码设置有哪些红线？答：6位纯数字、两次一致、30秒内、不能连续简单密码。");
  footer(s);
}

// ============================================================
// SLIDE 18: QUICK REFERENCE + CLOSING
// ============================================================
{
  let s = pres.addSlide();
  header(s, 18, "操作速查表", "7步操作快速参考");

  const rows = [
    [
      { text: "步骤", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
      { text: "操作内容", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
      { text: "关键要点", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
      { text: "注意事项", options: { bold: true, color: C.white, fill: { color: C.primary }, align: "center" } },
    ],
    ["1", "输入交易代码030601进入场景", "交易代码或场景名称", "—"],
    ["2", "选择证件类型+读取+人脸识别", "4种证件类型", "代理办理需双录；未满16岁监护人代理"],
    ["3", "选择6种介质之一并填写信息", "介质类型+字段清单", "储蓄+存单必须资金存入；录入冠字号码"],
    ["4", "设置查询密码和取款密码", "6位纯数字×2", "两次一致；30秒内；非连续简单密码"],
    ["5", "交易确认+签名/指纹", "柜外清正楷签名", "老年/无能力用指纹；存折需写磁"],
    ["6", "凭证及回单打印", "纸质/电子回单", "电子回单可手机银行扫码查看"],
    ["7", "客户评价", "4级评价", "不评价默认\u201C未评价\u201D；系统自动归档"],
  ];

  const colW = [0.55, 2.8, 2.4, 3.25];
  s.addTable(rows, {
    x: M, y: 1.1, w: CW, h: 3.0,
    colW,
    border: { type: "solid", pt: 1, color: C.textLight },
    rowH: 0.42,
    fontSize: 8, fontFace: BF, color: C.text,
    valign: "middle",
    autoPage: false,
  });

  // Closing banner — dark background with light text for proper contrast
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 4.3, w: CW, h: 0.75,
    fill: { color: C.dark }, rectRadius: 0.06,
  });
  s.addText("培训结束 · 请在实际操作中严格遵守注意事项与红线要求", {
    x: M + 0.2, y: 4.35, w: CW - 0.4, h: 0.35,
    fontSize: 13, fontFace: TF, color: C.white, bold: true,
    align: "center", valign: "middle",
  });
  s.addText("交易代码 030601  |  权限：业务柜员  |  综合个人开户 · 柜面操作培训", {
    x: M + 0.2, y: 4.68, w: CW - 0.4, h: 0.3,
    fontSize: 10, fontFace: BF, color: "B0D4E8",
    align: "center", valign: "middle",
  });
  s.addNotes("【复习提问】现在我们做几个自测题：1. 交易代码是多少？2. 4种证件类型有哪些？3. 6种介质分别是什么？4. 密码设置的3个红线是什么？5. 什么情况触发集中授权？\n【行动号召】培训结束后，请大家在实际操作中严格遵守注意事项与红线要求。遇到不确定的情况，及时请教老员工或主管。有问题随时联系培训组。");
}

// ============================================================
// SAVE
// ============================================================
pres.writeFile({ fileName: "综合个人开户_柜面操作培训.pptx" })
  .then(fn => console.log("Generated: " + fn))
  .catch(err => console.error("Error:", err));
