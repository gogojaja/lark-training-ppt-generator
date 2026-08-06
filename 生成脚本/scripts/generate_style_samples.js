/**
 * generate_style_samples.js — 风格样例预览脚本
 *
 * 参考 lark-workflow-handbook-deck/SKILL.md Step 2.5.3。
 *
 * 接收一份样例页内容（JSON 格式），生成 3 种风格变体的 PPTX 样例页，
 * 每页一种风格，填充同一份真实内容方便用户对比选择。
 *
 * 三种风格变体：
 *   A: Ocean Gradient  （蓝青系）三栏布局（上说明+中字段卡+下注意事项），底部通栏橙色高亮注意事项
 *   B: Charcoal Minimal（深灰系）双栏布局（左文右图），右侧固定侧栏注意事项
 *   C: Warm Terracotta  （暖色系）卡片矩阵布局，折叠卡片式注意事项
 *
 * 用法：
 *   node scripts/generate_style_samples.js --content sample_content.json --output style_samples.pptx
 *   node scripts/generate_style_samples.js --content sample_content.json --styles A,B
 *   node scripts/generate_style_samples.js   # 使用内置示例内容，输出 style_samples.pptx
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// ============================================================
// 页面尺寸（16:9，与 generate.js 一致）
// ============================================================
const W = 10;
const H = 5.625;
const M = 0.5;
const CW = W - 2 * M; // content width = 9

const TF = "Microsoft YaHei"; // title font
const BF = "Microsoft YaHei"; // body font

// ============================================================
// 三种风格变体的配色与元信息
// ============================================================
const VARIANTS = {
  A: {
    key: "A",
    name: "Ocean Gradient",
    cn: "蓝青系 · 清晰专业",
    layout: "三栏布局（上说明 + 中字段卡 + 下注意事项）",
    notice: "底部通栏橙色高亮",
    palette: {
      primary: "065A82",
      secondary: "1C7293",
      dark: "21295C",
      accent: "E86A33",
      bg: "F5F7FA",
      white: "FFFFFF",
      text: "1F2937",
      textLight: "6B7280",
      cardBg: "FFFFFF",
      noteBg: "FFF7ED",
      noteBorder: "E86A33",
      errorBg: "FEF2F2",
      errorBorder: "DC2626",
      danger: "DC2626",
    },
  },
  B: {
    key: "B",
    name: "Charcoal Minimal",
    cn: "深灰系 · 沉稳留白",
    layout: "双栏布局（左文右图）",
    notice: "右侧固定侧栏",
    palette: {
      primary: "333333",
      secondary: "555555",
      dark: "1F2937",
      accent: "0A8EA9",
      bg: "F5F5F5",
      white: "FFFFFF",
      text: "2D2D2D",
      textLight: "8A8A8A",
      cardBg: "FFFFFF",
      noteBg: "EAF6F9",
      noteBorder: "0A8EA9",
      errorBg: "FBF1F1",
      errorBorder: "B23A3A",
      danger: "B23A3A",
    },
  },
  C: {
    key: "C",
    name: "Warm Terracotta",
    cn: "暖色系 · 亲和全员",
    layout: "卡片矩阵布局",
    notice: "折叠卡片式",
    palette: {
      primary: "B85C38",
      secondary: "D08B5B",
      dark: "5C3D2E",
      accent: "5C3D2E",
      bg: "FDF6F0",
      white: "FFFFFF",
      text: "3D2B1F",
      textLight: "8C7263",
      cardBg: "FFFFFF",
      noteBg: "FBEFE6",
      noteBorder: "B85C38",
      errorBg: "FBEAE3",
      errorBorder: "8A3A1E",
      danger: "8A3A1E",
    },
  },
};

// ============================================================
// 内置示例内容（无 --content 时使用）
// 对应一个操作步骤页：步骤标题"步骤3：选择开户介质"
// ============================================================
const DEFAULT_CONTENT = {
  page_type: "step_operation",
  step_title: "步骤3：选择开户介质",
  step_number: 3,
  operation_desc:
    "客户经过开户协议阅读和确认后，在介质选择界面根据客户意愿选择需要开通的介质类型，再由客户勾选需要开的业务种类。柜面渠道支持全部 6 种介质：银行卡、存折、存单、大额存单、一本通、电子账户。每种介质完成后均进入步骤4（密码设置）。",
  key_fields: [
    {
      name: "银行卡开户",
      ops: ["选择\u201C银行卡开户\u201D", "选择卡种", "将新卡放置在 IC 卡读取区域进行读取"],
      fields: ["卡种", "（卡片信息自动读取）"],
    },
    {
      name: "存折开户",
      ops: ["选择\u201C存折开户\u201D", "根据客户开户意愿填写各项信息", "系统提示进行现金\u201C冠字号码\u201D录入"],
      fields: ["业务种类", "通兑标志", "支取方式", "开户金额", "凭证号码", "转存方式", "利率类型", "期限"],
    },
    {
      name: "存单开户",
      ops: ["选择\u201C存单开户\u201D", "根据客户开户意愿填写各项信息", "系统提示进行现金\u201C冠字号码\u201D录入"],
      fields: ["业务种类", "通兑标志", "支取方式", "开户金额", "凭证号码", "转存方式", "利率类型", "期限"],
    },
    {
      name: "大额存单开户",
      ops: ["选择\u201C大额存单\u201D开户", "勾选产品类型后点击\u201C下一步\u201D", "选择\u201C科目来源\u201D（现金/转账）", "选择\u201C存单形式\u201D（电子账户/纸质存单）"],
      fields: ["科目来源", "存单形式", "开户金额", "凭证号码"],
    },
    {
      name: "一本通开户",
      ops: ["选择\u201C一本通开户\u201D", "根据客户意愿填写开户各项信息", "填写完成后点击\u201C下一步\u201D"],
      fields: ["通兑标志", "支取方式", "凭证号码"],
    },
    {
      name: "电子账户开户",
      ops: ["选择\u201C电子账户开户\u201D", "根据客户意愿填写开户信息", "将 IC 卡放入读取区域点击\u201C读卡\u201D", "提醒客户阅读协议后签字确认"],
      fields: ["卡种", "产品代码"],
    },
  ],
  notices: [
    "若客户开立储蓄账户且介质为存单，则客户必须进行资金存入",
    "存折和存单开户必须录入现金\u201C冠字号码\u201D，遗漏将影响合规检查",
    "大额存单的\u201C科目来源\u201D选择不同，后续操作路径不同（现金直接输入 vs 转账需介质识别）",
  ],
  pitfalls: [
    "储蓄账户+存单组合容易遗漏资金存入步骤，导致开户失败",
    "存折/存单开户遗漏冠字号码录入将影响合规检查",
  ],
};

// ============================================================
// 参数解析
// ============================================================
function parseArgs(argv) {
  const args = { content: null, output: "style_samples.pptx", styles: "A,B,C" };
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--content") {
      args.content = argv[++i];
    } else if (a === "--output") {
      args.output = argv[++i];
    } else if (a === "--styles") {
      args.styles = argv[++i];
    } else if (a === "-h" || a === "--help") {
      console.log(
        "用法: node scripts/generate_style_samples.js --content <content.json> --output <out.pptx> --styles A,B,C"
      );
      process.exit(0);
    }
  }
  return args;
}

function loadContent(contentPath) {
  if (!contentPath) return DEFAULT_CONTENT;
  const abs = path.resolve(contentPath);
  if (!fs.existsSync(abs)) {
    console.warn(`[warn] 内容文件不存在: ${abs}，将使用内置示例内容。`);
    return DEFAULT_CONTENT;
  }
  try {
    return JSON.parse(fs.readFileSync(abs, "utf8"));
  } catch (e) {
    console.warn(`[warn] 解析内容文件失败: ${e.message}，将使用内置示例内容。`);
    return DEFAULT_CONTENT;
  }
}

// ============================================================
// 通用：变体标题栏（每页有标题栏显示变体名称）
// ============================================================
function variantHeader(slide, variant, content) {
  const P = variant.palette;
  slide.background = { color: P.bg };
  // 顶部主色条
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.07, fill: { color: P.primary },
  });
  // 变体名称标签（左上角色块）
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 0.18, w: 0.42, h: 0.42,
    fill: { color: P.primary }, rectRadius: 0.05,
  });
  slide.addText(variant.key, {
    x: M, y: 0.18, w: 0.42, h: 0.42,
    fontSize: 16, fontFace: TF, color: P.white, bold: true,
    align: "center", valign: "middle",
  });
  // 变体标题
  slide.addText(`风格变体 ${variant.key} — ${variant.name}`, {
    x: M + 0.52, y: 0.14, w: CW - 1.0, h: 0.34,
    fontSize: 18, fontFace: TF, color: P.dark, bold: true,
    valign: "middle", fit: "shrink",
  });
  // 副标题：布局 + 注意事项方式 + 步骤标题
  slide.addText(
    `${variant.cn}  |  ${variant.layout}  |  注意事项：${variant.notice}  |  样例：${content.step_title}`,
    {
      x: M + 0.52, y: 0.46, w: CW - 1.0, h: 0.24,
      fontSize: 9, fontFace: BF, color: P.textLight,
      valign: "middle", fit: "shrink",
    }
  );
  // 分隔线
  slide.addShape(pres.shapes.RECTANGLE, {
    x: M, y: 0.78, w: CW, h: 0.015, fill: { color: P.primary },
  });
}

// 通用：步骤标题行（步骤编号圆环 + 标题）
function stepTitleRow(slide, variant, content, y) {
  const P = variant.palette;
  slide.addShape(pres.shapes.OVAL, {
    x: M, y, w: 0.34, h: 0.34, fill: { color: P.primary },
  });
  slide.addText(String(content.step_number || "3"), {
    x: M, y, w: 0.34, h: 0.34,
    fontSize: 13, fontFace: TF, color: P.white, bold: true,
    align: "center", valign: "middle",
  });
  slide.addText(content.step_title, {
    x: M + 0.42, y: y - 0.03, w: CW - 0.42, h: 0.4,
    fontSize: 15, fontFace: TF, color: P.dark, bold: true,
    valign: "middle", fit: "shrink",
  });
}

// 通用：页脚
function footer(slide, variant) {
  const P = variant.palette;
  slide.addText("风格样例预览 · 制度手册转宣讲 PPT 工作流（Step 2.5.3）", {
    x: M, y: H - 0.28, w: CW, h: 0.2,
    fontSize: 7, fontFace: BF, color: P.textLight,
    align: "center",
  });
}

// 通用：易错点小条
function pitfallsStrip(slide, variant, items, y, h) {
  const P = variant.palette;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y, w: CW, h,
    fill: { color: P.errorBg }, rectRadius: 0.04,
    line: { color: P.errorBorder, width: 1 },
  });
  slide.addText("✕ 易错点提示", {
    x: M + 0.12, y: y + 0.03, w: 1.4, h: h - 0.06,
    fontSize: 9, fontFace: TF, color: P.danger, bold: true,
    valign: "middle",
  });
  slide.addText(items.map(t => ({ text: `• ${t}`, options: { breakLine: true } })), {
    x: M + 1.5, y: y + 0.03, w: CW - 1.62, h: h - 0.06,
    fontSize: 8, fontFace: BF, color: P.text,
    valign: "middle", fit: "shrink",
  });
}

// ============================================================
// 变体 A：Ocean Gradient — 三栏布局（上说明 + 中字段卡 + 下注意事项）
// 底部通栏橙色高亮注意事项
// ============================================================
function renderVariantA(content) {
  const v = VARIANTS.A;
  const P = v.palette;
  const s = pres.addSlide();
  variantHeader(s, v, content);

  // ① 上：操作说明（横通栏）
  const descY = 0.92;
  stepTitleRow(s, v, content, descY);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 1.34, w: CW, h: 0.62,
    fill: { color: P.cardBg }, rectRadius: 0.05,
    line: { color: P.secondary, width: 1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: M, y: 1.34, w: 0.06, h: 0.62, fill: { color: P.primary },
  });
  s.addText("操作说明", {
    x: M + 0.14, y: 1.37, w: 1.4, h: 0.2,
    fontSize: 9, fontFace: TF, color: P.primary, bold: true,
  });
  s.addText(content.operation_desc, {
    x: M + 0.14, y: 1.55, w: CW - 0.24, h: 0.38,
    fontSize: 8, fontFace: BF, color: P.text,
    valign: "top", fit: "shrink",
  });

  // ② 中：6 个介质字段卡（3×2 矩阵）
  const matrixY = 2.06;
  const matrixH = 1.62;
  const gap = 0.12;
  const cardW = (CW - 2 * gap) / 3;
  const cardH = (matrixH - gap) / 2;
  content.key_fields.forEach((kf, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = M + col * (cardW + gap);
    const y = matrixY + row * (cardH + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: P.cardBg }, rectRadius: 0.04,
      line: { color: P.secondary, width: 1 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: 0.04, fill: { color: P.primary },
    });
    s.addText(kf.name, {
      x: x + 0.08, y: y + 0.06, w: cardW - 0.16, h: 0.2,
      fontSize: 9, fontFace: TF, color: P.dark, bold: true,
      fit: "shrink",
    });
    s.addText(kf.fields.map(f => ({ text: f, options: { bullet: true, breakLine: true, paraSpaceAfter: 0 } })), {
      x: x + 0.08, y: y + 0.27, w: cardW - 0.16, h: cardH - 0.32,
      fontSize: 7, fontFace: BF, color: P.text,
      valign: "top", fit: "shrink",
    });
  });

  // ③ 下：底部通栏橙色高亮注意事项
  const noteY = 3.8;
  const noteH = 1.0;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: noteY, w: CW, h: noteH,
    fill: { color: P.noteBg }, rectRadius: 0.05,
    line: { color: P.noteBorder, width: 1.5 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: M, y: noteY, w: CW, h: 0.05, fill: { color: P.accent },
  });
  s.addText("⚠ 注意事项（底部通栏 · 橙色高亮）", {
    x: M + 0.14, y: noteY + 0.08, w: CW - 0.28, h: 0.22,
    fontSize: 10, fontFace: TF, color: P.accent, bold: true,
  });
  s.addText(
    content.notices.map((t, i) => ({
      text: `(${i + 1}) ${t}`,
      options: { breakLine: true, paraSpaceAfter: 2 },
    })),
    {
      x: M + 0.14, y: noteY + 0.3, w: CW - 0.28, h: noteH - 0.34,
      fontSize: 8, fontFace: BF, color: P.text,
      valign: "top", fit: "shrink",
    }
  );

  // 易错点
  pitfallsStrip(s, v, content.pitfalls, 4.88, 0.4);
  footer(s, v);
}

// ============================================================
// 变体 B：Charcoal Minimal — 双栏布局（左文右图），右侧固定侧栏注意事项
// ============================================================
function renderVariantB(content) {
  const v = VARIANTS.B;
  const P = v.palette;
  const s = pres.addSlide();
  variantHeader(s, v, content);

  stepTitleRow(s, v, content, 0.92);

  // 左栏：操作说明 + 6 介质字段文本
  const leftX = M;
  const leftW = 5.2;

  // 操作说明卡
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: leftX, y: 1.34, w: leftW, h: 0.92,
    fill: { color: P.cardBg }, rectRadius: 0.05,
    line: { color: P.secondary, width: 1 },
  });
  s.addText("操作说明", {
    x: leftX + 0.14, y: 1.38, w: leftW - 0.28, h: 0.22,
    fontSize: 9, fontFace: TF, color: P.primary, bold: true,
  });
  s.addText(content.operation_desc, {
    x: leftX + 0.14, y: 1.6, w: leftW - 0.28, h: 0.62,
    fontSize: 8, fontFace: BF, color: P.text,
    valign: "top", fit: "shrink",
  });

  // 6 介质字段列表卡
  const listY = 2.36;
  const listH = 2.5;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: leftX, y: listY, w: leftW, h: listH,
    fill: { color: P.cardBg }, rectRadius: 0.05,
    line: { color: P.secondary, width: 1 },
  });
  s.addText("关键字段（6 种介质）", {
    x: leftX + 0.14, y: listY + 0.05, w: leftW - 0.28, h: 0.22,
    fontSize: 9, fontFace: TF, color: P.primary, bold: true,
  });
  const fieldLines = content.key_fields.map(kf => ({
    text: `${kf.name}：${kf.fields.join(" / ")}`,
    options: { bullet: true, breakLine: true, paraSpaceAfter: 3 },
  }));
  s.addText(fieldLines, {
    x: leftX + 0.14, y: listY + 0.3, w: leftW - 0.28, h: listH - 0.36,
    fontSize: 8, fontFace: BF, color: P.text,
    valign: "top", fit: "shrink",
  });

  // 右栏：上方"右图"占位卡（用界面卡片框兜底视觉），下方固定侧栏注意事项
  const rightX = leftX + leftW + 0.25;
  const rightW = W - M - rightX; // 约 3.55

  // 右图占位（界面卡片框模拟）
  const imgY = 1.34;
  const imgH = 1.5;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: rightX, y: imgY, w: rightW, h: imgH,
    fill: { color: P.primary }, rectRadius: 0.05,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: rightX, y: imgY, w: rightW, h: 0.28,
    fill: { color: P.dark },
  });
  s.addText("介质选择界面（示意图）", {
    x: rightX + 0.12, y: imgY, w: rightW - 0.24, h: 0.28,
    fontSize: 8, fontFace: TF, color: P.white, bold: true,
    valign: "middle",
  });
  // 模拟 6 个介质按钮
  const btnGap = 0.08;
  const btnW = (rightW - 0.24 - btnGap) / 2;
  const btnH = 0.42;
  content.key_fields.slice(0, 6).forEach((kf, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const bx = rightX + 0.12 + col * (btnW + btnGap);
    const by = imgY + 0.36 + row * (btnH + 0.06);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: bx, y: by, w: btnW, h: btnH,
      fill: { color: P.white }, rectRadius: 0.04,
      line: { color: P.accent, width: 1 },
    });
    s.addText(kf.name, {
      x: bx + 0.04, y: by, w: btnW - 0.08, h: btnH,
      fontSize: 7, fontFace: BF, color: P.text, bold: true,
      align: "center", valign: "middle", fit: "shrink",
    });
  });

  // 右侧固定侧栏：注意事项
  const sideY = imgY + imgH + 0.12;
  const sideH = 4.9 - sideY; // 到 4.9
  s.addShape(pres.shapes.RECTANGLE, {
    x: rightX, y: sideY, w: 0.06, h: sideH, fill: { color: P.accent },
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: rightX + 0.06, y: sideY, w: rightW - 0.06, h: sideH,
    fill: { color: P.noteBg }, rectRadius: 0.04,
    line: { color: P.noteBorder, width: 1 },
  });
  s.addText("⚠ 注意事项", {
    x: rightX + 0.18, y: sideY + 0.06, w: rightW - 0.3, h: 0.22,
    fontSize: 9, fontFace: TF, color: P.accent, bold: true,
  });
  s.addText(
    content.notices.map((t, i) => ({
      text: `(${i + 1}) ${t}`,
      options: { breakLine: true, paraSpaceAfter: 3 },
    })),
    {
      x: rightX + 0.18, y: sideY + 0.3, w: rightW - 0.3, h: sideH - 0.36,
      fontSize: 7.5, fontFace: BF, color: P.text,
      valign: "top", fit: "shrink",
    }
  );

  // 易错点
  pitfallsStrip(s, v, content.pitfalls, 5.0, 0.32);
  footer(s, v);
}

// ============================================================
// 变体 C：Warm Terracotta — 卡片矩阵布局，折叠卡片式注意事项
// ============================================================
function renderVariantC(content) {
  const v = VARIANTS.C;
  const P = v.palette;
  const s = pres.addSlide();
  variantHeader(s, v, content);

  stepTitleRow(s, v, content, 0.92);

  // 操作说明通栏
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: M, y: 1.34, w: CW, h: 0.5,
    fill: { color: P.cardBg }, rectRadius: 0.05,
    line: { color: P.secondary, width: 1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: M, y: 1.34, w: 0.06, h: 0.5, fill: { color: P.primary },
  });
  s.addText(content.operation_desc, {
    x: M + 0.16, y: 1.36, w: CW - 0.28, h: 0.46,
    fontSize: 8, fontFace: BF, color: P.text,
    valign: "middle", fit: "shrink",
  });

  // 卡片矩阵：6 介质卡片 3×2
  const matrixY = 1.94;
  const matrixH = 1.5;
  const gap = 0.12;
  const cardW = (CW - 2 * gap) / 3;
  const cardH = (matrixH - gap) / 2;
  content.key_fields.forEach((kf, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = M + col * (cardW + gap);
    const y = matrixY + row * (cardH + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: P.cardBg }, rectRadius: 0.05,
      line: { color: P.secondary, width: 1 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: cardH, fill: { color: P.primary },
    });
    s.addText(kf.name, {
      x: x + 0.14, y: y + 0.05, w: cardW - 0.2, h: 0.2,
      fontSize: 9, fontFace: TF, color: P.primary, bold: true,
      fit: "shrink",
    });
    s.addText(kf.fields.map(f => ({ text: f, options: { bullet: true, breakLine: true, paraSpaceAfter: 0 } })), {
      x: x + 0.14, y: y + 0.26, w: cardW - 0.2, h: cardH - 0.3,
      fontSize: 7, fontFace: BF, color: P.text,
      valign: "top", fit: "shrink",
    });
  });

  // 折叠卡片式注意事项：3 张横向卡片，带"标题条 + 正文"的折叠卡外观
  const foldY = 3.56;
  const foldH = 1.18;
  const foldGap = 0.15;
  const foldW = (CW - 2 * foldGap) / 3;
  content.notices.forEach((t, i) => {
    const x = M + i * (foldW + foldGap);
    // 卡片底
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: foldY, w: foldW, h: foldH,
      fill: { color: P.noteBg }, rectRadius: 0.05,
      line: { color: P.noteBorder, width: 1 },
    });
    // 折叠卡标题条（深棕）
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: foldY, w: foldW, h: 0.26,
      fill: { color: P.dark }, rectRadius: 0.05,
    });
    // 盖住标题条下圆角，使只有上方圆角
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: foldY + 0.13, w: foldW, h: 0.13, fill: { color: P.dark },
    });
    s.addText(`▶ 注意事项 (${i + 1})`, {
      x: x + 0.1, y: foldY, w: foldW - 0.2, h: 0.26,
      fontSize: 8, fontFace: TF, color: P.white, bold: true,
      valign: "middle",
    });
    // 正文
    s.addText(t, {
      x: x + 0.1, y: foldY + 0.3, w: foldW - 0.2, h: foldH - 0.36,
      fontSize: 7.5, fontFace: BF, color: P.text,
      valign: "top", fit: "shrink",
    });
  });

  // 易错点
  pitfallsStrip(s, v, content.pitfalls, 4.86, 0.42);
  footer(s, v);
}

// ============================================================
// 主流程
// ============================================================
const args = parseArgs(process.argv);
const content = loadContent(args.content);

// 解析 --styles
const wanted = args.styles
  .split(",")
  .map(x => x.trim().toUpperCase())
  .filter(x => VARIANTS[x]);

if (wanted.length === 0) {
  console.error(`[error] 无效的 --styles 参数: ${args.styles}（可选值 A/B/C）`);
  process.exit(1);
}

const pres = new pptxgen();
pres.author = "lark-workflow-handbook-deck";
pres.title = "风格样例预览";
pres.subject = "Step 2.5.3 风格样例对比";
pres.layout = "LAYOUT_16x9";
pres.defineLayout({ name: "LAYOUT_16x9", width: W, height: H });
pres.layout = "LAYOUT_16x9";

wanted.forEach(key => {
  if (key === "A") renderVariantA(content);
  else if (key === "B") renderVariantB(content);
  else if (key === "C") renderVariantC(content);
});

const outPath = path.resolve(args.output);
pres
  .writeFile({ fileName: outPath })
  .then(fn => {
    console.log(`已生成 ${wanted.length} 个风格样例页 -> ${fn}`);
    console.log(`包含变体: ${wanted.join(", ")}`);
  })
  .catch(err => {
    console.error("生成失败:", err);
    process.exit(1);
  });
