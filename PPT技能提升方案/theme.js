/**
 * 浅绿色主题设计令牌
 * 风格定位：清新、柔和、护眼 —— 适合培训手册、产品介绍、政务服务类PPT
 * 主风格：正式工作汇报（浅绿变体）
 * 辅风格：清新活力
 * 设计张力：柔和底色 + 中饱和度彩色卡片
 * 
 * 使用方式：
 *   const { color, typography, size, spacing, roleColors } = require('./theme.js');
 */

// ============================================================
// 一、颜色令牌（按角色命名，不按页面命名）
// ============================================================
const color = {
  // ---- 基础中性色 ----
  bg:           "F5FFF7",  // 页面背景：极浅绿（暖白偏绿，护眼）
  surface:      "FFFFFF",  // 卡片/面板底色：纯白
  hairline:     "D5E8D4",  // 细分割线：浅绿灰
  hairlineAlt:  "E8F5E9",  // 更细的分割线：接近白色

  // ---- 文字色（四级可读层级）----
  ink:          "1B3A2B",  // 主文字：深墨绿（近黑，阅读舒适）
  text:         "2C3E2F",  // 正文：深灰绿（默认正文字色）
  muted:        "5D7A66",  // 次级文字：中灰绿（注释、辅助说明）
  subtle:       "85998A",  // 弱化文字：浅灰绿（来源、页脚、装饰）

  // ---- 品牌主色 ----
  primary:      "27AE60",  // 主色：柔和中绿（标题、按钮、表头）
  primaryDark:  "1E8449",  // 主色加深（悬停、强调）
  primaryLight: "82E0AA",  // 主色变浅（装饰线、淡背景）
  primaryPale:  "D5F5E3",  // 主色极浅（大背景块、高亮底）

  // ---- 辅助色 ----
  accent:       "F4D03F",  // 强调色：暖黄（重点提示、装饰线、图标）
  accentDark:   "F39C12",  // 强调加深（警告级提示）
  secondary:    "5DADE2",  // 辅色：天蓝（信息卡片、第二优先级）
  secondaryLit: "AED6F1",  // 辅色变浅
  tertiary:     "AF7AC5",  // 第三色：淡紫（点缀、分类标识）

  // ---- 语义状态色 ----
  positive:     "58D68D",  // 正向/达标/成功：浅绿
  success:      "27AE60",  // 成功（同主色，语义别名）
  caution:      "F5B041",  // 警告/偏差：暖橙
  warning:      "F39C12",  // 强警告
  risk:         "EC7063",  // 风险/未达标：柔和红
  danger:       "E74C3C",  // 危险/错误
  info:         "5DADE2",  // 信息提示：浅蓝

  // ---- 深色块（封面/结束页）----
  dark:         "1E5631",  // 深绿（深色页面背景）
  darkInk:      "FFFFFF",  // 深色页上的文字
  darkMuted:    "A9DFBF",  // 深色页上的次级文字
  darkSubtle:   "82E0AA",  // 深色页上的弱化文字

  // ---- 专用背景（提示框）----
  noteBg:       "FEF9E7",  // 注意事项底：浅黄
  noteBorder:   "F4D03F",  // 注意事项边框：暖黄
  infoBg:       "EBF8FB",  // 信息提示底：浅蓝
  infoBorder:   "5DADE2",  // 信息提示边框
  successBg:    "EAFAF1",  // 成功提示底：浅绿
  successBorder:"27AE60",  // 成功提示边框
  errorBg:      "FDF2F0",  // 错误提示底：浅红
  errorBorder:  "EC7063",  // 错误提示边框
};

// ============================================================
// 二、字体令牌（按角色命名）
// ============================================================
const typography = {
  // ---- 字体族 ----
  fontTitle:    "Microsoft YaHei",   // 标题字体
  fontBody:     "Microsoft YaHei",   // 正文字体
  fontMono:     "Consolas",          // 等宽字体（代码、编号）

  // ---- 字号角色（16:9 幻灯片，单位 pt）----
  deckTitle:    42,     // 封面主标题
  deckSubtitle: 22,     // 封面副标题
  claim:        20,     // 页面标题（主张）
  sectionLabel: 9,      // 分区标签/章节导航
  body:         9,      // 正文（默认）
  bodySmall:    8,      // 小号正文（表格、密集卡片）
  annotation:   7.5,    // 注释/卡片描述
  source:       7,      // 来源/页脚
  pageNum:      13,     // 页码徽章

  // ---- 字重 ----
  weightBold:   true,
  weightNormal: false,
};

// ============================================================
// 三、尺寸令牌
// ============================================================
const size = {
  // ---- 画布（16:9，单位英寸）----
  slideW:       10,
  slideH:       5.625,

  // ---- 外边距 ----
  margin:       0.5,     // 左右外边距
  contentW:     9.0,     // 内容区宽度 = slideW - 2*margin

  // ---- 垂直节奏（六段式位置，单位英寸，从顶部算）----
  contextBandY:  0.92,   // 上下文带（分区标签）y 坐标
  claimBandY:    0.15,   // 主张带（标题）y 坐标
  claimBandH:    0.42,   // 主张带高度
  separationH:   0.38,   // 分离带高度（标题到内容的间距）
  evidenceY:     1.05,   // 证据区起始 y
  footerY:       5.325,  // 页脚 y
  footerH:       0.22,   // 页脚高度

  // ---- 页眉元素 ----
  headerBarH:    0.06,   // 顶部装饰线高度
  pageBadgeW:    0.42,   // 页码徽章宽度
  pageBadgeH:    0.38,   // 页码徽章高度
  headerDividerH:0.015,  // 标题下方分割线高度

  // ---- 卡片 ----
  cardRadius:    0.05,   // 卡片圆角半径
  cardBorderW:   1,      // 卡片边框线宽（磅）
  cardAccentW:   0.05,   // 卡片左侧装饰条宽度
  cardInnerPad:  0.14,   // 卡片内边距

  // ---- 表格 ----
  tableRowH:     0.4,    // 表格行高（默认）
  tableRowHDense:0.3,    // 紧凑表格行高
  tableBorderW:  1,      // 表格边框线宽
  tableBorderColor: "BDC3C7",  // 表格边框色（灰）
};

// ============================================================
// 四、间距令牌（三级尺度）
// ============================================================
const spacing = {
  // ---- 宏观（Macro）：页面级间距 ----
  sectionGap:    0.3,    // 大区块之间的间距
  cardGapX:      0.3,    // 大卡片列间距
  cardGapY:      0.15,   // 大卡片行间距

  // ---- 中观（Meso）：组件级间距 ----
  colGap:        0.2,    // 列间距
  rowGap:        0.12,   // 行间距
  tileGap:       0.08,   // 小卡片/瓦片间距

  // ---- 微观（Micro）：文本级间距 ----
  paraSpaceAfter: 1,     // 段落间距（磅）
  bulletIndent:   0.1,   // 列表缩进
  labelValueGap:  0.05,  // 标签与值的间距
};

// ============================================================
// 五、角色色板（泳道图/分类用）
// ============================================================
const roleColors = {
  // ---- 五角色标准泳道 ----
  customer:   color.positive,   // 客户：浅绿
  hall:       color.accent,     // 厅堂人员：暖黄
  teller:     color.secondary,  // 柜员：浅蓝
  system:     color.secondaryLit,// 系统：更浅蓝
  auth:       color.risk,       // 授权人员：浅红
  manager:    color.tertiary,   // 管理人员：浅紫

  // ---- 业务分类色 ----
  finance:    color.primary,    // 财务/现金：绿
  service:    color.secondary,  // 服务：蓝
  product:    color.tertiary,   // 产品：紫
  risk:       color.danger,     // 风险：红
  operation:  color.accent,     // 运营：黄
};

// ============================================================
// 六、页面轮廓模板（快速引用）
// ============================================================
const layoutTemplates = {
  // 封面布局
  cover: {
    titleY:   1.6,
    titleH:   0.9,
    subtitleY:2.55,
    subtitleH:0.5,
    accentY:  3.2,
    accentW:  2.5,
    descY:    3.4,
    descH:    0.35,
    metaY:    4.3,
    metaH:    0.3,
    leftBarW: 0.08,
  },

  // 标准内容页（页眉+内容+页脚）
  content: {
    headerY:    0,
    headerH:    0.82,
    contentY:   1.05,
    contentH:   4.0,
    footerY:    size.footerY,
  },

  // 卡片网格：2列 × 2行
  grid2x2: {
    cardW:    4.35,
    cardH:    1.7,
    gapX:     spacing.cardGapX,
    gapY:     spacing.cardGapY,
    startX:   size.margin,
    startY:   size.evidenceY,
  },

  // 卡片网格：3列 × 2行
  grid3x2: {
    cardW:    2.85,
    cardH:    1.6,
    gapX:     spacing.colGap,
    gapY:     spacing.cardGapY,
    startX:   size.margin,
    startY:   size.evidenceY,
  },

  // 时间线/步骤流：7 步横向
  timeline7: {
    cardW:    1.2,
    cardH:    3.5,
    gap:      0.07,
    startY:   size.evidenceY + 0.05,
  },

  // 泳道图：5 行
  swimlane5: {
    labelW:   0.95,
    laneH:    0.62,
    startY:   size.evidenceY,
    nodeW:    1.05,
    nodeH:    0.45,
  },
};

// ============================================================
// 七、组件默认配置
// ============================================================
const components = {
  // 信息卡片
  infoCard: {
    bg:         color.surface,
    border:     color.secondary,
    radius:     size.cardRadius,
    borderW:    size.cardBorderW,
    accentW:    size.cardAccentW,
    titleSize:  11,
    titleColor: color.secondary,
    bodySize:   typography.body,
    bodyColor:  color.text,
    innerPad:   size.cardInnerPad,
  },

  // 注意事项框
  noticeBox: {
    bg:         color.noteBg,
    border:     color.noteBorder,
    radius:     0.06,
    title:      "⚠ 注意事项",
    titleSize:  10,
    titleColor: color.accentDark,
    bodySize:   8,
    bodyColor:  color.text,
  },

  // 错误/易错点框
  errorBox: {
    bg:         color.errorBg,
    border:     color.errorBorder,
    radius:     0.06,
    title:      "✕ 易错点提示",
    titleSize:  10,
    titleColor: color.danger,
    bodySize:   8,
    bodyColor:  color.text,
  },

  // 步骤徽章
  stepBadge: {
    size:       0.32,
    bg:         color.primary,
    textColor:  color.darkInk,
    fontSize:   12,
    labelSize:  12,
    labelColor: color.ink,
  },

  // 表格表头
  tableHeader: {
    bg:         color.primary,
    textColor:  color.darkInk,
    fontSize:   typography.body,
    bold:       true,
    align:      "center",
  },

  // 深色提示条（底部强调）
  darkBanner: {
    bg:         color.dark,
    textColor:  color.darkInk,
    accentColor:color.accent,
    radius:     0.04,
    fontSize:   9,
  },
};

// ============================================================
// 八、风格简报（Style Brief）
// ============================================================
const styleBrief = {
  name: "mint-fresh",
  adjectives: ["清新", "柔和", "专业"],
  tension: "柔和底色 + 中饱和彩色卡片（轻而不飘）",
  density: "中高密度，靠分组和对齐增加密度，不缩小字号",
  colorLogic: "浅绿主色 + 蓝黄红紫四色辅色 + 深绿深色块",
  fontLogic: "微软雅黑统一字体，字重对比作为主要区分手段",
  imageBehavior: "以图形/图标为主，少用照片；如需图片则加绿色色罩",
  rhythm: "六段式垂直节奏，页眉一致，页脚统一，内容区变化轮廓",
  useCases: ["培训手册", "产品介绍", "政务服务", "教育课件", "健康医疗"],
};

// ============================================================
// 导出
// ============================================================
module.exports = {
  color,
  typography,
  size,
  spacing,
  roleColors,
  layoutTemplates,
  components,
  styleBrief,
};
