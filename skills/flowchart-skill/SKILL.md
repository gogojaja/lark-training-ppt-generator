# Skill: flowchart-skill

# 流程图快速生成技能包

> 版权：`../shared/references/COPYRIGHT.md`

## 1. 元数据

- **技能名称**：flowchart-skill
- **技能版本**：v1.0.0
- **发布日期**：2026-08-10
- **技能定位**：从 Word 文档提取业务流程，通过 CSV 节点表手工调整后，一键生成 PPT 流程图
- **适用场景**：业务流程图、操作手册流程图、制度流程图、培训流程图

## 2. 五步工作流

```
Word文档 → ①拆分 → ②选配色 → ③生成CSV → ④手工调整 → ⑤生成PPT
```

### Step 1：按章节拆分 Word 文档

**工具**：`tools/split_docx_by_level.py`（已有）

```bash
# 按一级标题拆分
py -3 tools/split_docx_by_level.py 输入文档/操作手册.docx 生成产物/拆分结果

# 按二级标题拆分
py -3 tools/split_docx_by_level.py 输入文档/操作手册.docx 生成产物/拆分结果 2
```

**最佳实践**：
- 优先按 Heading 1 拆分，每个章节独立处理
- 大文档（>50MB）先拆分再逐章提取流程
- 拆分后查看 `INDEX.md` 确认章节结构

### Step 2：确定配色方案

**预设配色**（`skills/flowchart-skill/presets/`）：

| 预设 | 主色 | 菱形 | 正常分支 | 异常分支 | 适用场景 |
|------|------|------|---------|---------|---------|
| `green` | 浅绿 C6EFCE | 浅黄 FFF2CC | 浅蓝 DDEBF7 | 浅红 FCE4EC | 合规/风控/审计 |
| `blue` | 浅蓝 D6E4F0 | 浅黄 FFF2CC | 浅绿 E2EFDA | 浅红 FCE4EC | 业务/运营/技术 |
| `red` | 浅红 FCE4EC | 浅黄 FFF2CC | 浅蓝 DDEBF7 | 淡红 F8D7DA | 应急/故障/升级 |
| `yellow` | 浅黄 FFF2CC | 浅蓝 DDEBF7 | 浅绿 E2EFDA | 浅红 FCE4EC | 培训/宣导/引导 |

**选择原则**：
- 主流程用柔和色（低饱和度），避免刺眼
- 菱形判断统一浅黄，保持一致性
- 异常分支用红色系，正常分支用蓝色/绿色系
- 标题栏用深色（深蓝/深绿），文字白色

### Step 3：生成流程节点 CSV 表

**方式 A · AI 语义化生成（推荐，主路径）**：由 AI（skill 执行者）读取 Word 内容，按下方 schema 语义化输出最终 CSV

```
流程：AI 读取业务文档（或拆分后的章节）→ 逐段提炼流程步骤 → 标注判断/分支 → 输出规范 CSV
工具：文档文本提取可复用 tools/split_docx_by_level.py（按章节拆分）或 docx_to_flow_csv.py --json-only（调试）
```

> **AI 生成要点（语义判断由模型负责，机械排版由脚本负责——遵循项目"模型无关/工具先于模型"原则）**：
> - 只输出 schema 所需的核心动作，**每节点 ≤15 字**（压缩而非截断）
> - 区分：**处理步骤**=rect 主流程；**判断/条件**=diamond（是否/若…则/校验/审核/异常 等）
> - 判断节点必须有分支：正常（`normal`）+ 异常（`error`），branch_to 指向 41+ 分支节点
> - 主流程 seq=1..N，分支 seq=41+；分支颜色：error 浅红 FCE4EC/C00000，normal 浅蓝 DDEBF7/1F3864
> - 单入口单出口；文本精炼为动作句（动词开头）
> 生成后填入小工具「CSV 节点表」直接生成 PPT，无需再次人工大改。

**方式 B · 手工整理**：按下方 schema 制表（AI 不可用时的替代）

**方式 C · 规则脚本离线草稿（备选，质量有限）**：纯规则启发式，仅作快速了解文档结构用
```bash
py -3 tools/docx_to_flow_csv.py 输入文档/操作手册.docx --out 生成产物/草稿.csv
```
> 由于 Word→CSV 本质是语义理解，规则脚本**无法产出可直接使用**的 CSV（截断/噪音/分支误判）。
> 仅建议用于初步摸清文档段落；**正式 CSV 请用方式 A（AI 生成）**。

**CSV 格式**（`skills/flowchart-skill/templates/flowchart_nodes.csv`）：

```csv
seq,node_type,content,shape,width_cm,height_cm,bg_color,text_color,branch_to,branch_label,branch_kind
1,main,客户到达网点取号,rect,5.0,0.6,C6EFCE,006100,,,
2,main,Pad身份识别与分流,rect,5.0,0.6,C6EFCE,006100,,,
3,main,选择证件类型并读取身份证件,diamond,4.5,1.0,FFF2CC,7F6000,41,缺失,error
41,branch,客户信息维护,rect,5.0,0.6,FCE4EC,C00000,,,
```

**字段说明**：

| 字段 | 必填 | 说明 |
|------|------|------|
| seq | Y | 序号（整数，主流程 1-N，分支 41+） |
| node_type | Y | `main`=主流程 / `branch`=分支节点 |
| content | Y | 节点文本内容 |
| shape | Y | `rect`=矩形 / `diamond`=菱形 / `circle`=圆形 / `round_rect`=圆角矩形 |
| width_cm | Y | 宽度（cm），矩形推荐 5.0，菱形推荐 4.5 |
| height_cm | Y | 高度（cm），矩形推荐 0.6，菱形推荐 1.0 |
| bg_color | Y | 背景色（6位 hex，不含 #） |
| text_color | Y | 字体颜色（6位 hex，不含 #） |
| branch_to | | 分支目标序号（主流程节点填写） |
| branch_label | | 分支标签（如"是"/"否"/"缺失"） |
| branch_kind | | `normal`=正常分支 / `error`=异常分支 |

**维度参考**（行业最佳实践）：

| 形状 | 推荐宽度 | 推荐高度 | 说明 |
|------|---------|---------|------|
| 矩形（rect） | 5.0cm | 0.6cm | 主流程步骤 |
| 菱形（diamond） | 4.5cm | 1.0cm | 判断/条件 |
| 圆形（circle） | 1.5cm | 1.5cm | 开始/结束 |
| 圆角矩形（round_rect） | 5.0cm | 0.6cm | 子流程 |

### Step 4：手工调整 CSV 表

**调整要点**：
1. **流程完整性**：检查是否有遗漏步骤或多余步骤
2. **分支逻辑**：确认每个菱形判断的"是/否"分支是否完整
3. **节点顺序**：seq 序号确保主流程连续（1,2,3...），分支用 41+ 编号
4. **文本精简**：每个节点控制在 15 字以内，避免框内文字溢出
5. **颜色一致性**：同类型节点使用相同配色

**常见调整**：
- 合并相似步骤（如"读取证件"+"校验证件"→"读取并校验证件"）
- 拆分复杂判断（一个菱形只对应一个二选一判断）
- 补充异常分支（每个菱形判断都应有异常路径）

### Step 5：生成 PPT 流程图

```bash
# 使用预设配色
py -3 生成脚本/csv_to_flowchart.py nodes.csv --preset green --out 流程图.pptx

# 自定义标题
py -3 生成脚本/csv_to_flowchart.py nodes.csv --preset blue --title "综合账户激活流程" --out 流程图.pptx

# 默认不生成连接线（如需连线显式加 --connectors）
py -3 生成脚本/csv_to_flowchart.py nodes.csv --preset green --connectors --out 流程图.pptx

# 仅输出 JSON（调试用）
py -3 生成脚本/csv_to_flowchart.py nodes.csv --json-only --out flow.json
```

## 3. 文件结构

```
skills/flowchart-skill/
├── SKILL.md                      # 本文件
├── presets/                      # 配色预设
│   ├── green.json                # 清新绿
│   ├── blue.json                 # 商务蓝
│   ├── red.json                  # 警示红
│   └── yellow.json               # 温暖黄
└── templates/
    ├── flowchart_nodes.csv       # CSV 节点表模板（精简版）
    └── flowchart_full_config.csv # CSV 全参数配置模板（推荐）

生成脚本/
├── csv_to_flowchart.py           # CSV → PPT 转换器（主入口）
├── gen_flowchart_branch.py       # PPT 流程图生成器（底层引擎）

tools/
└── docx_to_flow_csv.py           # Word→CSV 规则引擎（纯规则草稿，可选）

样例/
└── 综合账户激活模板/            # 标准模板样例（参考本目录)
    ├── README.md                 # 模板使用说明
    ├── 综合账户激活.docx         # 源业务文档
    ├── 综合账户激活_流程图.json  # 语义模式 JSON
    ├── 综合账户激活_流程图_语义模式.pptx  # 生成 PPT
    ├── flowchart_nodes.csv       # 精简版 CSV
    └── flowchart_full_config.csv # 全参数 CSV
```

## 4. 行业最佳实践

### 4.1 流程图设计规范

| 规范 | 说明 |
|------|------|
| 单一入口 | 流程图只有一个起点（开始节点） |
| 单一出口 | 流程图只有一个终点（结束节点），或明确的多出口标注 |
| 判断完整 | 每个菱形判断必须有"是"和"否"两个分支 |
| 从上到下 | 主流程从上到下流动，分支向右展开 |
| 文字精简 | 每个节点 15 字以内，避免长句 |
| 颜色一致 | 同类型节点使用相同配色，便于快速识别 |

### 4.2 常见节点类型

| 类型 | 形状 | 用途 | 颜色 |
|------|------|------|------|
| 开始 | 圆形 | 流程起点 | 深色 |
| 处理 | 矩形 | 操作步骤 | 浅绿/浅蓝 |
| 判断 | 菱形 | 条件分支 | 浅黄 |
| 正常分支 | 矩形 | 正常路径 | 浅蓝/浅绿 |
| 异常分支 | 矩形 | 异常路径 | 浅红 |
| 结束 | 圆形 | 流程终点 | 深色 |

### 4.3 质量检查清单

- [ ] 主流程步骤是否完整（无遗漏、无多余）
- [ ] 每个菱形判断是否有"是/否"两个分支
- [ ] 异常分支是否都有处理措施
- [ ] 节点文字是否精简（≤15字）
- [ ] 配色是否一致（同类节点同色）
- [ ] 连线是否清晰（无交叉、无遗漏）
- [ ] 整体布局是否从上到下、从左到右

## 5. 故障排除

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| 需要连接线 | 默认不绘制连线 | 使用 `--connectors` 或 CSV config `no_connectors=false` |
| 文字溢出框外 | 文本过长 | 精简节点文字至 15 字以内 |
| 颜色不一致 | CSV 颜色值错误 | 使用预设配色 `--preset green` |
| 菱形太小放不下文字 | 尺寸不合理 | 调整 width_cm=4.5, height_cm=1.0 |
| 分支节点位置偏移 | branch_to 序号错误 | 检查 seq 与 branch_to 对应关系 |

**文档版本**：v1.0.0　**最后更新**：2026-08-10
