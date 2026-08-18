# 综合账户激活 - 流程图模板样例

> 本目录作为后续同类业务的**标准模板**，后续所有该类业务（账户类、柜面类、渠道类流程）均参照本目录的文件结构与工作流。

## 目录文件说明

| 文件 | 说明 |
|------|------|
| `综合账户激活.docx` | 源业务文档（Word） |
| `综合账户激活_流程图.json` | 语义模式流程图 JSON（步骤+分支描述） |
| `综合账户激活_流程图_语义模式.pptx` | 生成的最终 PPT 流程图 |
| `flowchart_nodes.csv` | 精简版 CSV 节点表（无全局配置） |
| `flowchart_full_config.csv` | **全参数配置版** CSV 节点表（推荐使用） |
| `csv_to_flowchart.py` | CSV → PPT 一键转换脚本（主入口） |
| `gen_flowchart_branch.py` | PPT 流程图生成引擎（底层，一般无需改动） |

## 标准工作流

```
① 准备源文档 → ② 分析业务步骤 → ③ 填写 CSV 节点表 → ④ 生成 PPT → ⑤ 检查输出
```

### Step 1：准备源文档

将源 Word 文档（如`综合账户激活.docx`）放入 `输入文档/` 目录。

### Step 2：分析业务步骤

从文档中提取：
- **主流程步骤**：按顺序列出操作性步骤（矩形）
- **判断节点**：需要条件分支的步骤（菱形）
- **异常分支**：判断不通过时的处理路径（红色）
- **正常分支**：判断通过后的路径（蓝色/绿色）

### Step 3：填写 CSV 节点表

**推荐使用** `flowchart_full_config.csv`（全参数配置版）：

```csv
# 全局配置区
config,title,流程标题,流程图标题
config,preset,green,配色预设：green/blue/red/yellow
config,no_connectors,false,禁用连接线
config,step_gap_cm,1.2,纵向间隔
config,box_width_cm,5.0,矩形宽
config,box_height_cm,0.6,矩形高
config,diamond_width_cm,4.5,菱形宽
config,diamond_height_cm,1.0,菱形高

# 主流程节点区
seq,node_type,content,shape,width_cm,height_cm,bg_color,text_color,branch_to,branch_label,branch_kind
1,main,客户到达网点取号,rect,,,,,,,,
...

# 分支节点区
41,branch,客户信息维护,rect,,,,,,,,
```

### Step 4：生成 PPT

```bash
# 标准生成（自动读取 CSV 中全局配置）
py -3 生成脚本/csv_to_flowchart.py 样例/综合账户激活模板/flowchart_full_config.csv --out 生成产物/流程图.pptx

# 命令行覆盖（优先级高于 CSV 全局配置）
py -3 生成脚本/csv_to_flowchart.py nodes.csv --preset blue --title "新流程" --no-connectors --out 生成产物/新流程图.pptx
```

### Step 5：检查输出

用 PowerPoint / WPS 打开检查：
- [ ] 节点顺序与源文档一致
- [ ] 菱形判断分支完整（是/否）
- [ ] 颜色规范（主流程/菱形/分支/异常各有区别）
- [ ] 文字不溢出框外（节点 ≤15 字）
- [ ] 标题栏颜色与预设匹配

## CSV 全参数说明

### 全局配置（type=config）

| key | 说明 | 默认值 |
|-----|------|--------|
| title | 流程图标题 | 文件名 |
| preset | 配色预设（green/blue/red/yellow 或 JSON） | 无（内置绿） |
| no_connectors | 是否禁用连线（true/false） | false |
| step_gap_cm | 纵向间隔（cm） | 1.2 |
| box_width_cm | 主流程矩形宽（cm） | 5.0 |
| box_height_cm | 主流程矩形高（cm） | 0.6 |
| diamond_width_cm | 菱形宽（cm） | 4.5 |
| diamond_height_cm | 菱形高（cm） | 1.0 |
| title_bg | 标题栏背景色 | 1F3864 |
| title_text | 标题栏文字色 | FFFFFF |

### 节点数据（type=node / type=branch）

| 字段 | 说明 | 默认 |
|------|------|------|
| seq | 序号（主流程 1-N；分支 41+） | 必填 |
| node_type | main / branch | main |
| content | 节点文本 | 必填 |
| shape | rect/diamond/circle/round_rect | rect |
| width_cm | 节点宽（cm） | 全局值 |
| height_cm | 节点高（cm） | 全局值 |
| bg_color | 背景色（6位hex） | 预设色 |
| text_color | 文字色（6位hex） | 预设色 |
| branch_to | 分支目标 seq（仅主流程判断节点） | 空 |
| branch_label | 分支标签（"否"/"缺失"等） | 空 |
| branch_kind | normal/error | normal |

## 配色预设速查

| 预设 | 主流程 | 菱形 | 正常分支 | 异常分支 | 适用场景 |
|------|--------|------|---------|---------|---------|
| green | C6EFCE | FFF2CC | DDEBF7 | FCE4EC | 合规/风控/审计 |
| blue | D6E4F0 | FFF2CC | E2EFDA | FCE4EC | 业务/运营/技术 |
| red | FCE4EC | FFF2CC | DDEBF7 | F8D7DA | 应急/故障/升级 |
| yellow | FFF2CC | DDEBF7 | E2EFDA | FCE4EC | 培训/宣导/引导 |

## 扩展：自定义配色

创建自定义预设 JSON（如 `my_theme.json`）：

```json
{
  "main":    {"fill": "C6EFCE", "text": "006100"},
  "diamond": {"fill": "FFF2CC", "text": "7F6000"},
  "branch":  {"fill": "DDEBF7", "text": "1F3864"},
  "error":   {"fill": "FCE4EC", "text": "C00000"},
  "title_bg": "1F3864",
  "title_text": "FFFFFF"
}
```

使用时：`--preset my_theme.json`

---

**样例版本**：v1.0.0
**创建日期**：2026-08-10
**适用业务**：账户类/柜面类/渠道类业务流程