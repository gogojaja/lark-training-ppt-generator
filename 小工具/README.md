# PPT流程图工具 · 使用说明

> 本工具将「Word 文档拆分」与「PPT 流程图生成」封装为本地可视化软件。
> **完全离线、无需联网、无需命令行**，双击即可使用。

---

## 一、快速开始

### 方式1 · 双击启动（推荐）
双击 `小工具/启动_PPT流程图工具.bat`，即可打开图形界面。

### 方式2 · 命令行启动
```bash
py -3 小工具/PPT流程图工具.py
```

> **前提**：已安装 Python 3（无需任何第三方库，全部使用标准库）。
> 检查方法：命令行运行 `py --version`，有输出即已安装。

> **说明**：`.bat` 启动器（ASCII脚本）→ `flowchart_tool_launcher.pyw`（ASCII启动器）→ `PPT流程图工具.py`（中文主程序）。
> 采用 ASCII 启动器是为了避免中文路径/代码页（codepage）导致 `.bat` 双击失效。

---

## 二、功能①：Word 文档拆分

**用途**：将大型 Word 文档按章节大纲级别拆分，保留图片，输出多个独立 docx。

### 操作步骤
1. **① 选择源文档**：点击「浏览…」选择 `.docx` 文件
2. **拆分设置**：选择拆分级别
   - `1` = 按一级标题拆分（推荐，每章一个文件）
   - `2` = 按二级标题拆分
   - `3` = 按三级标题拆分
3. **输出目录**：选择拆分结果存放位置（留空则生成在源文档旁）
4. 点击「开始拆分」→ 查看执行日志

### 示例
```
输入：操作手册.docx
级别：1
输出：D:\拆分结果\
      ├── 01_个人开户.docx
      ├── 02_账户激活.docx
      └── ...
```

---

## 三、功能②：PPT 流程图生成

**用途**：从业务文档生成流程节点 CSV，再从 CSV 节点表生成业务流程图 PPT。

### 工作流总览

```
Word文档 → ①拆分章节 → ②AI生成CSV → ③(可选)手工调整 → ④生成PPT
                      ↗
            规则预览（辅助摸底）
```

### 步骤 0：准备 CSV 节点表

#### 方式 A · AI 语义化生成（推荐，主路径）

调用 **flowchart-skill**，由 AI 读取 Word 文档内容并语义化输出标准 CSV 节点表。遵循「**模型无关 / 工具先于模型**」原则：语义判断由模型负责，机械排版/配色/坐标由脚本固化。

**四步工作流**：
1. **文档预处理**：大文档先拆分为独立章节，每章单独生成流程图
2. **AI 语义提炼**：按规范 Prompt 提取流程节点（详见 flowchart-skill Step 3）
3. **质量校验**：按 7 项自检清单核对（完整性/分支/异常/文字/编号/配色/结构）
4. **输出 CSV**：全参数 CSV 格式（含 config 全局配置区）

> 详细规范见 `skills/flowchart-skill/SKILL.md` Step 3。

#### 方式 B · 规则预览（辅助摸底，质量有限）

在「〇、CSV 来源」区选择 `.docx` 文档，点击「**规则预览**」可快速查看离线规则引擎输出。

**规则预览特性**（v1.1.0 新增）：
- ✅ 自动加载：预览后**自动填入** CSV 节点表路径，可直接生成或编辑
- ✅ 自动标题：从文件名推导流程图标题并**自动填入**
- ✅ 全参数格式：输出为带 config 全局配置区的完整 CSV（含 4 套预设配色）
- ✅ 详细统计：主流程节点数 / 判断节点数 / 分支节点数分类统计
- ⚠️ 仅作参考：规则启发式质量有限，**正式 CSV 建议由 AI 生成**

输出位置：与 Word 同目录的 `源文档名_规则预览.csv`

#### 方式 C · 手工整理

使用 CSV 模板手工填写，或从规则预览的草稿基础上修改。
模板位置：`skills/flowchart-skill/templates/flowchart_full_config.csv`

### 步骤 1：生成流程图

1. **CSV 节点表**：点击「浏览…」选择 `.csv` 文件
   - 可使用 `skills/flowchart-skill/templates/` 下的模板，或点「查看CSV模板」按钮
2. **配色方案**：green（绿）/ blue（蓝）/ red（红）/ yellow（黄）
3. **标题**：流程图标题（留空则用文件名）
4. **连接线设置**：默认不连线，勾选「生成连接线」则绘制
5. 点击「生成流程图」→ 查看日志

### CSV 格式说明

**全参数 CSV 格式**（推荐，含全局配置区）：

```csv
# === [全局配置区] ===
config,title,综合账户激活业务流程,流程图标题
config,preset,green,配色预设
config,no_connectors,true,禁用连接线（默认true）
config,step_gap_cm,1.2,纵向间隔（cm）

# === [主流程节点区] ===
seq,node_type,content,shape,width_cm,height_cm,bg_color,text_color,branch_to,branch_label,branch_kind
1,main,客户到达网点取号,rect,,,,,,,,
2,main,人脸识别身份核实,diamond,,,,41,不通过,error

# === [分支节点区] ===
41,branch,上级现场审核,rect,,,,,,,,
```

**字段说明**：

| 字段 | 必填 | 说明 |
|------|------|------|
| seq | Y | 序号（主流程1-N，分支41+） |
| node_type | Y | `main`=主流程 / `branch`=分支节点 |
| content | Y | 节点文字（≤15字最佳） |
| shape | Y | `rect`矩形 / `diamond`菱形 / `circle`圆形 / `round_rect`圆角矩形 |
| width_cm | | 宽度（cm），留空用全局默认 |
| height_cm | | 高度（cm），留空用全局默认 |
| bg_color | | 背景色（6位hex），留空用预设色 |
| text_color | | 文字色（6位hex），留空用预设色 |
| branch_to | | 分支目标序号（仅判断节点填） |
| branch_label | | 分支标签（如"否"/"缺失"） |
| branch_kind | | `normal`正常 / `error`异常 |

> 颜色/尺寸留空则自动套用配色预设，也可手工填写 6 位色号覆盖。

---

## 四、命令行用法（进阶）

### 规则预览命令行

```bash
# 生成规则预览 CSV（全参数格式）
py -3 tools/docx_to_flow_csv.py 输入文档.docx --out 草稿.csv --preset green

# 自定义标题
py -3 tools/docx_to_flow_csv.py 输入文档.docx --title "个人开户流程"

# 仅查看文档统计（不生成CSV）
py -3 tools/docx_to_flow_csv.py 输入文档.docx --stats

# 仅输出 JSON（调试用）
py -3 tools/docx_to_flow_csv.py 输入文档.docx --json-only --out draft.json
```

### 流程图生成命令行

```bash
# 使用预设配色
py -3 生成脚本/csv_to_flowchart.py nodes.csv --preset green --out 流程图.pptx

# 自定义标题
py -3 生成脚本/csv_to_flowchart.py nodes.csv --title "综合账户激活流程"

# 生成连接线（默认不连线）
py -3 生成脚本/csv_to_flowchart.py nodes.csv --connectors

# 仅输出 JSON（调试用）
py -3 生成脚本/csv_to_flowchart.py nodes.csv --json-only --out flow.json
```

---

## 五、目录结构

```
lark-training-ppt-generator/
├── 小工具/                  ← 本工具
│   ├── PPT流程图工具.py    ← 主程序（图形界面）
│   ├── flowchart_tool_launcher.pyw ← ASCII启动器（供.bat调用）
│   ├── 启动_PPT流程图工具.bat  ← 双击启动器
│   └── README.md           ← 本说明
├── tools/
│   ├── split_docx_by_level.py   ← 文档拆分引擎
│   └── docx_to_flow_csv.py      ← Word→CSV 规则引擎（纯规则草稿）
├── 生成脚本/
│   ├── csv_to_flowchart.py      ← CSV→PPT 转换
│   └── gen_flowchart_branch.py  ← PPT 流程图生成引擎
└── skills/flowchart-skill/
    ├── SKILL.md                 ← 流程图技能（含AI生成工作流）
    ├── presets/                 ← 配色方案（green/blue/red/yellow）
    └── templates/               ← CSV 模板（精简版 + 全参数版）
```

---

## 六、常见问题

| 问题 | 解决 |
|------|------|
| 双击 `.bat` 无窗口 | 未安装 Python，或未勾选「加入PATH」。先安装 Python 3 |
| 无法选择文件 | 点击右侧「浏览…」，确保弹窗中选中文件 |
| 生成失败 | 查看「执行日志」，通常为 CSV 格式错误 |
| 需要连线 | 流程图标签页勾选「生成连接线」 |
| 规则预览 CSV 能用吗？ | 仅作参考摸底，正式 CSV 建议由 AI 生成（flowchart-skill 方式 A） |
| 支持哪些配色？ | 4 套预设：green（合规绿）/ blue（商务蓝）/ red（警示红）/ yellow（培训黄） |

---

## 七、更新日志

### v1.1.0（2026-08-11）
- **规则预览增强**：
  - 预览后自动加载 CSV 到节点表，自动设置标题
  - 输出格式升级为全参数 CSV（含 config 全局配置区）
  - 新增 4 套预设配色适配（green/blue/red/yellow）
  - 增强输出统计（主流程/判断/分支分类统计）
  - 新增 `--stats` 参数（文档统计 + 大纲结构预览）
  - 新增 `--title` 参数（自定义流程图标题）
- **文档联动**：与 flowchart-skill v1.1.0 AI 语义提取工作流对齐
- **README 更新**：补充全参数 CSV 格式说明、命令行用法、更新日志

### v1.0.0（2026-08-10）
- 首版发布：双标签页 GUI（Word 文档拆分 + PPT 流程图生成）
- 离线零依赖：仅 Python 标准库，无需联网
- ASCII 启动器：规避中文路径/代码页问题
- 4 套配色预设 + CSV 模板
- 连接线默认不绘制，可选开启

---

**版本**：v1.1.0
**日期**：2026-08-11
