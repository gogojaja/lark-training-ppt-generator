# 生成脚本运行环境说明

「制度手册转宣讲 PPT 工作流」的本地生成脚本基于 **Node.js + pptxgenjs** 实现，可离线生成 PPTX 样例（不依赖 lark-cli）。

> **零依赖流程图工具**：`gen_flowchart.py` 用 Python 标准库直接读写 PPTX/docx（zip+xml），**无需 node/pip/第三方库**，本机 `py -3` 即可运行。

## 环境要求

| 组件 | 说明 |
|------|------|
| Node.js | LTS 版（未安装时运行 `setup_environment.bat` 会提示） |
| pptxgenjs | 由 `setup_environment.bat` 自动 `npm install` |

## 快速开始

```bat
:: 1. 一键准备环境（检查 node + 安装 pptxgenjs）
setup_environment.bat

:: 2. 生成主 PPT（读取 slide_plan.json 对应的 18 页版式）
node generate.js

:: 3. 生成风格样例（Step 3 风格预览用）
node scripts/generate_style_samples.js
```

## Word → 一页流程图 PPT（零依赖）

```bat
:: 方式一：读 Word 文档（第一段作标题，其余段落按顺序作为流程步骤）
py -3 gen_flowchart.py 流程操作.docx --out 流程图.pptx

:: 方式二：命令行直接给步骤（分号或换行分隔）
py -3 gen_flowchart.py --steps "预约叫号;身份识别;资料录入;业务开通;授权复核;归档结束" --title 柜面开户操作流程 --out 流程图.pptx
```

- 输出为**一页 16:9 纵向流程图**：开始（红）→ 步骤（蓝）→ 结束（绿），方框 + 下箭头。
- 不依赖 node/python-pptx/pip，仅用 Python 标准库（zipfile+ElementTree 手写 PPTX XML）。

## 分支流程图（gen_flowchart_branch.py，语义模式）

```bat
:: 语义模式：JSON 只写步骤与分支，脚本自动排版 + 配色，无需手算坐标
py -3 gen_flowchart_branch.py flow.json --out 流程图.pptx
```

`flow.json` 结构：

```json
{
  "title": "个人批量开户业务办理流程（纵向）",
  "steps": [
    {"text": "登录系统进入场景"},
    {"text": "客户信息是否齐全",
     "branch": {"text": "跳转客户信息维护", "label": "否", "kind": "err"}}
  ]
}
```

- 含 `branch` 的步骤自动变为**菱形判断**，右侧生成**分支框**（肘形连线 + 是/否标签）。
- `branch.kind`: `err` = 异常分支（浅红）、`br` = 正常分支（浅蓝）。

**固化配色规范**（脚本常量，可在高级模式 JSON 中覆盖）：

| 节点 | 填充色 | 文字色 |
|------|--------|--------|
| 主流程框 | 浅绿 `C6EFCE` | 深绿 `006100` |
| 菱形判断 | 浅黄 `FFF2CC` | 深黄 `7F6000` |
| 正常分支 | 浅蓝 `DDEBF7` | 深蓝 `1F3864` |
| 异常分支 | 浅红 `FCE4EC` | 红 `C00000` |

**固化纵向布局参数**（`gen_flowchart_branch.py` 顶部常量）：主流程列 X=900000、框宽 2600000、框高 216000（0.6cm）、行距 275000、分支宽 2700000、分支间距 300000。

**高级模式**（完全控制坐标/连线/颜色）：JSON 含 `nodes` + `edges` 字段，节点用 `kind: box/short/diamond`，连线用 `style: v/h/el/el-right/el-left`，并支持 `label` 标注。

> **设计目标（模型无关）**：排版坐标、配色、连线逻辑全部固化在脚本常量中，**模型只负责提供步骤语义内容**，不参与任何像素/坐标/色值计算；切换模型后同一 JSON 输出完全一致。

## 输出

- `generate.js` → `综合个人开户_柜面操作培训.pptx`
- `generate_style_samples.js` → `style_samples.pptx`（风格变体预览）
- `gen_flowchart.py` → 指定 `--out` 的一页流程图 PPTX
- `gen_flowchart_branch.py` → 指定 `--out` 的带分支流程图 PPTX（语义/高级模式）

## 与本机现状

> **注意**：当前执行机未安装 node。可直接运行 `setup_environment.bat` 获得安装指引；
> 也可用 `py -3`（已可用）运行 `D:/trae/lark-slides/scripts/xml_text_overlap_check.py` 做回读验证。

## 本地优先与可选云端协同

主流程以**本地离线生成**为主，不依赖 `lark-cli` 和飞书云端环境即可完成 PPTX 产出；飞书相关能力仅作为可选扩展，用于云端同步、发布和协同。

- 默认模式：本地生成、校验、输出
- 可选模式：Lark 文档读取 / 统一导出 / 发布共享
- 失败场景：若未配置飞书环境，主流程不应被阻断

---

**文档版本**：v1.2.0　**最后更新**：2026-08-17