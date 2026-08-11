# 技能包索引 (SKILL_INDEX)

> 最后更新：2026-08-10

## 技能包列表

| # | 技能包 | 版本 | 说明 | 触发词 |
|---|--------|------|------|--------|
| 1 | dev-project-team-skill | v21.3.1 | 软件研发多角色编排器 | 全生命周期、启用角色、阶段评审 |
| 2 | document-processing | v1.0.0 | 文档处理技能包 | 拆分文档、章节拆分、docx拆分 |
| 3 | lark-training-ppt-generator | v5.0.0 | 培训PPT生成工作流 | 生成PPT、培训材料、演示文稿 |
| 4 | flowchart-skill | v1.0.0 | 流程图快速生成技能包 | 流程图、CSV转PPT、节点表、配色预设 |

## 技能包详细信息

### 1. DevProjectTeamSkill v21.3.1

**定位**：全生命周期多角色编排器

**角色包**：
- role-project-init（项目启动）
- role-requirements-analysis（需求）
- role-architecture（架构）
- role-development（开发）
- role-testing（测试）
- role-deployment（投产）
- role-governance（总控保障）

**加载路径**：`.trae/skills/dev-project-team-skill/SKILL.md`

---

### 2. DocumentProcessing v1.0.0

**定位**：文档处理技能包

**核心能力**：
- Word文档章节拆分
- 文档结构分析
- 格式转换

**工具脚本**：
- `tools/docx_splitter.py` — Word文档拆分工具

**加载路径**：`skills/document-processing/SKILL.md`

---

### 3. LarkTrainingPptGenerator v5.0.0

**定位**：培训课程PPT生成工作流

**核心能力**：
- 配置驱动主题系统
- 组件化架构
- 流程图生成
- Schema验证

**加载路径**：`SKILL.md`（根目录）

---

## 使用指南

### 加载技能包

```
# 加载编排器
加载 DevProjectTeamSkill

# 加载文档处理
加载 document-processing

# 加载PPT生成
加载 lark-training-ppt-generator
```

### 组合使用

```
# 多角色联合
启用需求分析师+测试工程师

# 文档拆分+PPT生成
拆分Word文档 → 生成流程图 → 制作培训PPT
```
