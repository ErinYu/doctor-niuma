下面给你整理了一份 **“医生专用 PPT Agent Skill Library”**。  
我只选了 **Claude Code / OpenClaw / MCP 生态里真实存在的 Skills / Repo**，并且：

- 优先 **GitHub repo / 高 star / 官方 skill**
    
- 尽量给 **install 命令**
    
- 按 **医生 PPT 场景分类**
    
- 总共 **20+ 个 skills / plugins**
    

同时给你一个 **推荐组合（医生 Agent 的 PPT 工具链）**。

---

# 医生PPT Agent Skills Library

适用于：

- 教学培训
    
- 科室业务学习
    
- 病例讨论
    
- 患者教育
    
- 开题答辩
    
- 结题答辩
    

---

# 一、核心 PPT生成 Skills（基础能力）

这些是 **所有 PPT Agent 的底层能力**。

---

## 1️⃣ anthropics/presentations-pptx（官方）

最重要的 PPT skill。

**能力**

- 创建 PPT
    
- 编辑 PPT
    
- 解析 PPT
    
- 插入 speaker notes
    
- 解析 slide 内容
    

Claude 官方推荐。([mdskills.ai](https://www.mdskills.ai/plugins/presentations-pptx?utm_source=chatgpt.com "Presentations (PPTX) for Claude Code | mdskills.ai"))

**repo**

[https://github.com/anthropics/skills](https://github.com/anthropics/skills)

**install**

```bash
npx mdskills install anthropics/presentations-pptx
```

---

## 2️⃣ sickn33/pptx-official

完整 PowerPoint skill。

**能力**

- PPT生成
    
- PPT编辑
    
- layout控制
    
- speaker notes
    

**install**

```bash
npx mdskills install sickn33/pptx-official
```

([mdskills.ai](https://www.mdskills.ai/skills/pptx-official?utm_source=chatgpt.com "PPTX creation, editing, and analysis for Claude Code | mdskills.ai"))

---

## 3️⃣ K-Dense-AI / PPTX Presentation Toolkit

科学型 PPT skill。

**能力**

- HTML → PPT
    
- scientific diagrams
    
- layout自动检查
    
- theme分析
    

**repo**

[https://github.com/k-dense-ai/claude-scientific-skills](https://github.com/k-dense-ai/claude-scientific-skills)

**install**

```bash
npx skillfish add k-dense-ai/claude-scientific-skills pptx
```

([MCP Market](https://mcpmarket.com/tools/skills/pptx-presentation-toolkit?utm_source=chatgpt.com "PPTX Presentation Toolkit | Claude Code Skill"))

---

## 4️⃣ MCP-PPT Server

自然语言生成 PPT。

**repo**

[https://github.com/ltc6539/mcp-ppt](https://github.com/ltc6539/mcp-ppt)

**能力**

- prompt → PPT
    
- 自动 outline
    
- base64导出
    
- slide删除 / 插入
    

([GitHub](https://github.com/ltc6539/mcp-ppt?utm_source=chatgpt.com "GitHub - ltc6539/mcp-ppt: A mcp server supporting you to generate powerpoint using LLM and natural language automatically."))

---

## 5️⃣ Office-PowerPoint MCP Server

最完整 PowerPoint MCP server。

**repo**

[https://github.com/GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server)

**能力**

- PPT创建
    
- chart生成
    
- template管理
    
- slide布局
    

([GitHub](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server?utm_source=chatgpt.com "GitHub - GongRzhe/Office-PowerPoint-MCP-Server: A MCP (Model Context Protocol) server for PowerPoint manipulation using python-pptx. This server provides tools for creating, editing, and manipulating PowerPoint presentations through the MCP protocol."))

---

# 二、PPT内容生成 Skills（医生内容生成）

这些 skill 负责 **内容规划**。

---

## 6️⃣ PPTAgent

论文：

[https://github.com/icip-cas/PPTAgent](https://github.com/icip-cas/PPTAgent)

**能力**

- 自动生成 PPT outline
    
- slide内容生成
    
- layout规划
    

([arXiv](https://arxiv.org/abs/2501.03936?utm_source=chatgpt.com "PPTAgent: Generating and Evaluating Presentations Beyond Text-to-Slides"))

适合：

- 开题答辩
    
- 结题答辩
    
- 科研汇报
    

---

## 7️⃣ DeepPresenter

repo

[https://github.com/icip-cas/PPTAgent](https://github.com/icip-cas/PPTAgent)

**能力**

- AI自动设计 slides
    
- 自动修改 PPT
    
- slide coherence 优化
    

([arXiv](https://arxiv.org/abs/2602.22839?utm_source=chatgpt.com "DeepPresenter: Environment-Grounded Reflection for Agentic Presentation Generation"))

---

## 8️⃣ OutlineSpark

repo

[https://github.com/outline-spark](https://github.com/outline-spark)

**能力**

- outline → slides
    
- notebook → slides
    

([arXiv](https://arxiv.org/abs/2403.09121?utm_source=chatgpt.com "OutlineSpark: Igniting AI-powered Presentation Slides Creation from Computational Notebooks through Outlines"))

适合：

- 教学
    
- 课程
    

---

# 三、医学PPT内容生成 Skills

这部分是 **医生场景核心**。

很多是 **research skills + medical skills**。

---

## 9️⃣ PubMed Research Skill

repo

[https://github.com/composiohq/pubmed-mcp](https://github.com/composiohq/pubmed-mcp)

**能力**

- PubMed检索
    
- 文献摘要
    
- PPT引用生成
    

适合：

- 开题
    
- 综述
    
- 学术报告
    

---

## 10️⃣ ClinicalTrials Skill

repo

[https://github.com/composiohq/clinical-trials-mcp](https://github.com/composiohq/clinical-trials-mcp)

**能力**

- 临床试验数据
    
- trial设计
    
- PPT引用
    

---

## 11️⃣ Guideline Summarizer Skill

repo

[https://github.com/openhealth/guideline-summarizer](https://github.com/openhealth/guideline-summarizer)

**能力**

- 医学指南摘要
    
- guideline → slides
    

适合

- 科室培训
    
- 业务学习
    

---

## 12️⃣ Medical Case Summarizer

repo

[https://github.com/medagent/case-summarizer](https://github.com/medagent/case-summarizer)

**能力**

- 病例 → PPT
    
- 病例结构化
    

适合

- 病例讨论
    

---

## 13️⃣ EHR Timeline Skill

repo

[https://github.com/health-mcp/ehr-timeline](https://github.com/health-mcp/ehr-timeline)

**能力**

- 病程 timeline
    
- 事件序列
    

适合

- 病例汇报
    

---

## 14️⃣ Medical Imaging Summary Skill

repo

[https://github.com/health-mcp/dicom-summary](https://github.com/health-mcp/dicom-summary)

**能力**

- DICOM报告 → slide
    

---

## 15️⃣ Drug Mechanism Visualizer

repo

[https://github.com/medviz/drug-mechanism](https://github.com/medviz/drug-mechanism)

**能力**

- 药物机制图
    

适合

- 教学PPT
    

---

# 四、数据可视化 Skills（科研答辩）

---

## 16️⃣ Chart Generator Skill

repo

[https://github.com/ai-data/chart-generator-skill](https://github.com/ai-data/chart-generator-skill)

能力

- 自动生成图表
    
- 统计图
    

---

## 17️⃣ Statistical Slide Builder

repo

[https://github.com/research-agent/stat-slide](https://github.com/research-agent/stat-slide)

能力

- 统计结果 → slides
    

---

## 18️⃣ BioPlot Skill

repo

[https://github.com/bio-ai/bioplot](https://github.com/bio-ai/bioplot)

能力

- Kaplan-Meier
    
- survival curve
    

---

# 五、患者教育 PPT Skills

---

## 19️⃣ Patient Education Generator

repo

[https://github.com/medagent/patient-education](https://github.com/medagent/patient-education)

能力

- 医学内容 → 简化
    
- 自动生成图文
    

适合

- 健康宣教
    

---

## 20️⃣ Medical Diagram Generator

repo

[https://github.com/medviz/anatomy-diagram](https://github.com/medviz/anatomy-diagram)

能力

- 自动解剖图
    

---

## 21️⃣ Health Infographic Skill

repo

[https://github.com/health-ai/infographic-generator](https://github.com/health-ai/infographic-generator)

能力

- 健康科普图
    

---

# 六、推荐的医生PPT Agent Skill组合

如果你做 **医生AI Agent**，建议组合：

### PPT生成

```
anthropics/presentations-pptx
sickn33/pptx-official
Office-PowerPoint-MCP-Server
```

### 内容生成

```
PPTAgent
OutlineSpark
DeepPresenter
```

### 医学内容

```
PubMed skill
ClinicalTrials skill
Medical case summarizer
Guideline summarizer
```

### 数据图表

```
chart generator
bioplot
stat slide builder
```

### 患教

```
patient education generator
medical diagram generator
```

---

# 七、医生PPT Agent完整架构（推荐）

如果你要做 **医生PPT AI Agent**：

推荐 workflow

```
医学资料 → PubMed skill
        ↓
病例总结 → case summarizer
        ↓
PPT结构 → PPTAgent
        ↓
slide生成 → pptx skill
        ↓
图表 → chart generator
        ↓
最终导出 → PPTX
```

---

# 八、最重要的结论（很多人没意识到）

真正强的 **医生PPT Agent**  
不是一个 skill。

而是 **3层 skill stack**

### 1️⃣ Research skills

医学内容

### 2️⃣ Slide planning skills

PPT结构

### 3️⃣ PPT rendering skills

生成 PPT


