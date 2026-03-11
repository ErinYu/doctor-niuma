---
title: "Medical PPT Enhancement - 深度研究报告类幻灯片生成"
type: feat
status: active
date: 2026-03-12
---

# Medical PPT Enhancement - 深度研究报告类幻灯片生成

## 概述

增强 MedClaw 的医学 PPT 生成能力，支持**深度研究报告类**幻灯片制作。包括：纯文字自动检索文献、文件附件解析、智能澄清追问、自动文件投递。

**当前状态**：基础 PPT 生成已实现（`medical-presentation` skill + `generate_ppt.py`），文件发送管道已就绪。

**目标**：让医生能用自然语言描述需求，MedClaw 自动检索文献、解析附件、生成专业 PPT 并直接发送到飞书/微信。

---

## 问题背景

### 现有能力

| 组件 | 状态 | 说明 |
|------|------|------|
| `medical-presentation` skill | ✅ 已有 | 基于 python-pptx，支持 4 种风格、5 种布局 |
| `generate_ppt.py` | ✅ 已有 | 核心生成器，已集成到容器 |
| `send_file` MCP 工具 | ✅ 刚实现 | Agent 可调用发送文件到用户 |
| 飞书文件发送 | ✅ 刚实现 | `FeishuChannel.sendFile()` 已实现 |

### 现有限制

1. **无文献检索能力**：纯文字输入时需手动查资料，不支持 PubMed/指南检索
2. **无文件附件解析**：PDF/PPT/Excel/图片附件无法读取
3. **澄清追问不足**：信息不足时追问不够智能
4. **未自动调用 send_file**：生成 PPT 后需手动告知用户文件位置

---

## 解决方案

### Phase 1: 集成 PubMed MCP 文献检索

**目标**：让 agent 能自动检索医学文献和指南。

**选择**：`JackKuo666/PubMed-MCP-Server` (104⭐，Python 实现，与容器技术栈一致)

**实施方案**：

1. **容器内克隆 MCP 服务器**：
   ```bash
   # Dockerfile 添加
   RUN git clone https://github.com/JackKuo666/PubMed-MCP-Server.git /app/mcp-servers/pubmed-mcp
   RUN cd /app/mcp-servers/pubmed-mcp && pip3 install -r requirements.txt
   ```

2. **MCP 配置** (`container/agent-runner/mcp-config.json`)：
   ```json
   {
     "mcpServers": {
       "pubmed": {
         "command": "python3",
         "args": ["/app/mcp-servers/pubmed-mcp/src/pubmed_mcp/server.py"]
       }
     }
   }
   ```

3. **Skill 指令更新** — 添加文献检索步骤：
   ```markdown
   ## 内容研究（增强版）

   **纯文字 + 学术报告类**：
   - 使用 `pubmed_search` 搜索：`"[topic]" AND ("guideline"[Filter] OR "review"[Filter])`
   - 时间过滤：近 3-5 年（`"2020"[PDAT]: "2026"[PDAT]`）
   - 优先获取指南（guideline）、系统综述（systematic review）
   - 提取关键数据：发病率、诊断标准、治疗推荐

   **有指定指南时**：
   - 使用 WebSearch 搜索指南 PDF（如 `"中国 2 型糖尿病防治指南 2023 PDF"`）
   - 用 Read 工具读取下载的文件路径
   ```

**验证标准**：
- [ ] Agent 能用 `pubmed_search` 找到相关文献
- [ ] 能从文献中提取结构化数据（发病率、诊断标准、治疗方案）
- [ ] 搜索结果标注 PMID 和年份

---

### Phase 2: 文件附件解析

**目标**：支持用户上传 PDF/PPT/Excel/图片，自动提取内容。

**实施方案**：

1. **PDF 解析**（指南、论文、出院记录）：
   ```bash
   # 容器依赖
   RUN apt-get install -y poppler-utils
   RUN pip3 install pdfplumber PyPDF2
   ```

   Skill 中添加 PDF 读取工具：
   ```python
   # container/agent-runner/pdf_reader.py
   import pdfplumber
   import sys

   def extract_text(path):
       with pdfplumber.open(path) as pdf:
           return "\n".join(page.extract_text() for page in pdf.pages)

   if __name__ == "__main__":
       print(extract_text(sys.argv[1]))
   ```

2. **Excel 解析**（临床研究数据）：
   ```bash
   RUN pip3 install openpyxl
   ```

   ```python
   # container/agent-runner/excel_reader.py
   import openpyxl

   def summarize_excel(path):
       wb = openpyxl.load_workbook(path)
       # 提取：表头、数据行数、统计类型、关键列
   ```

3. **图片 OCR**（CT 报告、化验单）：
   ```bash
   RUN pip3 install pytesseract
   RUN apt-get install -y tesseract-ocr tesseract-ocr-chi-sim
   ```

4. **Skill 指令更新**：
   ```markdown
   ## 文件附件处理

   **PDF 文件**：
   - 指南/论文：调用 `python3 /app/read_pdf.py <path>` 提取全文
   - 重点关注：推荐等级、治疗路径、诊断标准

   **Excel 文件**：
   - 调用 `python3 /app/read_excel.py <path>` 提取结构
   - 识别：数据表头、统计方法、P 值、结论

   **图片文件**：
   - 调用 `tesseract <path> stdout -l chi_sim+eng` 提取文字
   - 适用于：化验单、CT 报告截图
   ```

**验证标准**：
- [ ] 能从 PDF 指南中提取章节结构
- [ ] 能从 Excel 中识别数据表和统计结果

---

### Phase 2b: 图片 OCR 解析（可选）

**目标**：支持化验单、CT 报告截图的文字提取。

**实施方案**：

1. **安装 Tesseract OCR**：
   ```bash
   RUN apt-get install -y tesseract-ocr tesseract-ocr-chi-sim
   RUN pip3 install pytesseract
   ```

2. **Skill 指令**：
   ```markdown
   **图片文件**：
   - 调用 `tesseract <path> stdout -l chi_sim+eng` 提取文字
   - 适用于：化验单、CT 报告截图
   - **高风险**：OCR 可能出错，数字识别需人工核对
   ```

**验证标准**：
- [ ] 能从化验单截图中提取数值（可选，因准确性限制）

---

### Phase 3: 智能澄清追问

**目标**：用户输入过于简略时，主动收集必要信息，减少返工。

**实施方案**：

1. **追问参数矩阵**（已在 skill 中定义，需强化）：
   ```markdown
   | 参数 | 必填性 | 推断规则 | 追问措辞 |
   |------|--------|----------|----------|
   | topic | 必填 | 无 | "请问 PPT 主题是什么？" |
   | goal | 高 | 根据类型推断 | "这个 PPT 用于什么目的？（教学/汇报/科普）" |
   | audience | 中 | 默认住院医 | "听众是谁？（住院医/主治/患者/专家）" |
   | duration | 中 | 默认 15 分钟 | "汇报时长多少？" |
   | style | 低 | 根据 goal 推断 | "风格偏好？（学术/临床/简洁）" |
   ```

2. **追问触发逻辑**：
   - topic 缺失 → **必须追问**，不生成 PPT
   - goal/audience 缺失 → **推断默认值**，列出假设，允许直接开始或纠正
   - 其他参数缺失 → 使用默认值，不追问

3. **一次性追问模板**：
   ```markdown
   为了做出最适合你的 PPT，有几个问题：

   1. **主题**：请明确疾病或内容（如"2型糖尿病"）
   2. **用途**：教学培训/病例讨论/开题答辩？
   3. **听众**：住院医/主治/患者？
   4. **时长**：汇报多少分钟？
   5. **特殊要求**：有指定指南或参考文献吗？

   （有附件请直接发给我，我会自动解析）
   ```

**验证标准**：
- [ ] 主题缺失时 100% 追问
- [ ] 其他缺失信息 80% 能推断出合理默认值
- [ ] 追问措辞简洁、一次性问完

---

### Phase 4: 自动文件投递

**目标**：PPT 生成后自动调用 `send_file` 发送到用户。

**实施方案**：

1. **在 Skill 末尾添加 send_file 调用**：
   ```markdown
   ## 第五步：输出结果（增强版）

   生成 PPT 后，立即调用 send_file：

   1. **确认文件存在**：检查 /workspace/group/[filename].pptx
   2. **调用 MCP 工具**：使用 send_file({file_path: "/workspace/group/[filename].pptx", file_name: "[显示名].pptx"})
   3. **发送确认消息**：
      ```markdown
      ✅ PPT 已生成并发送！

      📁 文件：[filename].pptx（[X] 页）
      📋 类型：[ppt_type]
      🎨 风格：[style]

      **大纲预览**：
      1. 标题页
      2. [章节1]
      3. [页面2]
      ...

      💡 如需调整内容、风格、页数，直接告诉我。
      ```
   ```

2. **确认消息格式**：
      ```markdown
      ✅ PPT 已生成并发送！

      📁 文件：[filename].pptx（[X] 页）
      📋 类型：[ppt_type]
      🎨 风格：[style]

      **大纲预览**：
      1. 标题页
      2. [章节1]
      3. [页面2]
      ...

      💡 如需调整内容、风格、页数，直接告诉我。
      ```
   ```

2. **JSON schema 更新**（确保 send_file 能找到文件）：
   ```json
   {
     "output_path": "/workspace/group/[topic]_[type]_[date].pptx"
   }
   ```

**验证标准**：
- [ ] 生成的 PPT 自动出现在用户飞书/微信
- [ ] 确认消息包含文件名、页数、类型、风格
- [ ] 文件名规范、可搜索

---

## 技术架构

```
用户输入（飞书/微信）
       ↓
medical-presentation skill
       ↓
┌─────────────────────────────────────────────┐
│ 澄清追问（信息不足时）                        │
└─────────────────────────────────────────────┘
       ↓
┌─────────────────────────────────────────────┐
│ 内容研究                                     │
│  • pubmed_search (MCP)                       │
│  • WebSearch (指南 PDF)                      │
│  • Read 工具 (附件解析)                       │
└─────────────────────────────────────────────┘
       ↓
┌─────────────────────────────────────────────┐
│ 大纲设计（基于 ppt_type 模板）                │
└─────────────────────────────────────────────┘
       ↓
生成 PPT JSON → 调用 generate_ppt.py
       ↓
    /workspace/group/*.pptx
       ↓
send_file MCP 工具
       ↓
FeishuChannel.sendFile() → 用户收到文件 ✅
```

---

## 系统影响分析

### 交互图
```
Agent 生成 PPT
  → 调用 send_file({file_path: "/workspace/group/xxx.pptx"})
  → IPC 写入 {type: "send_file", relativePath: "xxx.pptx"}
  → ipc.ts 拾取并验证权限
  → index.ts 解析路径、防穿越检查
  → FeishuChannel.sendFile()
    → client.im.file.create() 上传文件
    → client.im.message.create(msg_type: "file") 发送消息
  → 用户在飞书收到文件
```

### 错误传播
| 失败点 | 检测方式 | 处理策略 |
|--------|----------|----------|
| PubMed API 超时 | MCP 工具返回 error | 降级到 WebSearch |
| PDF 解析失败 | read_pdf 返回非 200 | 告知用户文件损坏 |
| PPT 生成失败 | generate_ppt.py exit code 1 | 记录错误日志，要求简化内容 |
| send_file 失败 | IPC 错误日志 | 提示文件路径，手动取用 |

### 状态生命周期
- **无状态操作**：PPT 生成是独立的，不修改数据库
- **文件清理策略**：生成的 PPT 永久保留在 `/workspace/group/`
- **无回滚需求**：生成失败不影响其他数据

### 关键假设
- **文件可访问性**：假设用户上传的附件（PDF/Excel/图片）已可通过 `/workspace/group/` 访问
  - **实现**：飞书/钉钉/微信需在接收文件时下载到 group 目录
  - **当前缺口**：需在 channel 层实现文件接收功能（不在本计划范围）

### API 表面一致性
- **所有 channel**：`sendFile(jid, filePath, fileName)` 可选方法
- **已实现**：Feishu 完整实现
- **待实现**：钉钉、微信（降级为文本提示）

### 集成测试场景
1. 纯文字输入 → PubMed 检索 → 生成 PPT → 自动发送
2. 上传指南 PDF → 解析内容 → 生成解读 PPT → 自动发送
3. 上传去年 PPT → 读取内容 → 更新指南版本 → 优化视觉 → 自动发送
4. （可选）上传病例图片 → OCR 提取 → 生成病例讨论 PPT → 自动发送

---

## 验收标准

### 功能需求
- [ ] **Phase 1**：能通过 PubMed MCP 检索文献并提取数据
- [ ] **Phase 2a**：能解析 PDF、Excel 附件
- [ ] **Phase 2b**：能从图片中提取文字（可选）
- [ ] **Phase 3**：信息不足时主动追问，追问率 < 15%（即 >85% 请求能直接执行）
- [ ] **Phase 4**：生成的 PPT 自动发送到用户飞书

### 非功能需求
- [ ] **性能**：纯文字请求（含文献检索）< 60 秒
- [ ] **性能**：有附件请求 < 90 秒
- [ ] **质量**：生成的 PPT 内容准确率 > 90%（人工抽检）
- [ ] **安全**：PHI 自动脱敏（姓名 → 患者 X）

### 质量门槛
- [ ] 单元测试：`generate_ppt.py` 覆盖主要 layout
- [ ] 集成测试：端到端测试文件发送流程
- [ ] 文档：更新 SKILL.md 使用说明

---

## 依赖与风险

### 依赖项
| 依赖 | 版本要求 | 用途 |
|------|----------|------|
| `python-pptx` | ≥ 1.0.0 | PPT 生成（已有） |
| `pdfplumber` | ≥ 0.10.0 | PDF 解析（Phase 2a） |
| `openpyxl` | ≥ 3.1.0 | Excel 解析（Phase 2a） |
| `tesseract-ocr` | ≥ 5.0 | 图片 OCR（Phase 2b，可选） |
| PubMed MCP | latest | 文献检索（Phase 1） |
| `@larksuiteoapi/node-sdk` | latest | 飞书文件发送（已有） |

### 风险分析
| 风险 | 影响 | 概率 | 缓解措施 |
|------|------|------|----------|
| PubMed API 限流 | 文献检索失败 | 中 | 降级到 WebSearch |
| PDF 解析失败 | 附件内容丢失 | 低 | 提示用户文件损坏，要求文本粘贴 |
| 大文件上传超时 | 飞书发送失败 | 低 | 限制文件 < 10MB，或分拆 PPT |
| OCR 识别错误 | 化验单数据错误 | 高（如实施） | Phase 2b 标记为可选，高风险场景要求人工核对 |
| MCP 服务器克隆失败 | Phase 1 无法实施 | 低 | 提供 fallback WebSearch 方案 |

---

## 实施优先级

| Phase | 优先级 | 工作量 | 价值 |
|-------|--------|--------|------|
| Phase 1: PubMed 集成 | P0 | 2 天 | 高 — 学术报告必需 |
| Phase 3: 智能追问 | P0 | 1 天 | 高 — 减少 50% 返工 |
| Phase 4: 自动投递 | P0 | 0.5 天 | 高 — 用户体验 |
| Phase 2a: PDF/Excel 解析 | P1 | 2 天 | 中 — 增强场景 |
| Phase 2b: 图片 OCR | P2 | 1 天 | 低 — 可选，准确性受限 |

**推荐实施顺序**：Phase 4 → Phase 3 → Phase 1 → Phase 2a → Phase 2b

理由：Phase 4（自动投递）改动最小、立即可用；Phase 3（智能追问）纯 prompt 改进；Phase 1 需要集成 MCP；Phase 2a 需要容器依赖；Phase 2b OCR 准确性有限，优先级最低。

---

## 参考资料

### 外部参考
- [anthropics/skills - PPTX skill](https://github.com/anthropics/skills/tree/main/skills/pptx) — 官方 PPT 生成技能参考
- [JackKuo666/PubMed-MCP-Server](https://github.com/JackKuo666/PubMed-MCP-Server) — PubMed MCP 服务器
- [python-pptx 文档](https://python-pptx.readthedocs.io/) — PPT 生成库
- [pdfplumber 文档](https://github.com/jsvine/pdfplumber) — PDF 解析库

### 内部参考
- `container/skills/medical-presentation/SKILL.md` — 现有 PPT skill
- `container/agent-runner/generate_ppt.py` — PPT 生成器
- `src/channels/feishu.ts:179-211` — 飞书文件发送实现
- `container/agent-runner/src/ipc-mcp-stdio.ts:65-102` — send_file MCP 工具

### 相关工具验证报告
**已验证可用的工具**：
- ✅ `anthropics/skills` (pptx) — 90.6k⭐，官方维护
- ✅ `JackKuo666/PubMed-MCP-Server` — 104⭐，支持搜索/PDF 下载
- ✅ `python-pptx` — 成熟库，Docker 兼容
- ✅ `icip-cas/PPTAgent` — 3.5k⭐，文档转 PPT（备选）

**已确认不可用的工具**（幻觉/废弃）：
- ❌ `sickn33/pptx-official` — 不存在
- ❌ `composiohq/pubmed-mcp` — 不存在
- ❌ `GongRzhe/Office-PowerPoint-MCP-Server` — 已归档
- ❌ `ltc6539/mcp-ppt` — 仅 65⭐，8 commits
- ❌ 所有 `health-mcp/*` 和 `openhealth/*` 仓库 — 不存在
