# PPT 生成待办事项

## 🔥 P0: 最高优先级

### 1. 文件上传与理解（飞书机器人集成）

**目标：** 支持用户上传文件（PDF/Word/图片），根据文件内容生成 PPT

**待实现：**

#### 1.1 飞书文件交互设计
- [ ] 调研飞书机器人文件上传 API
  - [ ] 支持的文件类型（PDF、DOCX、图片、TXT）
  - [ ] 文件大小限制
  - [ ] 获取文件内容的 API 接口
- [ ] 设计用户交互流程：
  ```
  用户上传文件 → 机器人确认收到 → 询问 PPT 需求 → 生成 PPT
  ```
- [ ] 支持多文件同时上传
- [ ] 文件 + 文字描述混合输入

#### 1.2 文件内容解析
- [ ] 选择合适的模型解析文件内容：
  | 文件类型 | 推荐模型 | 原因 |
  |----------|----------|------|
  | PDF（纯文字） | Claude/GPT-4 | 文本提取能力强 |
  | PDF（图文混排） | Claude Vision / GPT-4V / **Gemini 3.1** | 多模态理解 |
  | DOCX | Claude/GPT-4 | 结构化文本解析 |
  | 图片 | Claude Vision / GPT-4V / **Gemini 3.1** | OCR + 内容理解 |
  | 图片生成 | **Gemini 3.1 Flash Image** | 内置图像生成能力 |

- [ ] **Gemini 3.1 Flash Image 模型集成**（推荐用于图文处理）：
  ```python
  from google import genai
  from google.genai import types

  client = genai.Client(
      api_key="$ZENMUX_API_KEY",
      vertexai=True,
      http_options=types.HttpOptions(api_version='v1', base_url='https://zenmux.ai/api/vertex-ai')
  )

  # 文本生成
  response = client.models.generate_content(
      model="google/gemini-3.1-flash-image-preview",
      contents="How does AI work?"
  )
  print(response.text)

  # 图像生成（可用于 PPT 配图）
  prompt = "Create a medical illustration of heart anatomy"
  response = client.models.generate_content(
      model="google/gemini-3.1-flash-image-preview",
      contents=[prompt],
      config=types.GenerateContentConfig(
          response_modalities=["TEXT", "IMAGE"]
      )
  )

  # 处理响应（文本 + 图片）
  for part in response.parts:
      if part.text is not None:
          print(part.text)
      elif part.inline_data is not None:
          image = part.inline_data
          # 保存生成的图片
          image.save("generated_image.png")
  ```

- [ ] PDF 图片提取：
  - [ ] 使用 `pdf2image` 将 PDF 页转为图片
  - [ ] 使用 Vision 模型理解图片内容
- [ ] 表格识别与提取（PDF/图片中的表格）

#### 1.3 AI 图片生成与嵌入
- [ ] 使用 Gemini 3.1 Flash Image 生成 PPT 配图：
  ```python
  # 根据幻灯片内容生成配图
  slide_topic = "surgical procedure comparison"
  prompt = f"Create a clean medical illustration showing {slide_topic}, professional style, suitable for presentation"

  response = client.models.generate_content(
      model="google/gemini-3.1-flash-image-preview",
      contents=[prompt],
      config=types.GenerateContentConfig(
          response_modalities=["TEXT", "IMAGE"]
      )
  )

  # 提取图片并保存到工作目录
  for part in response.parts:
      if part.inline_data is not None:
          image_path = f"/workspace/group/generated_images/{slide_topic}.png"
          part.inline_data.save(image_path)
  ```
- [ ] 图片插入到 PPT：
  - [ ] 自动根据页面内容判断是否需要配图
  - [ ] 使用 `image_left` / `image_right` / `image_top` 布局
  - [ ] 图片路径写入 PPT JSON 的 `images` 字段
- [ ] 图片生成场景：
  | 场景 | Prompt 示例 |
  |------|-------------|
  | 医学插图 | "medical illustration of [器官/手术流程]" |
  | 流程图 | "flowchart showing [治疗流程/诊断路径]" |
  | 对比图 | "side-by-side comparison diagram of [A vs B]" |
  | 数据可视化 | "infographic showing [统计数据]" |

#### 1.4 内容结构化
- [ ] 从文件中提取关键信息：
  - [ ] 标题/主题
  - [ ] 章节结构
  - [ ] 关键数据/图表
  - [ ] 图片资源
- [ ] 生成 PPT 大纲（JSON 格式）
- [ ] 用户确认大纲后再生成 PPT

---

### 2. PPT 视觉效果优化（借鉴 frontend-slides）

**目标：** 参考 `docs/frontend-slides` 的设计系统，提升 PPT 视觉效果

**待实现：**

#### 2.1 引入 frontend-slides 设计系统
- [ ] 复用 12 种风格预设（STYLE_PRESETS.md）：
  - Dark: Bold Signal, Electric Studio, Creative Voltage, Dark Botanical
  - Light: Notebook Tabs, Pastel Geometry, Split Pastel, Vintage Editorial
  - Specialty: Neon Cyber, Terminal Green, Swiss Modern, Paper & Ink
- [ ] 复用字体配对系统：
  | 风格 | 标题字体 | 正文字体 |
  |------|----------|----------|
  | Bold Signal | Archivo Black | Space Grotesk |
  | Dark Botanical | Cormorant | IBM Plex Sans |
  | Notebook Tabs | Bodoni Moda | DM Sans |
- [ ] 复用色彩方案（CSS 变量形式）

#### 2.2 动画效果
- [ ] 入场动画（animation-patterns.md）：
  - [ ] Fade + Slide Up
  - [ ] Scale In
  - [ ] Blur In
- [ ] 背景效果：
  - [ ] Gradient Mesh（渐变网格）
  - [ ] Grid Pattern（网格图案）
- [ ] 按场景匹配动画风格：
  | 场景 | 动画风格 |
  |------|----------|
  | 学术汇报 | Professional（简洁快速） |
  | 商务演示 | Dramatic（缓慢大气） |
  | 技术分享 | Techy（霓虹/网格） |

#### 2.3 布局优化
- [ ] 复用 viewport 拟合规则（viewport-base.css）
- [ ] 内容密度限制：
  | 页面类型 | 最大内容 |
  |----------|----------|
  | 标题页 | 1 标题 + 1 副标题 |
  | 内容页 | 1 标题 + 4-6 要点 |
  | 表格页 | 1 标题 + 1 表格（≤6 列）|
- [ ] 避免内容溢出，超出自动分页

#### 2.4 装饰元素
- [ ] 侧边装饰条（accent_bar）
- [ ] 角落装饰（corner_accent）
- [ ] 分割线（divider）
- [ ] 图标式要点（icon_bullets）

---

## ✅ 已完成

### 3. 文字+文件输入场景支持
- [x] 解析用户文字描述，提取关键要求
- [x] 同时读取上传文件内容
- [x] 合并文字要求与文件内容，生成结构化大纲

### 4. 布局类型（12+ 种）
- [x] three_col, left_sidebar, right_sidebar
- [x] image_left, image_right, image_top
- [x] quote, center_focus, comparison
- [x] process, chart, table

### 5. 飞书卡片渲染优化
- [x] `sendCard` MCP 工具支持
- [x] IPC 层 `send_card` 消息类型处理
- [x] router 层 `routeCardOutbound` 函数

### 6. PPT 视觉多样化规则
- [x] 布局多样化规则
- [x] 内容拆分规则（超过 6 个要点拆页）

---

## 📋 Backlog（低优先级）

- [ ] 自动图片推荐/生成（AI 驱动）
- [ ] 模板继承和自定义模板
- [ ] 嵌入视频支持
- [ ] 高级图表类型（雷达图、树状图等）
- [ ] 确保文字+文件信息不冲突、互补融合
