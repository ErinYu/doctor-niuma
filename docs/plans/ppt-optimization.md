# PPT 优化工作项

基于 `docs/frontend-slides/` 设计系统 + NotebookLM 研究的改进清单。

## NotebookLM 设计原则（2026-03 研究成果）

### 核心理念
- **一页一概念**：每个 slide 只讲一个 idea，不堆砌
- **二选一法则**：要么大面积留白，要么信息密集，绝不两者平分
- **模板不存在**：每页都是重新排版，避免 cookie-cutter 感
- **自解释设计**：每页 slide 无需演讲者也能被理解（McKinsey 报告式）

### 视觉签名
- 标题占 slide 面积 30-50%，字号是正文的 10 倍
- 每页不超过 40 个词
- 配色控制在 2-3 色
- 无页码、无 footer、无 logo、无 running header
- 章节编号 (01, 02, 03) 作为导航元素
- 荧光笔高亮关键短语
- 单色剪影图片代替彩色照片
- 细线分隔 + 几何边框
- L 形裁剪标记营造海报感

### 布局模式
- 左右 60/40 或 70/30 分割（非对称）
- 三列卡片对比
- 纵向时间轴（文字左右交错）
- 图标网格 / 视觉列表
- 泡泡图 / Venn 图（线框风格）
- 大数字高亮（关键数据超大字号）

## P0 — 立即实施

### 1. 内容密度控制 + 自动分页
- 标题页: 1 标题 + 1 副标题，不超过 2 行
- 内容页: 1 标题 + 最多 4-6 个要点，每个要点 1 行
- 引用页: 1 句引言，最多 3 行
- 对比页: 左右各最多 3 个要点
- 超出限制时自动拆分为多页

### 2. 排版多样性（至少 4 种布局）
- 分栏布局: 左右 50/50 或 60/40 分割
- 卡片布局: 2-3 个并排卡片（对比、流程）
- 大数字高亮: 关键数据用超大字号突出
- 图文混排: 图片占 50%+ 面积，文字在侧边
- 时间轴: 纵向或横向流程图

## P1 — 高优

### 3. 签名设计元素（按风格分类装饰）
| 风格 | 签名元素 | PPTX 实现 |
|------|---------|----------|
| Bold Signal | 彩色卡片焦点 + 大号章节编号 | 圆角矩形 + 超大字号文本框 |
| Dark Botanical | 抽象渐变圆 + 细竖线分隔 | 渐变填充椭圆 + 细线 shape |
| Notebook | 纸质容器 + 彩色标签页 | 白色矩形底板 + 顶部小矩形色块 |
| Swiss Modern | 网格系统 + 红色强调线 | 虚线网格 + 红色粗线 |
| Neon Cyber | 深色底 + 发光边框 | 深色背景 + 亮色边框矩形 |
| Vintage | 衬线字体 + 装饰边框 | 双线矩形边框 + serif 字体 |

### 4. 背景效果（渐变 + 装饰形状）
- 渐变背景: `slide.background.fill` 渐变
- 半透明装饰: 大圆/矩形放背景层，50-80% 透明度
- 网格线: 细线 shape 组合

## ✅ 已完成

- [x] 真正渐变背景（通过 XML 操控实现）
- [x] 半透明装饰形状（圆、矩形叠加层，营造层次感）
- [x] 风格专属签名装饰（每种风格有独特的视觉元素）
- [x] 内容密度控制（`_enforce_density()` 自动截断，防止信息过载）
- [x] 新增 3 种 NotebookLM 风格（`minimalist`, `editorial`, `neo_retro`）
- [x] 新增 4 种 NotebookLM 布局（`big_number`, `split_panel`, `timeline`, `card_grid`）
- [x] 新增 3 种装饰器函数（`_decorate_minimalist`, `_decorate_editorial`, `_decorate_neo_retro`)
- [x] 风格预览功能（`--preview` 模式，生成所有 13 种风格的对比 PPTX）
- [x] 升级图片生成模型到 `google/gemini-3-pro-image-preview`
- [x] **默认风格改为 `clinical`**（医学专业蓝色，绝不使用 neon_cyber 作为默认）
- [x] **激进的自动配图**（`auto_generate_images_aggressive()` 为所有内容页生成配图）
- [x] **SKILL.md 更新**（明确配图自动生成、表格使用场景）

### 预览功能使用

当用户不确定想要哪种风格时：
```bash
python3 /app/generate_ppt.py --preview --output /workspace/group/style_preview.pptx
```
生成一个 13 页的 PPTX，每页展示一种风格。用户打开后可以直观对比选择。

## 下一步
- [ ] 字体安装（思源黑体、霞鹜文楷等）—— 黿免系统字体 fallback
- [ ] 动画效果（PPTX 原生不支持，考虑 HTML 输出模式）

### 5. 字体升级（容器预装 + 配置）
- 思源黑体 + 思源宋体
- 阿里巴巴普惠体 + 霞鹜文楷
- LXGW WenKai（开源免费，手写感）

### 6. 配图 prompt 优化
- 图片尺寸匹配 slide 布局
- prompt 结合具体 slide 内容
- Gemini 超时/失败 fallback，不留白

## 设计原则（来自 frontend-slides）

- **"No AI slop"**: 禁止 Inter/Roboto/Arial 等泛用字体，禁止千篇一律的蓝白配色
- **每种风格必须有辨识度**: 不能只靠颜色区分，要有独特的视觉签名
- **留白 > 堆砌**: 信息密度严格控制，宁可多分页也不塞满
- **图片是主角**: 配图占比至少 40-50%，而非装饰点缀

## 模型配置

- 图片生成: `google/gemini-3-pro-image-preview` via Zenmux Vertex AI proxy
- 文本分析: Claude Agent SDK (anthropic/claude-opus-4.6)
