---
name: slideforge
description: >-
  Build NATIVE, editable Microsoft PowerPoint (.pptx) decks — never a locked image and never HTML.
  Two entry modes in one skill:
    (A) RECREATE — given a reference image (screenshot, design mock, other slide, whiteboard), reproduce it
        as genuine editable PowerPoint objects (boxes, text, tables, charts, connectors, color blocks), so every
        element can be re-selected, re-typed, re-colored in PowerPoint.
    (B) DESIGN — from a text brief only, produce a modern consulting-grade slide (Gartner/McKinsey/vendor-doc feel):
        gradient background, color-coded logical parts, type hierarchy (large emphasis / small supplements / bold
        important notes), a strong visual metaphor, and an open-source icon set.
  Both modes keep OneSlide's discipline: one page = one main conclusion + one main visual, source-traceable,
  synthesized data clearly labeled, only a single "Data source:" footer. Use when the user wants to make / recreate /
  design a PowerPoint, a slide, a deck, a one-pager, or 做一份 PPT / 生成幻灯片 / 照着这张图复刻 / 把这个做成能改的 PPT.
  Requires python-pptx (native OOXML generation); no image-generation API needed.
metadata:
  version: "1.0.0"
  license: MIT
  author: adexbn
  combines:
    - image-to-pptx (reference-image recreation as native editable objects)
    - one-slide (single-page, source-traceable, editable PowerPoint discipline)
licenses:
  - content: "SKILL.md, scripts/, AGENTS.md, README.md"
    license: MIT
  - content: "references/design-system.md"
    license: MIT
---

# slideforge — 原生可编辑 PPT 生成器

**一句话**：把「一段文字」或「一张参考图」变成真正可编辑的 Microsoft PowerPoint（`.pptx`），
每个框、每行字、每张表、每条连接线都是原生对象 —— 不是整张图，不是 HTML。

## 铁律（任何时候都要遵守）

1. **只输出原生可编辑 `.pptx`**。禁止把一页当作整张 PNG/JPG 贴进来；禁止输出 HTML/网页幻灯片代
   替产品交付；禁止用 `save as PNG` 冒充价值。
2. **可编辑 = PowerPoint 原生对象**：文字用文本框、表格用 PPT 表格、图形用自选图形/形状、箭头用连
   接线、配色用实色填充或渐变。
3. **一页一个主结论 + 一个主图 + 一个主视觉**（OneSlide 单页边界）。不要为一页塞进第二个中心思想。
4. **来源可追溯**：区分「用户提供的事实」「稳定推断」「计算得出」「模型补全（待确认）」「外部核验」。
   合成数据必须写 `数据来源：... 合成示例数据，非真实数据`。绝不把推断/合成伪装成事实。
5. **只有一行页脚**：`数据来源：...`。方法、口径、公式、术语解释放进「PowerPoint 备注」，不放画面。
6. **不做无来源授权的地图/国界、不做地理分布**（用排序/矩阵/表格替代，返回阻断说明）。
7. **强调默认关闭**：不做按位置/顺序/排名的随机加粗、随机变色、随机填充。允许稳定可读性交替底色与
   `allowed_emphasis`（带 target/reason/method/source_ids）的业务强调，仅用加粗 + 浅色底纹，不改边框色。
8. **不删关键内容/不缩到不可读/不改变中心思想来凑一页**；放不下就返回范围超载，建议最值得做的那页。

## 两种入口模式

### 模式 A —— RECREATE（输入一张图 → 复刻成可编辑 PPT）
用户给一张参考图（截图、别人的 slide、设计稿、白板、模板页）。流程：

1. **读图**（如果当前 agent 能看多模态，直接 `read_image`；否则让用户描述关键结构）。
2. **解构图**，输出一个结构化 scene：画布尺寸（默认 16:9 `13.333×7.5in`，或按图的比例）、
   每个元素的 `type / content / x / y / w / h / fill / text_color / font_size / bold / z_index`。
   元素类型：`title` 标题、`panel` 面板、`box` 方块、`text` 文字、`table` 表格、`chart` 原生图表、
   `connector` 连接线/箭头、`icon` 图标、`image_layer`(真实照片/插画，只能当贴图)、`section` 分区。
3. **分类边界**：
   - 结构化内容（方框、表格、流程、色块、标题、图标、分栏）→ **原生对象重建**（高保真可编辑）。
   - 像素级内容（真实照片、手绘插画、3D 渲染、复杂艺术画面）→ 只能作为**位图嵌入**，明确告知用户
     "这部分不可编辑"。此时仍把外框/标题/说明做成原生对象，只把照片本身贴进去。
   - 图表（柱状/折线/饼）→ 优先用 **python-pptx 原生 chart**，让用户能改数据。
4. **重建**：用 `scripts/slideforge_build.py` 按场景 JSON 生成 `.pptx`。配 `--render` 用本机
   PowerPoint COM 导出 PNG 做视觉核对（可选，但建议）。
5. **核对**：渲染后读 PNG 检查重叠/越界/文字溢出/配色还原，修复确定性缺陷；不为了"更像"而反复漂
   移。人工目测后才能判定通过。

   **边界诚实**：结构化的图能高保真；照片/插画类只能当贴图。千万不要把一张照片图"用原生框硬仿"
   然后宣称完全可编辑——那既失真又自欺。

### 模式 B —— DESIGN（输入一段文字 → 现代咨询风可编辑 PPT）
用户给主题/要点。流程：

1. **先问三件事**（除非用户已给全）：内容与受众、风格偏好（或上传自己的 `.pptx` 模板）、
   页数/是否先单页冒烟。
2. **设计单页**：参考 `references/design-system.md` 的配色、字号层级、网格、圆角、渐变、图标规范，
   选一个有感的视觉隐喻（例如"多个品牌 → 汇入一个核心"）。
3. **构建**：同样走 `scripts/slideforge_build.py`，产出原生 `.pptx`。
4. **核对**：渲染 + 目测。

## 参考文件（按需读取，不要一次全读，省 context）
- `references/design-system.md` —— 设计系统规范（配色、字号层级、网格、渐变、图标、版式）。
- `references/recreate-contract.md` —— 复刻模式的结构化 scene 契约（每种元素的字段表）。
- `references/integrity-contract.md` —— 来源可追溯/合成披露/强调开关的判定标准（OneSlide 思路的落
  地规则）。

## 运行文件
- `scripts/slideforge_build.py` —— 核心引擎：按场景 JSON 生成原生 `.pptx`。
- `scripts/render_check.py` —— 用本机 PowerPoint COM 把 `.pptx` 导出 PNG 供目测；`--check` 只探测
  渲染后端是否可用。
- `examples/` —— 可运行的例子（一张参考图 + 一份场景 JSON + 生成的 pptx）。
- `tools/` —— 复刻/设计时要用的辅助（图标模板等），不要在这些文件里做无关修改。

## 运行要求
- Python 3.10+，`python-pptx`（生成原生 OOXML）。这是唯一硬依赖。
- 可选：本机 PowerPoint（做 `--render` 目测核对）或 LibreOffice；都不是生成所必需。
- 不需要任何图像生成 API（gpt-image-2 / Nano Banana 等）—— 我们把"美"写成设计系统，而不是靠图像
  模型重画。

## 交付报告
- 产出：真实 `.pptx`（原生可编辑）。
- 说明：哪部分是原生可编辑对象，哪部分（如有）仍是位图贴入；数据来源与任何合成披露；渲染核对状态
  （`RENDERED_AND_VERIFIED` 或 `not_tested`）。

## 授权
- SKILL.md、scripts/、AGENTS.md、README.md、references/：MIT。
- 内置图标：Tabler Icons（MIT）。
