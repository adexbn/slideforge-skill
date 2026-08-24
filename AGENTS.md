# AGENTS.md — slideforge 给 codex / aider / cursor / claude 等 agent 的入口

本仓库是一个把「**一段文字**或**一张参考图**」变成**原生可编辑 Microsoft PowerPoint (.pptx)** 的技能包。
**权威文档在 [`SKILL.md`](./SKILL.md)**。任何涉及"做 / 生成 / 复刻 / 编辑 PPT、演示文稿、幻灯片、单页、
把这张图做成能改的 PPT"的请求，都先完整读 `SKILL.md` 再动手。下面是索引，不是正文。

## 一分钟索引

- **两种入口模式**（详见 SKILL.md）：
  - **RECREATE**：输入一张参考图 → 解码成结构化 scene → 重建为原生可编辑 `.pptx`。
  - **DESIGN**：输入一段文字 → 按 `references/design-system.md` 出现代咨询风原生单页。
- **核心引擎**：`python scripts/slideforge_build.py scene.json -o out.pptx`（原生 OOXML，零图像 API）。
- **视觉核对（可选）**：`python scripts/render_check.py deck.pptx`（本机 PowerPoint COM / LibreOffice 导出 PNG）。
- **契约文档**：
  - `references/recreate-contract.md` —— 复刻模式的 scene 元素字段表。
  - `references/design-system.md` —— DESIGN 模式的配色/字体/网格/渐变/图标规范。
  - `references/integrity-contract.md` —— 来源可追溯/合成披露/强调开关（OneSlide 思路）。
- **依赖**：Python 3.10+ 和 `python-pptx`。不需要任何图像生成 API。

## 对 agent 的硬约束

1. **只输出原生可编辑 `.pptx`**。禁止整页贴图冒充交付、禁止 HTML、禁止 `save as PNG` 当产品价值。
2. **先想清楚是 RECREATE 还是 DESIGN**。有参考图走 A；只有文字走 B；两者都要写清用户要什么。
3. **读图**（能看多模态就直接 `read_image`），解构图，再重建；不要凭猜硬搭。
4. **一页一个主结论 + 一个主图**（单页边界）。放不下返回范围超载，不硬塞第二页。
5. **来源可追溯 + 合成披露**：合成数据标"合成示例数据，非真实数据"。详见 integrity-contract。
6. **渲染 + 目测核对**后再交付；说明哪些是原生可编辑对象、哪些（如有）仍是位图贴入。
7. **不依赖图像模型**。把"美"写成设计系统，而不是靠 gpt-image/Nano Banana 重画。

## 不要做的事
- 不要输出整页图 / HTML / 用截图冒充 .pptx。
- 不要跳过 `recreate-contract.md` 的元素分类，把"能做成框/表/图的"硬贴成图片。
- 不要把合成/推断数据说成事实。
- 不要在缺 `python-pptx` 时硬跑；先确认依赖。
