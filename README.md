# slideforge

**把「一段文字」或「一张参考图」变成真正可编辑的 Microsoft PowerPoint (`.pptx`)。**

不是输出整张贴图，不是 HTML —— 每个框、每行字、每张表、每条连接线、每个原生图表都是 PowerPoint 原生对象，
在 PowerPoint 里都能被选中、改字、改色、改版式。

> 融合了两条已验证的路线：
> - **image-to-pptx**：输入一张参考图，解码其布局/配色/文字/表格/图表，重建为原生可编辑对象。
> - **one-slide**：单页、来源可追溯、合成数据明确披露、只留一行数据来源的咨询式单页纪律。

## 两种入口

### 模式 A · RECREATE（输入一张图 → 复刻成可编辑 PPT）
给一张参考图（截图、别人的 slide、设计稿、白板、模板页），slideforge 会：
1. 读图、解构图（面板/box/表格/连接线/图表/图标各是什么）；
2. 把结构化内容重建为**原生对象**（可编辑）；
3. 像素级内容（真实照片/手绘插画/3D 渲染）作为独立图片层贴入，并诚实说明"这层不可编辑"；
4. 用本机 PowerPoint 渲染导出 PNG 做视觉核对。

### 模式 B · DESIGN（输入一段文字 → 现代咨询风可编辑 PPT）
给主题/要点，slideforge 按[设计系统](references/design-system.md)（Gartner / McKinsey / 大厂风格）输出：
渐变背景、彩色分区、字体层级（重点大/补充小/要点加粗）、一个强视觉隐喻、开源图标。

## 为什么不用图像模型
"美"可以写成可复用的设计规则（配色、字号层级、网格、渐变、图标），而不是靠 gpt-image / Nano Banana 重画。
这样产出是**原生可编辑**的，且不依赖任何图像 API / 密钥。

## 快速开始
```bash
pip install python-pptx

# DESIGN：写一个 scene.json，构建原生 .pptx
python scripts/slideforge_build.py examples/design_jv_plan.json -o out/jv_plan.pptx

# RECREATE：对着参考图标定 scene.json，再构建
python scripts/slideforge_build.py examples/recreate_data_sync.json -o out/data_sync.pptx

# 可选：渲染成 PNG 做视觉核对
python scripts/render_check.py out/data_sync.pptx
```

## 目录
```
slideforge/
├── SKILL.md              # 技能主文档（DSH / Claude Code / Codex 通用）
├── AGENTS.md             # 给 codex 等 agent 的入口索引
├── README.md
├── LICENSE               # MIT
├── scripts/
│   ├── slideforge_build.py   # 核心引擎：scene JSON -> 原生 .pptx
│   └── render_check.py       # 用本机 PPT/LibreOffice 渲染 PNG 做 QA
├── references/
│   ├── recreate-contract.md  # RECREATE 模式 scene 字段表
│   ├── design-system.md      # DESIGN 模式设计系统
│   └── integrity-contract.md # 来源追溯/合成披露/强调开关
└── examples/
    ├── design_jv_plan.json   # DESIGN 示例（跨品牌数据打通单页，通用占位）
    └── recreate_data_sync.json  # RECREATE 示例（数据覆盖与增量同步，通用占位）
```

## 安装到各 agent（同一份 SKILL.md 通用）
| 环境 | 放哪 |
|---|---|
| DSH / DeepSeek Harness（全局） | `~/.dsh/skills/slideforge/` |
| DSH（项目级） | `<项目根>/.agents/skills/slideforge/` 或 `.dsh/skills/` |
| Claude Code | `~/.claude/skills/slideforge/` |
| Codex / aider / cursor | 仓库根 `AGENTS.md` + 把 `slideforge/` 复制到项目的 `.agents/skills/` |

然后在 DSH / Claude / Codex 里对 agent 说「生成 / 复刻 / 做一份 PPT」即可，它会自动触发本 skill。

## 依赖
- Python 3.10+，`python-pptx`（唯一硬依赖，生成原生 OOXML）。
- 可选：本机 PowerPoint 或 LibreOffice（`render_check.py` 做视觉核对用），非生成所必需。

## 授权
MIT（`SKILL.md`、`scripts/`、`references/`、`AGENTS.md`、`README.md`、`examples/`）。内置图标为 Tabler Icons（MIT）。
