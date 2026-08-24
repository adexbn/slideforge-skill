# recreate-contract.md —— RECREATE 模式的结构化 scene 契约

RECREATE 模式把一张参考图**解码**成下面的 scene JSON，再交给 `scripts/slideforge_build.py` 重建为原生 `.pptx`。
这个契约是"看图 → 结构化 → 重建"之间的唯一语言。

## 画布
```json
{ "canvas": "16:9" }
```
- `16:9` → 13.333×7.5in（默认）；`9:16` → 7.5×13.333；`3:4` → 7.5×10（知识图文/竖版）。
- 按参考图宽高比选择；非 16:9 时按目标画布重排，**不要裁切/拉伸原图**。

## slide 顶层结构
```json
{
  "canvas": "16:9",
  "slides": [
    {
      "background": {"type":"gradient","color1":"#F7FAFD","color2":"#E9F1FA","angle":90},
      "elements": [ /* 见下 */ ]
    }
  ]
}
```

## 元素类型（字段表）

### text —— 文本框（标题/正文/标签）
```json
{"type":"text","x":0.85,"y":0.4,"w":11,"h":0.5,"size":30,"color":"#15375F",
 "bold":true,"align":"left","anchor":"top","content":"标题"}
```

### box —— 自选图形（框架、卡片、色块、形状）
```json
{"type":"box","x":0.6,"y":1.5,"w":4,"h":2.5,"fill":"#F0F5FB","line":"#AEB8C6",
 "line_w":1.0,"radius":0.08,"content":[["行1",11,"#172033",false],["行2",11,"#172033",false]],
 "font_size":11,"bold":false,"align":"center","anchor":"middle"}
```
- `radius` 存在 → 圆角矩形；省略 → 直角矩形。
- 用一个小 `box`（高 0.06–0.09）做「顶部强调条」。

### table —— 原生 PowerPoint 表格
```json
{"type":"table","x":4,"y":2.6,"w":3.4,"h":1.8,
 "header":["A","B"],
 "rows":[["1","x"],["2","y"]],
 "rowfills":[["#FFF0E2","#FFF0E2"],["#E4F4E9","#E4F4E9"]],
 "header_fill":"#15375F","font_size":10.5}
```
- 用 `rowfills` 做**逐行/逐列**状态色（如绿=INSERT、橙=UPDATE、红=DELETE）。这是区分逻辑的常用做法。

### connector —— 连接线/箭头（流程、指向）
```json
{"type":"connector","x1":4,"y1":2.2,"x2":6,"y2":2.2,"color":"#2F6FB2","w":2.5,"arrow":true,"dash":null}
```
- 直线连接线是原生对象；需要曲线/折线时可分多段，或用 block-arrow 自选图形。

### bar_chart / (可扩展 line_chart, pie_chart) —— 原生图表
```json
{"type":"bar_chart","x":8,"y":2,"w":4,"h":3,"title":"Sales",
 "categories":["A","B","C"],
 "series":[{"name":"Sales","values":[1,2,3]}],
 "color":"#2F6FB2","legend":true}
```
- 用 PPT 原生图表让用户能改数据；不要拿图片冒充图表。

### icon —— 图标（语义标记）
```json
{"type":"icon","x":1,"y":2,"w":0.5,"h":0.5,"name":"database","color":"#2F6FB2","size":8}
```
- 生产流程：先在 build.py 里画一个简单占位（圆形+字母），随后用 PPT COM 插入真实 Tabler SVG
  覆盖同位置，保持矢量可缩放。占位仅保证"没有 SVG 时布局可见"。

### image —— 真实图片层（照片/插画，仅当位图嵌入）
```json
{"type":"image","x":3,"y":2,"w":4,"h":2.5,"path":"/abs/path.png"}
```
- **只有像素级内容**（真实照片、手绘插画、3D 渲染、复杂艺术画面）才用 image。此时把它作为独立图片
  对象叠在原生对象之上，并明确告知用户"这层图片不可编辑"。永远不要用它替代"能做成框/表格/文字的
  结构化内容"。

## 复刻时的分类决策（重要）
| 参考图里的内容 | 用什么 |
|---|---|
| 方框、面板、色块、标题、标签 | `box` / `text`（原生可编辑） |
| 表格、行列、状态色 | `table`（原生可编辑） |
| 流程、箭头、指向 | `connector`（原生可编辑） |
| 柱状/折线/饼图 | `bar_chart`（原生可编辑、可改数据） |
| 图标 | `icon`（Tabler SVG 矢量） |
| 真实照片、手绘插画、3D、艺术画面 | `image`（仅位图嵌入，告知不可编辑） |

记住：**能做成框/表/字/图的，绝不贴图**；真没法做成对象的，才作为贴图并诚实说明。
