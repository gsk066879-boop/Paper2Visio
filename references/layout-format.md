# 可编辑 Visio 布局接口

脚本负责确定性绘制，不替代文章理解、生图或人工视觉检查。按本次参考图建立 JSON；所有坐标是像素式设计单位，原点左上，x 向右、y 向下，实际英寸为单位除以 `units_per_inch`。

顶层：`width`、`height`、`units_per_inch`（默认80）、`background`（默认 `#FAF8E8`）、`font`（默认 `Microsoft YaHei UI`）、`min_font_size`（默认10，设计单位）、`elements` 列表。主图不得使用图片图元。

每个图元必须有唯一 `id`。绘制顺序为列表顺序，连接线最后绘制。

- `rect`：`x,y,w,h`；可选 `fill,line,r,weight`（线宽单位 pt）；无填充或无边框用 null。
- `text`：`x,y,w,h,text,size,color,bold,align`；size 是设计单位，align 为0左/1中/2右。换行用真实 `\n`。文字独立可编辑。脚本不自动缩小文本；溢出会报错，先修改布局。可选 `runs`，每项是 Python 字符下标 `start,end` 与 `color`，可局部着色。
- `poly`：`points:[[x,y],...]`、`fill,line,weight`；闭合多边形需首尾同点。
- `line`：`points`、`color,weight,arrow`；用于括号、表格线、归纳线等自由几何，不具备语义端点粘附。
- `connector`：`from`、`to` 为已有 rect 图元 id；`from_port`、`to_port` 为 `[fx,fy]`，在目标框局部坐标中左下为 `[0,0]`、右上为 `[1,1]`。常用底中 `[0.5,0]`、顶中 `[0.5,1]`、右中 `[1,0.5]`、左中 `[0,0.5]`。可选 `arrow`（默认true）、`color`、`weight`。采用 Visio 原生直角动态连接线，端点粘附，节点移动后维持连接。为实际连接预留无文字通道；自动路由仍须检查。

文字与框体可以分别选择编辑；移动完整模块时同时选中它的框体与文字（可在 Visio 中分组），连接端点附着在框体上。

最小示例（仅演示接口，不是论文模板）：

```json
{
  "width": 640, "height": 480, "units_per_inch": 80,
  "elements": [
    {"id":"a","kind":"rect","x":60,"y":50,"w":210,"h":75,"fill":"#EAF5FA","line":"#155675"},
    {"id":"at","kind":"text","x":65,"y":55,"w":200,"h":65,"text":"研究对象","size":20,"align":1,"bold":true},
    {"id":"b","kind":"rect","x":330,"y":260,"w":210,"h":75,"fill":"#EAF5FA","line":"#155675"},
    {"id":"bt","kind":"text","x":335,"y":265,"w":200,"h":65,"text":"分析方法","size":20,"align":1,"bold":true},
    {"id":"ab","kind":"connector","from":"a","to":"b","from_port":[1,0.5],"to_port":[0,0.5]}
  ]
}
```

运行 `--check-only` 不创建 Visio 文档，验证图元、文字尺寸、范围和端点引用。字体测量使用已安装的微软雅黑字体文件；不通过就修正布局，不强制绕过。

完整生成写出 `<name>.vsdx`、`<name>-preview.png`、`<name>-layout.json`、`<name>-validation.json`、`<name>-render.py`、`<name>-replay.ps1`。提供 `--reference` 时额外复制参考图并创建对照页。重放脚本依赖目标计算机的 Python、pywin32、Pillow 和 Visio；输出新版本名，不覆盖原图。

验证报告包含主图原生对象/文字/连接数量、主图图片对象数量、连接端点记录、保存后重开结果及预览尺寸。它不声称通过视觉验收。不要用包内论文例图充当新文章的参考图。
