# Paper2Visio

Paper2Visio 是一个 Codex Skill，用于把论文或技术文章转化为高密度学术体系图：先提炼科学内容并生成视觉参考，再用 Microsoft Visio 原生图元重建、核验并交付可编辑的 `.vsdx` 文件。

它适合技术路线图、研究框架图、方法体系图和论文精读总图。Skill 会保留论文中的适用条件、证据边界和局限，不会为了版式补造方法、因果关系或实验结论。

## 工作流程

1. 阅读论文，记录研究对象、问题、方法、验证、结论与来源位置。
2. 生成内容简报、文献依据和可独立执行的英文生图提示词。
3. 使用当前环境可用的图像生成能力制作并检查参考图。
4. 根据参考图编写布局 JSON，并用 `scripts/render_visio.py` 生成原生可编辑 Visio。
5. 保存后重新打开文件，检查文字、连接端点、原生对象数量和导出预览。

## 安装

克隆仓库到 Codex Skills 目录：

```powershell
git clone https://github.com/gsk066879-boop/Paper2Visio.git "$env:USERPROFILE\.codex\skills\paper2visio"
python -m pip install -r "$env:USERPROFILE\.codex\skills\paper2visio\requirements.txt"
```

重新启动 Codex 后，可显式调用 `$paper2visio`；仓库也允许 Codex 根据任务内容自动选择该 Skill。

## 环境要求

- Windows 10/11
- Microsoft Visio 桌面版
- Python 3.10 或更高版本
- `pywin32` 和 `Pillow`
- 用于读取论文的 PDF/文档能力
- 可选的图像生成能力；不可用时会按 Skill 中的降级规则继续制作 Visio

## 使用

在 Codex 中附上论文或文章，然后输入：

```text
使用 $paper2visio 阅读我提供的文章，严格按完整话术生成参考图，再重建并核验可编辑 Visio，优化箭头后直接交付。
```

若只需要内容分析或生图提示词，请在当前请求中明确阶段范围。

## 渲染器

先检查布局，不启动 Visio：

```powershell
python scripts/render_visio.py path\to\layout.json --check-only
```

生成 Visio 文件和验证材料：

```powershell
python scripts/render_visio.py path\to\layout.json `
  --output-dir outputs `
  --name paper-roadmap `
  --reference path\to\reference.png
```

输出包括 `.vsdx`、PNG 预览、布局 JSON、验证报告、渲染脚本副本和 PowerShell 重放脚本。详细字段见 [references/layout-format.md](references/layout-format.md)。

## 仓库结构

```text
paper2visio/
|-- SKILL.md
|-- agents/openai.yaml
|-- assets/style-reference.png
|-- references/layout-format.md
|-- references/original-user-brief.txt
`-- scripts/render_visio.py
```

## 许可证

代码、Skill 指令和仓库自有文档采用 [MIT License](LICENSE)。`assets/style-reference.png` 仅作为版式与视觉风格参考；其中引用的论文题名、作者及研究内容仍归各自权利人所有，不因本仓库许可证而改变。

---

## English

Paper2Visio is a Codex Skill that turns papers and technical articles into dense academic system diagrams. It extracts source-grounded content, creates and reviews a visual reference, then reconstructs the diagram with native editable Microsoft Visio shapes and validates the saved `.vsdx` output.

See the Chinese sections above for installation, requirements, commands, and licensing details.
