---
name: academic-docx-formatter
description: "Standardize and normalize formatting for Chinese academic papers, competition reports (数媒/挑战杯/互联网+/OPC), and graduation theses in Microsoft Word (.docx) without altering text, equation, or image content. Enforces left alignment (eliminates character stretching), 2-character Chinese first-line indentation, SimSun/Times New Roman fonts, professional table styling, figure/table caption centering, and OpenXML East Asian typography rules."
---

# Academic DOCX Formatter (学术与竞赛文档 Word 格式规范化 Skill)

本 Skill 用于对现有的 Word（`.docx`）文档进行纯格式层面的统一规范化与排版美化。
**核心原则：内容零改动（Zero Content Alteration）**。绝不增删篡改任何正文字词、数据指标、图表标题、公式符号或嵌入式图片，仅在排版样式、字体体系、对齐规则、首行缩进及底层 OpenXML 规范上进行系统性标准化重构。

---

## 适用场景与触发短语

当用户提出以下需求时激活本 Skill：
- **格式规范需求**：“统一word格式”、“调整word排版”、“按学术规范排版”、“把格式统一成标准格式”。
- **不改动内容限制**：“在不改变内容的基础之上”、“只改格式不要改字”、“保留全部原文字句”。
- **排版缺陷修复**：“段落字距拉得太开”、“英文/数字撑得字距很稀疏”、“两端对齐太丑了改左对齐”、“首行缩进不对齐”、“图片和图题没有居中”。
- **竞赛与答辩材料整改**：数媒竞赛、挑战杯、“互联网+”、毕业设计说明书、白皮书的最终交付排版。

> [!TIP]
> **与 `humanize-paper` 的边界分工**：
> - `humanize-paper`：专精于**文本内容与论证逻辑重构**（去 AI 腔、提炼科学动机、强化消融对比）。
> - `academic-docx-formatter`：专精于**版面表现与 Word 排版工业标准**（零改字、统一字体字号、消除拉伸、锁死缩进与页数）。

---

## 排版标准规范对照表

| 版面元素 | 规范要求 | 底层 OpenXML / Word 属性 |
| :--- | :--- | :--- |
| **纸张与页边距** | 标准 A4（21.0 cm × 29.7 cm），上/下 2.54 cm，左/右 2.70 cm（版芯有效宽 15.6 cm） | `top="1440" bottom="1440" left="1530" right="1530"` |
| **正文对齐** | **左对齐（Left Alignment）**，杜绝两端对齐（Justify）因西文/数字导致的文字强行撑开拉伸 | `w:jc w:val="left"` |
| **正文段落缩进** | 严格统一为**首行缩进 2 中文字符**（10.5pt 对应 21.0pt = 420 dxa） | `w:ind w:firstLine="420" w:firstLineChars="200"` |
| **正文字体绑定** | 中文：**宋体（SimSun）**；西文/数字：**Times New Roman**；字号：10.5 pt（五号） | `w:rFonts w:eastAsia="SimSun" w:ascii="Times New Roman"` |
| **正文行距间距** | 1.15 ~ 1.25 倍行距，段前 0 pt，段后 1.5 ~ 2.5 pt | `lineSpacing=1.20, space_after=Pt(2.0)` |
| **一级标题** | 13.5 ~ 14 pt（四号）黑体（SimHei）加粗，左对齐，首行缩进 0，段前 5~6pt，段后 2~3pt | `keep_with_next=True, firstLine="0"` |
| **二级标题** | 12 pt（小四）黑体（SimHei）加粗，左对齐，首行缩进 0，段前 3.5~4pt，段后 1.5~2pt | `keep_with_next=True, firstLine="0"` |
| **三级标题** | 10.5 pt（五号）黑体（SimHei）加粗，左对齐，首行缩进 0，段前 2.5pt，段后 1pt | `keep_with_next=True, firstLine="0"` |
| **插图与图题** | 图片与下方图题**水平居中**，图题 9.0 ~ 9.5 pt 黑体加粗，首行缩进 0 | `align=CENTER, firstLine="0"` |
| **表格与表题** | 表格居中；表题置于表格上方居中（9.5pt 黑体）；表头浅灰底色（`#D9D9D9`），细实线框 | `tblHeader`, `cantSplit`, `align=CENTER` |
| **标点与换行** | 开启东亚版式标点压缩（kinsoku）、行尾溢出（overflowPunct）、禁止 Shift+Enter 跨行撑开 | `compressPunctuation`, `doNotExpandShiftReturn` |

---

## 自动化工具执行流程

本 Skill 内置了高鲁棒性的自动化 Python 引擎，位于 `scripts/` 目录下：

```bash
# 1. 扫描与审计排版缺陷（输出问题段落清单）
python scripts/audit_docx.py input.docx -v

# 2. 一键无损标准化排版（保持全部内容不变）
python scripts/format_docx.py input.docx -o output_formatted.docx --audit

# 3. （可选）一键高质量转换为 PDF 并校验页数预算
python scripts/export_pdf.py output_formatted.docx
```

### Python API 内部集成

```python
from scripts.format_docx import format_document
from scripts.audit_docx import audit_docx

# 针对已有文档执行一键标准化
formatted_path = format_document(
    input_path="我的作品说明书.docx",
    output_path="我的作品说明书_规范终版.docx",
    body_line_spacing=1.20,
    table_header_bg="D9D9D9"
)

# 验证排版合规性
is_ok = audit_docx(formatted_path)
```

---

## 典型排版疑难与根治方案

### 1. 为什么坚决使用“左对齐”而不是“两端对齐”？
- **症结**：在 Word 默认两端对齐（`w:jc w:val="both"`）机制下，当行尾出现较长西文词组（如 `Edge Impulse EON Tuner`、`ESP32-C3`）、代码缩写或段内软回车（`\n`）时，Word 会强行拉伸中文字符与英文字母之间的留白，产生极其丑陋的巨大空白缝隙（例如：`C++          函          数          库          :`）。
- **治本**：将正文切换为 `w:jc w:val="left"`，配合 `compressPunctuation` 标点压缩，字间距自然致密，完全杜绝拉伸走样。

### 2. 避免使用段内软回车（Shift+Enter）
- **症结**：许多作者在同一个段落中使用 `\n` 进行序号换行（1. ... \n 2. ...），Word 会将其渲染为 `<w:br/>` 并视为行内折行，导致上一行末尾产生强制拉伸。
- **治本**：拆分成独立的 `<w:p>` 规范段落，加粗前缀采用黑体，内容采用宋体，各条目享有统一的段间距与缩进。

### 3. 表格跨页防碎裂
- 对表格所有行注入 `<w:cantSplit/>`，防止表格某一行被生硬切断在两页之中；首行注入 `<w:tblHeader/>`，实现跨页表头自动重复。
