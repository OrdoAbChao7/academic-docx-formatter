# Academic DOCX Formatter (学术与科技竞赛文档 Word 格式规范化工具)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-brightgreen.svg)](https://www.python.org/)
[![Antigravity Skill](https://img.shields.io/badge/Antigravity-Agent%20Skill-orange.svg)](#antigravity--agent-skill-集成)

> **让学术、毕业设计与竞赛申报 Word 文档一键实现出版级工业排版，承诺 100% 保持文字与图表内容零改动！**

---

## 🌟 核心设计原则 (Core Principles)

1. **内容零篡改 (Non-Destructive Guarantee)**：
   - 绝不修改、删除、重新生成任何正文字句、实验数据、公式符号或图片附件。
   - 仅对文档结构、版面尺寸、段落对齐、首行缩进、字体绑定及 OpenXML 底层样式进行工业级规范化。
2. **彻底解决排版恶疾 (Flawless Typography)**：
   - **告别字间距拉伸**：全面采用**规范左对齐（Left Alignment）**，彻底根除 Word 两端对齐（Justify）在遇到西文专业词汇、英文缩写或软换行时强行拉扯出的巨大空白缝隙（如 `C++          函          数          库          :`）。
   - **严格 2 字符首行缩进**：正文统一锁定 `w:firstLineChars="200"`（420 dxa = 21.0 pt），标题与图表严格锁定 `0` 缩进。
   - **中英双字型绑定**：中文统一为**宋体（SimSun）**，西文/数字统一为 **Times New Roman**，标题统一为**黑体（SimHei）**。
   - **插图与图表自动化居中**：图片及其图题（下方）、表格及其表题（上方）全部实现专业居中与防拆分（`cantSplit`）。
   - **底层东亚排版优化**：自动注入 `compressPunctuation`（标点压缩）、`kinsoku`（避头尾法则）以及 `doNotExpandShiftReturn`。

---

## 📋 排版规范一览表

| 元素分类 | 格式规范要求 | 对应 OpenXML / Word 属性 |
| :--- | :--- | :--- |
| **纸张版面** | A4 规格（21.0 × 29.7 cm），上下 2.54 cm，左右 2.70 cm（有效版芯 15.6 cm） | `w:pgSz w:w="11906" w:h="16838"` |
| **正文对齐** | **左对齐（Left）**（杜绝两端拉伸） | `w:jc w:val="left"` |
| **段落缩进** | 严格统一**首行缩进 2 中文字符** | `w:ind w:firstLine="420" w:firstLineChars="200"` |
| **字体字号** | 中文宋体，英文/数字 Times New Roman，五号（10.5 pt） | `w:rFonts w:eastAsia="SimSun" w:ascii="Times New Roman"` |
| **行距间距** | 1.15 ~ 1.25 倍行距，段前 0 pt，段后 1.5 ~ 2.0 pt | `lineSpacing=1.20, space_after=2.0pt` |
| **一级标题** | 黑体四号（13.5~14 pt）加粗，左对齐，首行无缩进，段前 5.5pt，段后 2.5pt | `keep_with_next=True, firstLine="0"` |
| **二级标题** | 黑体小四（12 pt）加粗，左对齐，首行无缩进，段前 4pt，段后 1.8pt | `keep_with_next=True, firstLine="0"` |
| **三级标题** | 黑体五号（10.5 pt）加粗，左对齐，首行无缩进，段前 2.5pt，段后 1.2pt | `keep_with_next=True, firstLine="0"` |
| **插图与图题** | 图像段落居中；下方图题居中（9.0~9.5 pt 黑体），首行无缩进 | `align=CENTER, firstLine="0"` |
| **表格与表题** | 表格上方表题居中（9.5 pt 黑体）；表头浅灰底色（`#D9D9D9`），表格居中且防跨页断裂 | `tblHeader`, `cantSplit`, `align=CENTER` |

---

## 🚀 快速上手 (Quick Start)

### 1. 环境准备
```bash
git clone https://github.com/OrdoAbChao7/academic-docx-formatter.git
cd academic-docx-formatter
pip install -r requirements.txt
```

### 2. 命令行一键格式化 (CLI)
```bash
# 1. 扫描与审计排版缺陷
python scripts/audit_docx.py input.docx -v

# 2. 一键无损标准化排版
python scripts/format_docx.py input.docx -o output_formatted.docx --audit

# 3. 转换为高质量 PDF 并核实页数预算 (支持 Word COM 原生渲染)
python scripts/export_pdf.py output_formatted.docx
```

### 3. 运行内置演示示例
```bash
python examples/demo_format.py
```

---

## 🤖 Antigravity / AI Agent Skill 集成

本项目完整支持 Google Antigravity、Claude Code 及 Cursor 智能体 Skill 规范。

### 安装方法：
将本项目放置于智能体 skills 目录下（例如 `~/.gemini/config/skills/academic-docx-formatter`），智能体即可自动识别并在下列对话中激活：

- *"请帮我统一这个 Word 文档的格式，不要改动任何文字内容。"*
- *"修复文档里的字间距拉伸和缩进问题，按规范学术格式排版。"*
- *"把竞赛申报书排版成标准的五号宋体、左对齐、首行缩进 2 字符。"*

> **提示**：配合 [`humanize-paper`](https://github.com/OrdoAbChao7) 可形成“**内容科学论证润色 + 排版工业级规范化**”的双重完整流水线！

---

## 🛠️ 项目结构 (Repository Structure)

```
academic-docx-formatter/
├── SKILL.md                  # Antigravity 智能体 Skill 核心规则描述
├── README.md                 # 中英文使用文档
├── LICENSE                   # MIT 开源协议
├── requirements.txt          # Python 依赖项
├── scripts/
│   ├── format_docx.py        # 核心格式化引擎 (无损排版与样式统一)
│   ├── audit_docx.py         # 格式合规性检查与排版缺陷审计器
│   ├── export_pdf.py         # Word COM / docx2pdf 高保真导出与页数校验
│   └── typography_patcher.py # OpenXML 东亚排版底层参数补丁注入器
└── examples/
    └── demo_format.py        # 10秒快速上手演示脚本
```

---

## 📄 开源许可证

本项目基于 [MIT License](LICENSE) 开源发布，欢迎自由商用、集成与二次开发。
