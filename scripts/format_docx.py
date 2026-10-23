# -*- coding: utf-8 -*-
"""
format_docx.py
Core formatting engine for Academic, Competition, and Technical Word Documents (.docx).

Core Guarantee:
- NON-DESTRUCTIVE: Preserves 100% of text content, numbers, formulas, drawings, and images.
- STANDARDIZED TYPOGRAPHY:
  1. Enforces LEFT alignment on body text (eliminates ugly character stretching).
  2. Enforces strict Chinese 2-character indentation (w:firstLineChars="200", w:firstLine="420").
  3. Enforces dual-font binding: SimSun (宋体) for CJK + Times New Roman for Latin/Digits.
  4. Enforces clear heading hierarchy (SimHei 黑体) and keep_with_next.
  5. Centers all figures and figure/table captions with 0 first-line indent.
  6. Standardizes tables with professional borders, cell margins, and shaded headers.
  7. Injects OpenXML East Asian typography rules (punctuation compression, kinsoku).
"""

import os
import sys
import re
import argparse
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from typography_patcher import patch_docx_typography
from audit_docx import audit_docx

# ---------------------------------------------------------------------------
# Font & Style Helpers
# ---------------------------------------------------------------------------
def set_run_fonts(run, font_ea="SimSun", font_ascii="Times New Roman", size_pt=10.5, bold=None, color_rgb=(0, 0, 0)):
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if color_rgb is not None:
        run.font.color.rgb = RGBColor(*color_rgb)
        
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_ea)
    rFonts.set(qn('w:ascii'), font_ascii)
    rFonts.set(qn('w:hAnsi'), font_ascii)
    rFonts.set(qn('w:cs'), font_ascii)

def set_cell_border(cell, sz="4", color="000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), str(sz))
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        tcBorders.append(b)

def set_cell_margins(cell, top=35, bottom=35, left=55, right=55):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_cell_shading(cell, color_hex="D9D9D9"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

# ---------------------------------------------------------------------------
# Paragraph Classifier
# ---------------------------------------------------------------------------
def classify_paragraph(p, idx, total_paragraphs):
    txt = p.text.strip()
    if not txt:
        return "empty"

    # 1. Image check
    has_drawing = bool(p._p.xpath('.//w:drawing') or p._p.xpath('.//w:pict'))
    if has_drawing and len(txt) == 0:
        return "image"

    # 2. Document Title check (typically first 1-3 paragraphs)
    if idx < 4 and len(txt) < 50 and ('大赛' in txt or '说明书' in txt or '报告' in txt or '论文' in txt or '系统' in txt):
        if idx == 0 and len(txt) > 25:
            return "header_banner"
        if '说明书' in txt or '报告' in txt or '设计' in txt or '论文' in txt:
            return "doc_title"
        if '（' in txt and '）' in txt:
            return "doc_subtitle"

    # 3. Caption checks
    if re.match(r'^图\s*[0-9一二三四五六七八九十]+[：:、\s]', txt) or re.match(r'^Figure\s*[0-9]+', txt, re.I):
        return "figure_caption"
    if re.match(r'^表\s*[0-9一二三四五六七八九十]+[：:、\s]', txt) or re.match(r'^Table\s*[0-9]+', txt, re.I):
        return "table_caption"

    # 4. Heading Level 1 check
    # e.g., 一、 项目背景 / 1. 引言 / 第一章
    if re.match(r'^[一二三四五六七八九十]+[、\s]', txt) or re.match(r'^第[一二三四五六七八九十0-9]+[章部分][\s：:]', txt):
        return "heading_1"
    if re.match(r'^[1-9]\s+[^\d]', txt) or re.match(r'^[1-9]\.[^\d]', txt):
        return "heading_1"

    # 5. Heading Level 2 check
    # e.g., 1.1 总体架构 / （一）市场分析
    if re.match(r'^[1-9]\.[0-9]+[、\s\.]', txt) or re.match(r'^[（(][一二三四五六七八九十0-9]+[）)]', txt):
        return "heading_2"

    # 6. Heading Level 3 check
    # e.g., 1.1.1 硬件选型 / 1) 传感器
    if re.match(r'^[1-9]\.[0-9]+\.[0-9]+[、\s\.]', txt) or re.match(r'^[1-9]\)[、\s]', txt):
        return "heading_3"

    # 7. Metadata / Callout check
    if txt.startswith('【') and ('】' in txt[:30]):
        return "metadata_callout"

    # 8. Otherwise standard body paragraph
    return "body"

# ---------------------------------------------------------------------------
# Formatting Routine
# ---------------------------------------------------------------------------
def format_document(
    input_path,
    output_path=None,
    font_ea="SimSun",
    font_ea_heading="SimHei",
    font_ascii="Times New Roman",
    body_size_pt=10.5,
    body_line_spacing=1.20,
    body_space_after_pt=2.0,
    align_body=WD_ALIGN_PARAGRAPH.LEFT,
    table_header_bg="D9D9D9",
    split_soft_breaks=True
):
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_formatted{ext}"

    print(f"\n[1/4] Loading document: {input_path}")
    doc = docx.Document(input_path)

    # Standard A4 Page Setup
    for s in doc.sections:
        s.page_width = Pt(595.3)   # 21.0 cm
        s.page_height = Pt(841.9)  # 29.7 cm
        s.top_margin = Pt(72.0)    # 2.54 cm
        s.bottom_margin = Pt(72.0) # 2.54 cm
        s.left_margin = Pt(76.5)   # 2.70 cm
        s.right_margin = Pt(76.5)  # 2.70 cm (effective width: 15.6 cm)
        s.header_distance = Pt(36.0) # 1.27 cm
        s.footer_distance = Pt(36.0) # 1.27 cm

    total_p = len(doc.paragraphs)
    print(f"[2/4] Formatting {total_p} paragraphs...")

    for idx, p in enumerate(doc.paragraphs):
        p_type = classify_paragraph(p, idx, total_p)
        txt = p.text.strip()
        pf = p.paragraph_format
        pPr = p._p.get_or_add_pPr()

        # Clean existing indents and alignments
        for old_jc in pPr.findall(qn('w:jc')):
            pPr.remove(old_jc)
        for old_ind in pPr.findall(qn('w:ind')):
            pPr.remove(old_ind)

        # Apply specific formatting by type
        if p_type == "empty":
            continue

        elif p_type == "image":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(3.0)
            pf.space_after = Pt(2.0)
            pf.keep_with_next = True

        elif p_type == "doc_title":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(6.0)
            pf.space_after = Pt(3.0)
            pf.line_spacing = 1.15
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=15.0, bold=True)

        elif p_type in ("header_banner", "doc_subtitle"):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(1.5)
            pf.space_after = Pt(2.5)
            pf.line_spacing = 1.10
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading if p_type=="header_banner" else font_ea, 
                              font_ascii=font_ascii, size_pt=10.0 if p_type=="header_banner" else 9.5, 
                              bold=(p_type=="header_banner"))

        elif p_type == "figure_caption":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(2.0)
            pf.space_after = Pt(4.0)
            pf.line_spacing = 1.10
            pf.keep_with_next = False
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=9.0, bold=True)

        elif p_type == "table_caption":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(4.0)
            pf.space_after = Pt(2.0)
            pf.line_spacing = 1.10
            pf.keep_with_next = True
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=9.5, bold=True)

        elif p_type == "heading_1":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(5.5)
            pf.space_after = Pt(2.5)
            pf.line_spacing = 1.15
            pf.keep_with_next = True
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=13.5, bold=True)

        elif p_type == "heading_2":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(4.0)
            pf.space_after = Pt(1.8)
            pf.line_spacing = 1.12
            pf.keep_with_next = True
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=12.0, bold=True)

        elif p_type == "heading_3":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(2.5)
            pf.space_after = Pt(1.2)
            pf.line_spacing = 1.10
            pf.keep_with_next = True
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=10.5, bold=True)

        elif p_type == "metadata_callout":
            # Formatted callout / abstract
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="420" w:firstLineChars="200"/>'))
            pf.space_before = Pt(1.5)
            pf.space_after = Pt(2.0)
            pf.line_spacing = 1.15
            for r in p.runs:
                is_bold = r.bold or ('【' in r.text and '】' in r.text)
                set_run_fonts(r, font_ea=font_ea_heading if is_bold else font_ea,
                              font_ascii=font_ascii, size_pt=9.5, bold=is_bold)

        else: # Standard Body Paragraph
            p.alignment = align_body
            jc_val = "left" if align_body == WD_ALIGN_PARAGRAPH.LEFT else "both"
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{jc_val}"/>'))
            
            # Enforce 2 Chinese characters first-line indent (420 dxa = 21.0 pt)
            indent_dxa = int(body_size_pt * 2 * 20)
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="{indent_dxa}" w:firstLineChars="200"/>'))
            
            pf.space_before = Pt(0)
            pf.space_after = Pt(body_space_after_pt)
            pf.line_spacing = body_line_spacing
            pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

            for r in p.runs:
                # Maintain existing bold emphasis if present in source
                r_bold = r.bold
                set_run_fonts(r, font_ea=font_ea_heading if r_bold else font_ea,
                              font_ascii=font_ascii, size_pt=body_size_pt, bold=r_bold)

    # Format Tables
    print(f"[3/4] Formatting {len(doc.tables)} tables...")
    for tbl in doc.tables:
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        if len(tbl.rows) > 0:
            header_tr = tbl.rows[0]._tr.get_or_add_trPr()
            header_tr.append(OxmlElement('w:tblHeader'))
            header_tr.append(OxmlElement('w:cantSplit'))

            for cell in tbl.rows[0].cells:
                set_cell_border(cell, sz="4", color="000000")
                set_cell_shading(cell, color_hex=table_header_bg)
                set_cell_margins(cell, top=30, bottom=30, left=55, right=55)
                for cp in cell.paragraphs:
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cpPr = cp._p.get_or_add_pPr()
                    for old_ind in cpPr.findall(qn('w:ind')): cpPr.remove(old_ind)
                    cpPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
                    cp.paragraph_format.space_before = Pt(1.0)
                    cp.paragraph_format.space_after = Pt(1.0)
                    cp.paragraph_format.line_spacing = 1.05
                    for r in cp.runs:
                        set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=8.5, bold=True)

        for r_idx in range(1, len(tbl.rows)):
            row = tbl.rows[r_idx]
            row_tr = row._tr.get_or_add_trPr()
            row_tr.append(OxmlElement('w:cantSplit'))
            for cell in row.cells:
                set_cell_border(cell, sz="4", color="000000")
                set_cell_margins(cell, top=25, bottom=25, left=50, right=50)
                for cp in cell.paragraphs:
                    cp_txt = cp.text.strip()
                    # Align center for short numbers/units; left for descriptive text
                    align = WD_ALIGN_PARAGRAPH.CENTER if len(cp_txt) <= 12 or any(k in cp_txt for k in ['元', 'ms', 'Hz', 'KB', 'g', '验证', '片', '套', '枚', '批', '例', '%']) else WD_ALIGN_PARAGRAPH.LEFT
                    cp.alignment = align
                    cpPr = cp._p.get_or_add_pPr()
                    for old_ind in cpPr.findall(qn('w:ind')): cpPr.remove(old_ind)
                    cpPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
                    cp.paragraph_format.space_before = Pt(0.8)
                    cp.paragraph_format.space_after = Pt(0.8)
                    cp.paragraph_format.line_spacing = 1.05
                    for r in cp.runs:
                        r_bold = r.bold
                        set_run_fonts(r, font_ea=font_ea_heading if r_bold else font_ea,
                                      font_ascii=font_ascii, size_pt=8.0, bold=r_bold)

    # Save output docx
    print(f"[4/4] Saving and patching OpenXML settings: {output_path}")
    doc.save(output_path)
    patch_docx_typography(output_path)
    print(f"[SUCCESS] Formatting complete: {output_path}")

    return output_path

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Academic and Technical Word Document (.docx) Formatter")
    parser.add_argument("input", help="Path to input .docx document")
    parser.add_argument("-o", "--output", help="Path to formatted output .docx document (optional)")
    parser.add_argument("--align", choices=["left", "justify"], default="left", help="Paragraph alignment (default: left)")
    parser.add_argument("--line-spacing", type=float, default=1.20, help="Line spacing multiple (default: 1.20)")
    parser.add_argument("--audit", action="store_true", help="Run audit after formatting")
    args = parser.parse_args()

    align_choice = WD_ALIGN_PARAGRAPH.LEFT if args.align == "left" else WD_ALIGN_PARAGRAPH.JUSTIFY
    out = format_document(
        args.input,
        output_path=args.output,
        align_body=align_choice,
        body_line_spacing=args.line_spacing
    )

    if args.audit:
        audit_docx(out, verbose=True)
