# -*- coding: utf-8 -*-
"""
format_docx.py
Comprehensive Academic and Technical Word Document (.docx) Formatter.

Core Principles:
- 100% NON-DESTRUCTIVE: Preserves all textual sentences, numbers, formulas, drawings, and images.
- DUAL PRESETS:
  * 'comprehensive' (详本方案): Body 12.0 pt (小四), Headings 16/14/12 pt, Line spacing 1.25.
  * 'compact' (紧凑申报方案): Body 10.5 pt (五号), Headings 13.5/12/10.5 pt, Line spacing 1.15~1.20.
- TYPOGRAPHIC DISCIPLINE:
  1. LEFT Alignment (eliminates ugly text-stretching from Latin/numbers/soft-returns).
  2. Strict Chinese 2-character indentation (firstLineChars="200").
  3. Dual font system: SimSun (宋体) for CJK + Times New Roman for Latin/Digits.
  4. Headings: SimHei (黑体) bold, left-aligned, keep_with_next=True.
  5. Captions: SimHei bold, centered, figure below, table above.
  6. Table Anti-fragmentation: Center alignment, tblHeader repeat, cantSplit rows, #D9D9D9 shading.
  7. Side-by-side Images: Centered, borderless 1x2 table formatting.
  8. OpenXML East Asian typography rules (compressPunctuation, kinsoku, doNotExpandShiftReturn).
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
# Presets Configuration
# ---------------------------------------------------------------------------
PRESETS = {
    "comprehensive": {
        "name": "工程详本 / 学术论文版 (12pt 小四)",
        "body_size_pt": 12.0,
        "body_line_spacing": 1.25,
        "body_space_after_pt": 3.0,
        "h1_size_pt": 16.0,
        "h1_before_pt": 15.0,
        "h1_after_pt": 7.0,
        "h2_size_pt": 14.0,
        "h2_before_pt": 9.0,
        "h2_after_pt": 4.5,
        "h3_size_pt": 12.0,
        "h3_before_pt": 5.0,
        "h3_after_pt": 2.5,
        "caption_size_pt": 10.0,
        "table_body_size_pt": 9.5,
        "table_pad_v": 45,
        "table_pad_h": 75,
    },
    "compact": {
        "name": "竞赛紧凑 / 限页申报版 (10.5pt 五号)",
        "body_size_pt": 10.5,
        "body_line_spacing": 1.18,
        "body_space_after_pt": 2.0,
        "h1_size_pt": 13.5,
        "h1_before_pt": 5.5,
        "h1_after_pt": 2.5,
        "h2_size_pt": 12.0,
        "h2_before_pt": 4.0,
        "h2_after_pt": 1.8,
        "h3_size_pt": 10.5,
        "h3_before_pt": 2.5,
        "h3_after_pt": 1.2,
        "caption_size_pt": 9.0,
        "table_body_size_pt": 8.0,
        "table_pad_v": 25,
        "table_pad_h": 50,
    }
}

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

def set_cell_borders(cell, sz="4", color="000000"):
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

def remove_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/><w:left w:val="none"/>'
        f'<w:bottom w:val="none"/><w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

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

    # 2. Math Formula check
    has_math = bool(p._p.xpath('.//m:oMath') or p._p.xpath('.//m:oMathPara'))
    if has_math:
        return "math_formula"

    # 3. Document Title / Front-matter check
    if idx < 5 and len(txt) < 60 and any(k in txt for k in ['竞赛', '大赛', '说明书', '报告', '论文', '系统']):
        if idx == 0 and len(txt) > 25:
            return "header_banner"
        if any(k in txt for k in ['说明书', '报告', '设计', '论文']):
            return "doc_title"
        if '（' in txt and '）' in txt:
            return "doc_subtitle"

    # 4. Caption checks
    if re.match(r'^图\s*[0-9一二三四五六七八九十]+[：:、\s]', txt) or re.match(r'^Figure\s*[0-9]+', txt, re.I):
        return "figure_caption"
    if re.match(r'^表\s*[0-9一二三四五六七八九十]+[：:、\s]', txt) or re.match(r'^Table\s*[0-9]+', txt, re.I):
        return "table_caption"

    # 5. Heading Level 1 check
    if re.match(r'^[一二三四五六七八九十]+[、\s]', txt) or re.match(r'^第[一二三四五六七八九十0-9]+[章部分][\s：:]', txt):
        return "heading_1"
    if re.match(r'^[1-9]\s+[^\d]', txt) or re.match(r'^[1-9]\.[^\d]', txt):
        return "heading_1"

    # 6. Heading Level 2 check
    if re.match(r'^[1-9]\.[0-9]+[、\s\.]', txt) or re.match(r'^[（(][一二三四五六七八九十0-9]+[）)]', txt):
        return "heading_2"

    # 7. Heading Level 3 check
    if re.match(r'^[1-9]\.[0-9]+\.[0-9]+[、\s\.]', txt) or re.match(r'^[1-9]\)[、\s]', txt):
        return "heading_3"

    # 8. Metadata / Callout check
    if txt.startswith('【') and ('】' in txt[:30]):
        return "metadata_callout"

    # 9. Otherwise standard body paragraph
    return "body"

# ---------------------------------------------------------------------------
# Formatting Routine
# ---------------------------------------------------------------------------
def format_document(
    input_path,
    output_path=None,
    preset="compact",
    font_ea="SimSun",
    font_ea_heading="SimHei",
    font_ascii="Times New Roman",
    align_body="left",
    table_header_bg="D9D9D9",
    split_soft_breaks=True
):
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_formatted{ext}"

    cfg = PRESETS.get(preset, PRESETS["compact"])
    print(f"\n=================================================================")
    print(f"  ACADEMIC DOCX FORMATTER: {cfg['name']}")
    print(f"=================================================================")
    print(f"[1/4] Loading document: {input_path}")
    doc = docx.Document(input_path)

    # Standard A4 Page Setup (15.6 cm effective width)
    for s in doc.sections:
        s.page_width = Pt(595.3)   # 21.0 cm
        s.page_height = Pt(841.9)  # 29.7 cm
        s.top_margin = Pt(72.0)    # 2.54 cm
        s.bottom_margin = Pt(72.0) # 2.54 cm
        s.left_margin = Pt(79.2)   # 2.79 cm
        s.right_margin = Pt(73.5)  # 2.59 cm
        s.header_distance = Pt(42.5) # 1.5 cm
        s.footer_distance = Pt(42.5) # 1.5 cm

    total_p = len(doc.paragraphs)
    print(f"[2/4] Formatting {total_p} paragraphs (align={align_body}, body_size={cfg['body_size_pt']}pt)...")

    align_choice = WD_ALIGN_PARAGRAPH.LEFT if align_body == "left" else WD_ALIGN_PARAGRAPH.JUSTIFY

    for idx, p in enumerate(doc.paragraphs):
        p_type = classify_paragraph(p, idx, total_p)
        txt = p.text.strip()
        pf = p.paragraph_format
        pPr = p._p.get_or_add_pPr()

        # Clean existing indents and alignments
        for old_jc in pPr.findall(qn('w:jc')): pPr.remove(old_jc)
        for old_ind in pPr.findall(qn('w:ind')): pPr.remove(old_ind)

        if p_type == "empty":
            continue

        elif p_type == "image":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(4.0)
            pf.space_after = Pt(2.0)
            pf.keep_with_next = True

        elif p_type == "math_formula":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(3.0)
            pf.space_after = Pt(3.0)
            pf.line_spacing = 1.25 # Prevent math symbol clipping

        elif p_type == "doc_title":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(8.0)
            pf.space_after = Pt(4.0)
            pf.line_spacing = 1.15
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=16.0, bold=True)

        elif p_type in ("header_banner", "doc_subtitle"):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(2.0)
            pf.space_after = Pt(2.5)
            pf.line_spacing = 1.10
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading if p_type=="header_banner" else font_ea, 
                              font_ascii=font_ascii, size_pt=10.5 if p_type=="header_banner" else 9.5, 
                              bold=(p_type=="header_banner"))

        elif p_type == "figure_caption":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(2.0)
            pf.space_after = Pt(5.0)
            pf.line_spacing = 1.10
            pf.keep_with_next = False
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=cfg["caption_size_pt"], bold=True)

        elif p_type == "table_caption":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(5.0)
            pf.space_after = Pt(2.0)
            pf.line_spacing = 1.10
            pf.keep_with_next = True
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=cfg["caption_size_pt"], bold=True)

        elif p_type == "heading_1":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(cfg["h1_before_pt"])
            pf.space_after = Pt(cfg["h1_after_pt"])
            pf.line_spacing = 1.15
            pf.keep_with_next = True
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=cfg["h1_size_pt"], bold=True)

        elif p_type == "heading_2":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(cfg["h2_before_pt"])
            pf.space_after = Pt(cfg["h2_after_pt"])
            pf.line_spacing = 1.12
            pf.keep_with_next = True
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=cfg["h2_size_pt"], bold=True)

        elif p_type == "heading_3":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            pf.space_before = Pt(cfg["h3_before_pt"])
            pf.space_after = Pt(cfg["h3_after_pt"])
            pf.line_spacing = 1.10
            pf.keep_with_next = True
            for r in p.runs:
                set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, size_pt=cfg["h3_size_pt"], bold=True)

        elif p_type == "metadata_callout":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="420" w:firstLineChars="200"/>'))
            pf.space_before = Pt(1.5)
            pf.space_after = Pt(2.5)
            pf.line_spacing = 1.15
            for r in p.runs:
                is_bold = r.bold or ('【' in r.text and '】' in r.text)
                set_run_fonts(r, font_ea=font_ea_heading if is_bold else font_ea,
                              font_ascii=font_ascii, size_pt=cfg["body_size_pt"]-1.0, bold=is_bold)

        else: # Standard Body Paragraph
            p.alignment = align_choice
            jc_val = "left" if align_choice == WD_ALIGN_PARAGRAPH.LEFT else "both"
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{jc_val}"/>'))
            
            # Enforce 2 Chinese characters first-line indent
            indent_dxa = int(cfg["body_size_pt"] * 2 * 20)
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="{indent_dxa}" w:firstLineChars="200"/>'))
            
            pf.space_before = Pt(0)
            pf.space_after = Pt(cfg["body_space_after_pt"])
            pf.line_spacing = cfg["body_line_spacing"]
            pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

            for r in p.runs:
                r_bold = r.bold
                set_run_fonts(r, font_ea=font_ea_heading if r_bold else font_ea,
                              font_ascii=font_ascii, size_pt=cfg["body_size_pt"], bold=r_bold)

    # Format Tables (and Side-by-side Image tables)
    print(f"[3/4] Formatting {len(doc.tables)} tables (anti-fragmentation & cell styling)...")
    for tbl in doc.tables:
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        # Detect if table is a 1x2 image-pairing table (side-by-side images)
        is_side_by_side_images = (len(tbl.rows) == 1 and len(tbl.columns) == 2 and 
                                  any(bool(c._tc.xpath('.//w:drawing') or c._tc.xpath('.//w:pict')) for c in tbl.rows[0].cells))

        if is_side_by_side_images:
            # Borderless centered side-by-side layout
            for c in tbl.rows[0].cells:
                remove_cell_borders(c)
                set_cell_margins(c, top=10, bottom=10, left=30, right=30)
                for cp in c.paragraphs:
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cpPr = cp._p.get_or_add_pPr()
                    for old_ind in cpPr.findall(qn('w:ind')): cpPr.remove(old_ind)
                    cpPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
            continue

        # Academic / Technical Data Table
        if len(tbl.rows) > 0:
            header_tr = tbl.rows[0]._tr.get_or_add_trPr()
            header_tr.append(OxmlElement('w:tblHeader'))
            header_tr.append(OxmlElement('w:cantSplit'))

            for cell in tbl.rows[0].cells:
                set_cell_borders(cell, sz="4", color="000000")
                set_cell_shading(cell, color_hex=table_header_bg)
                set_cell_margins(cell, top=cfg["table_pad_v"]+5, bottom=cfg["table_pad_v"]+5, 
                                 left=cfg["table_pad_h"], right=cfg["table_pad_h"])
                for cp in cell.paragraphs:
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cpPr = cp._p.get_or_add_pPr()
                    for old_ind in cpPr.findall(qn('w:ind')): cpPr.remove(old_ind)
                    cpPr.append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0" w:firstLineChars="0"/>'))
                    cp.paragraph_format.space_before = Pt(1.0)
                    cp.paragraph_format.space_after = Pt(1.0)
                    cp.paragraph_format.line_spacing = 1.05
                    for r in cp.runs:
                        set_run_fonts(r, font_ea=font_ea_heading, font_ascii=font_ascii, 
                                      size_pt=cfg["table_body_size_pt"]+0.5, bold=True)

        for r_idx in range(1, len(tbl.rows)):
            row = tbl.rows[r_idx]
            row_tr = row._tr.get_or_add_trPr()
            row_tr.append(OxmlElement('w:cantSplit'))
            for cell in row.cells:
                set_cell_borders(cell, sz="4", color="000000")
                set_cell_margins(cell, top=cfg["table_pad_v"], bottom=cfg["table_pad_v"], 
                                 left=cfg["table_pad_h"], right=cfg["table_pad_h"])
                for cp in cell.paragraphs:
                    cp_txt = cp.text.strip()
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
                                      font_ascii=font_ascii, size_pt=cfg["table_body_size_pt"], bold=r_bold)

    # Save output docx and apply low-level OpenXML typography patch
    print(f"[4/4] Saving and patching OpenXML settings: {output_path}")
    doc.save(output_path)
    patch_docx_typography(output_path)
    print(f"[SUCCESS] Formatting complete: {output_path}")

    return output_path

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Academic and Technical Word Document (.docx) Formatter")
    parser.add_argument("input", help="Path to input .docx document")
    parser.add_argument("-o", "--output", help="Path to formatted output .docx document (optional)")
    parser.add_argument("--preset", choices=["compact", "comprehensive"], default="compact", 
                        help="Typography preset: 'compact' (10.5pt 五号, 竞赛限页) or 'comprehensive' (12pt 小四, 论文详本)")
    parser.add_argument("--align", choices=["left", "justify"], default="left", help="Paragraph alignment (default: left)")
    parser.add_argument("--audit", action="store_true", help="Run audit after formatting")
    args = parser.parse_args()

    out = format_document(
        args.input,
        output_path=args.output,
        preset=args.preset,
        align_body=args.align
    )

    if args.audit:
        audit_docx(out, verbose=True)
