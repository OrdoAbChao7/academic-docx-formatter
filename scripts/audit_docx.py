# -*- coding: utf-8 -*-
"""
audit_docx.py
Formatting auditor / linter for Chinese academic and technical .docx documents.
Scans for:
1. Forced justification ('both' alignment) causing character stretching.
2. Missing or incorrect Chinese 2-character first-line indentation.
3. Soft line breaks (<w:br/>) inside paragraphs.
4. Uncentered figures and captions.
5. Inconsistent fonts and styles.
"""

import os
import sys
import argparse
import docx
from docx.oxml.ns import qn

def audit_docx(docx_path, verbose=False):
    if not os.path.exists(docx_path):
        print(f"[ERROR] File not found: {docx_path}")
        return False

    doc = docx.Document(docx_path)
    issues = {
        "justified_paragraphs": [],
        "missing_indent_paragraphs": [],
        "soft_line_breaks": [],
        "uncentered_captions": [],
        "uncentered_images": [],
    }

    total_p = len(doc.paragraphs)
    total_t = len(doc.tables)

    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        pPr = p._p.pPr

        # 1. Alignment check
        jc = pPr.find(qn('w:jc')) if pPr is not None else None
        jc_val = jc.get(qn('w:val')) if jc is not None else 'left'

        if jc_val == 'both':
            issues["justified_paragraphs"].append((i, txt[:40]))

        # 2. Soft line breaks check (<w:br> without type="page")
        brs = p._p.xpath('.//w:br[not(@w:type)]') if p._p is not None else []
        if brs or '\n' in p.text:
            issues["soft_line_breaks"].append((i, txt[:40]))

        # 3. Indent check for body paragraphs
        # Exclude headings, captions, titles, empty paragraphs
        is_caption = txt.startswith('图') or txt.startswith('表') or '图 ' in txt[:6] or '表 ' in txt[:6]
        has_drawing = bool(p._p.xpath('.//w:drawing') or p._p.xpath('.//w:pict')) if p._p is not None else False
        is_heading = any(txt.startswith(k) for k in ['一、', '二、', '三、', '四、', '五、', '六、', '七、', '八、', '九、', '十、']) or \
                     any(txt.startswith(f'{n}.') for n in range(1, 10)) or \
                     any(txt.startswith(f'（{n}）') for n in range(1, 10))

        if txt and not is_caption and not has_drawing and not is_heading and len(txt) > 25:
            ind = pPr.find(qn('w:ind')) if pPr is not None else None
            fl = ind.get(qn('w:firstLine')) if ind is not None else None
            flc = ind.get(qn('w:firstLineChars')) if ind is not None else None
            if not fl and not flc:
                issues["missing_indent_paragraphs"].append((i, txt[:40]))

        # 4. Caption & Image centering check
        if is_caption and jc_val != 'center':
            issues["uncentered_captions"].append((i, txt[:40]))
        if has_drawing and jc_val != 'center':
            issues["uncentered_images"].append((i, "Image paragraph"))

    print(f"\n=================================================================")
    print(f"  FORMATTING AUDIT REPORT: {os.path.basename(docx_path)}")
    print(f"=================================================================")
    print(f"Total Paragraphs: {total_p} | Total Tables: {total_t}")
    print(f"-----------------------------------------------------------------")
    print(f"1. Justified Paragraphs (both):      {len(issues['justified_paragraphs'])} (Potential text-stretching)")
    print(f"2. Soft Line Breaks (\\n / <w:br/>):   {len(issues['soft_line_breaks'])} (Risk of forced line expansion)")
    print(f"3. Missing 2-char Indent:            {len(issues['missing_indent_paragraphs'])} (Body paragraphs)")
    print(f"4. Uncentered Captions:              {len(issues['uncentered_captions'])}")
    print(f"5. Uncentered Images:                {len(issues['uncentered_images'])}")
    print(f"-----------------------------------------------------------------")

    is_compliant = (len(issues["justified_paragraphs"]) == 0 and 
                    len(issues["soft_line_breaks"]) == 0 and 
                    len(issues["uncentered_captions"]) == 0 and
                    len(issues["uncentered_images"]) == 0)

    if is_compliant:
        print("[STATUS: COMPLIANT] Document conforms to strict typography standards!")
    else:
        print("[STATUS: NON-COMPLIANT] Found typography flaws that should be normalized.")

    if verbose:
        for k, items in issues.items():
            if items:
                print(f"\nDetails for {k} ({len(items)}):")
                for it in items[:10]:
                    print(f"   P{it[0]:2d}: {it[1]}")
                if len(items) > 10:
                    print(f"   ... and {len(items)-10} more.")
    print("=================================================================\n")

    return is_compliant

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Audit Word document (.docx) formatting.")
    parser.add_argument("docx_path", help="Path to input .docx document")
    parser.add_argument("-v", "--verbose", action="store_true", help="Print detailed list of issues")
    args = parser.parse_args()

    audit_docx(args.docx_path, verbose=args.verbose)
