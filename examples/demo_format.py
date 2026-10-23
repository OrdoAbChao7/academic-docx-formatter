# -*- coding: utf-8 -*-
"""
demo_format.py
Quick start demonstration of academic-docx-formatter.
Creates an unformatted sample document, audits it, standardizes it, and audits again.
"""

import os
import sys
import docx
from docx.shared import Pt, Inches

# Add scripts directory to path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'scripts')))
from format_docx import format_document
from audit_docx import audit_docx

def create_sample_unformatted_docx(filename="sample_unformatted.docx"):
    doc = docx.Document()
    
    # Title
    p_title = doc.add_paragraph("全国大学生科技竞赛作品技术说明书")
    
    # Subtitle
    doc.add_paragraph("（参赛团队匿名盲审申报材料）")
    
    # Heading 1
    doc.add_paragraph("一、 项目研发背景与核心创新点")
    
    # Unformatted Body Paragraph (justified, no 2-char indent, mixed fonts, line breaks)
    p_body1 = doc.add_paragraph(
        "本项目针对现有可穿戴助残设备体积笨重、起居隐私泄露风险高以及边缘端算力受限等痛点，"
        "提出了一套基于近红外漫反射光电微动作感知与端侧 TinyML 极速脱网推理的无声交互系统。"
        "系统采用 3 枚微距传感器捕捉口唇细微位移，在 ESP32 芯片上实现了 1.0 ms 极速推理。"
    )
    p_body1.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY # Causes stretching bug
    
    # Heading 2
    doc.add_paragraph("1.1 关键技术指标与算力权衡")
    
    # Unformatted table
    table = doc.add_table(rows=3, cols=3)
    data = [
        ["指标项", "实测参数", "国家标准依据"],
        ["推理延迟", "1.0 ms", "毫秒级实时响应规范"],
        ["SRAM 占用", "6.7 KB", "微控制器轻量部署标准"]
    ]
    for r_i, row in enumerate(data):
        for c_i, val in enumerate(row):
            table.rows[r_i].cells[c_i].paragraphs[0].text = val
            
    doc.save(filename)
    print(f"Created sample unformatted document: {filename}")
    return filename

if __name__ == "__main__":
    sample_file = "sample_unformatted.docx"
    formatted_file = "sample_formatted.docx"
    
    create_sample_unformatted_docx(sample_file)
    
    print("\n>>> AUDITING UNFORMATTED DOCUMENT:")
    audit_docx(sample_file, verbose=True)
    
    print("\n>>> FORMATTING DOCUMENT:")
    format_document(sample_file, output_path=formatted_file)
    
    print("\n>>> AUDITING FORMATTED DOCUMENT:")
    audit_docx(formatted_file, verbose=True)
