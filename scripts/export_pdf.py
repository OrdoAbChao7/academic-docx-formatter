# -*- coding: utf-8 -*-
"""
export_pdf.py
Convert standardized .docx document to .pdf and verify page count.
Uses native Word COM on Windows (highest fidelity matching Office Word rendering)
with automatic fallback to docx2pdf / libreoffice.
"""

import os
import sys
import argparse

def convert_to_pdf(docx_path, pdf_path=None):
    if pdf_path is None:
        base, _ = os.path.splitext(docx_path)
        pdf_path = f"{base}.pdf"

    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)

    print(f"Exporting PDF: {abs_docx} -> {abs_pdf}")
    
    # Try Windows Word COM first
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            doc = word.Documents.Open(abs_docx)
            doc.SaveAs(abs_pdf, FileFormat=17) # 17 = wdFormatPDF
            doc.Close()
        finally:
            word.Quit()
            pythoncom.CoUninitialize()
        print(f"[SUCCESS] Exported via Word COM: {abs_pdf}")
        return abs_pdf
    except Exception as e:
        print(f"[INFO] Word COM not available or failed ({e}), trying docx2pdf...")

    try:
        from docx2pdf import convert
        convert(abs_docx, abs_pdf)
        print(f"[SUCCESS] Exported via docx2pdf: {abs_pdf}")
        return abs_pdf
    except Exception as e:
        print(f"[ERROR] PDF conversion failed: {e}")
        return None

def verify_page_count(pdf_path):
    if not os.path.exists(pdf_path):
        return None
    try:
        import pypdf
        reader = pypdf.PdfReader(pdf_path)
        count = len(reader.pages)
        print(f"Verified PDF page count: {count} pages")
        return count
    except Exception:
        try:
            import fitz
            doc = fitz.open(pdf_path)
            count = len(doc)
            print(f"Verified PDF page count: {count} pages")
            return count
        except Exception as e:
            print(f"Could not read PDF: {e}")
            return None

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Export Word document to PDF and verify page count.")
    parser.add_argument("docx_path", help="Path to input .docx document")
    parser.add_argument("-o", "--output", help="Path to output .pdf (optional)")
    args = parser.parse_args()

    pdf_out = convert_to_pdf(args.docx_path, args.output)
    if pdf_out:
        verify_page_count(pdf_out)
