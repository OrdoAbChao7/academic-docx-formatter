# -*- coding: utf-8 -*-
"""
typography_patcher.py
Low-level OpenXML settings & styles patcher for Microsoft Word (.docx).
Enforces professional East Asian typography rules:
1. characterSpacingControl="compressPunctuation" (compresses full-width CJK punctuation)
2. kinsoku (East Asian line breaking rules: avoids starting lines with punctuation)
3. overflowPunct (allows punctuation to hang slightly rather than wrapping awkwardly)
4. doNotExpandShiftReturn (never stretches soft returns / shift-enter across margin)
5. autoSpaceDE / autoSpaceDN (clean spacing between CJK and Latin/Digits)
"""

import os
import zipfile
import tempfile
import shutil

def patch_docx_typography(docx_path):
    """Patch word/settings.xml and word/styles.xml inside docx package."""
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            zin.extractall(temp_dir)

        settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
        if os.path.exists(settings_path):
            with open(settings_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 1. characterSpacingControl
            if 'w:characterSpacingControl w:val="doNotCompress"' in content:
                content = content.replace(
                    'w:characterSpacingControl w:val="doNotCompress"',
                    'w:characterSpacingControl w:val="compressPunctuation"'
                )
            elif '<w:characterSpacingControl' not in content:
                content = content.replace(
                    '<w:settings ',
                    '<w:settings ><w:characterSpacingControl w:val="compressPunctuation"/>'
                )

            # 2. kinsoku & overflowPunct
            tags_to_add = []
            if '<w:kinsoku' not in content:
                tags_to_add.append('<w:kinsoku/>')
            if '<w:overflowPunct' not in content:
                tags_to_add.append('<w:overflowPunct/>')
            if '<w:autoSpaceDE' not in content:
                tags_to_add.append('<w:autoSpaceDE w:val="1"/>')
            if '<w:autoSpaceDN' not in content:
                tags_to_add.append('<w:autoSpaceDN w:val="1"/>')

            # 3. doNotExpandShiftReturn (Never stretch Shift+Enter)
            if '<w:compat>' in content:
                if 'doNotExpandShiftReturn' not in content:
                    content = content.replace('<w:compat>', '<w:compat><w:doNotExpandShiftReturn/>')
            else:
                tags_to_add.append('<w:compat><w:doNotExpandShiftReturn/></w:compat>')

            if tags_to_add:
                idx = content.find('</w:settings>')
                if idx != -1:
                    content = content[:idx] + ''.join(tags_to_add) + content[idx:]

            with open(settings_path, 'w', encoding='utf-8') as f:
                f.write(content)

        # Patch styles.xml for default East Asia language
        styles_path = os.path.join(temp_dir, 'word', 'styles.xml')
        if os.path.exists(styles_path):
            with open(styles_path, 'r', encoding='utf-8') as f:
                content = f.read()
            content = content.replace('w:eastAsia="ja-JP"', 'w:eastAsia="zh-CN"')
            content = content.replace('w:eastAsia="en-US"', 'w:eastAsia="zh-CN"')
            with open(styles_path, 'w', encoding='utf-8') as f:
                f.write(content)

        # Repackage zip
        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, temp_dir)
                    zout.write(full_path, rel_path)
    finally:
        shutil.rmtree(temp_dir)
