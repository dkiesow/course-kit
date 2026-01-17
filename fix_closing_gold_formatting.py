#!/usr/bin/env python3
"""
Fix closing slide placeholder and Gold_Bullets formatting specifically
"""

import shutil
import zipfile
import os
import tempfile
import re

def fix_closing_and_gold_bullets():
    template_path = 'templates/4734_template.potx'
    fixed_template_path = 'templates/4734_template_final_fix.potx'
    
    # Work with a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract template
        with zipfile.ZipFile(template_path, 'r') as zip_read:
            zip_read.extractall(temp_dir)
        
        # Fix Closing slide (slideLayout14.xml) - add body placeholder
        layout14_path = os.path.join(temp_dir, 'ppt', 'slideLayouts', 'slideLayout14.xml')
        if os.path.exists(layout14_path):
            with open(layout14_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            print("Fixing Mizzou_Close layout - adding body placeholder...")
            
            # Add a body placeholder below the title
            # Insert before closing </p:spTree>
            body_placeholder = '''<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{12345678-1234-5678-9012-123456789012}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1257300" y="4800000"/><a:ext cx="15773400" cy="3200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>'''
            
            content = content.replace('</p:spTree>', body_placeholder + '</p:spTree>')
            
            with open(layout14_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print("  Added body placeholder to Mizzou_Close layout")
        
        # Fix Gold_Bullets (slideLayout7.xml) - remove individual formatting, use empty lstStyle
        layout7_path = os.path.join(temp_dir, 'ppt', 'slideLayouts', 'slideLayout7.xml')
        if os.path.exists(layout7_path):
            with open(layout7_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            print("Fixing Gold_Bullets layout - removing individual formatting...")
            
            # Replace the complex body content with simple placeholder that inherits from master
            # Find the body placeholder and simplify its txBody
            pattern = r'(<p:sp><p:nvSpPr><p:cNvPr id="3"[^>]*>.*?</p:nvSpPr></p:nvSpPr><p:spPr>.*?</p:spPr><p:txBody><a:bodyPr[^>]*><a:normAutofit/></a:bodyPr>)<a:lstStyle/>.*?</p:txBody>'
            
            replacement = r'\1<a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Fifth level</a:t></a:r></a:p></p:txBody>'
            
            content = re.sub(pattern, replacement, content, flags=re.DOTALL)
            
            with open(layout7_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print("  Simplified Gold_Bullets to inherit master formatting")
        
        # Repack the template
        with zipfile.ZipFile(fixed_template_path, 'w', zipfile.ZIP_DEFLATED) as zip_write:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zip_write.write(file_path, arc_name)
    
    print(f"\nFixed template saved as: {fixed_template_path}")
    return fixed_template_path

if __name__ == "__main__":
    fix_closing_and_gold_bullets()