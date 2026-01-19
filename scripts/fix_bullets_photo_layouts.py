#!/usr/bin/env python3
"""
Fix the Bullets_Photo layouts to have proper content placeholders by modifying XML directly
"""

import shutil
import zipfile
import os
import tempfile
import re

def fix_bullets_photo_placeholders():
    template_path = 'templates/4734_template.potx'
    fixed_template_path = 'templates/4734_template_bullets_fix.potx'
    
    # Work with a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract template
        with zipfile.ZipFile(template_path, 'r') as zip_read:
            zip_read.extractall(temp_dir)
        
        # Fix White_Bullets_Photo (slideLayout3.xml)
        layout3_path = os.path.join(temp_dir, 'ppt', 'slideLayouts', 'slideLayout3.xml')
        if os.path.exists(layout3_path):
            with open(layout3_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            print("Fixing White_Bullets_Photo layout...")
            
            # The text shape has userDrawn="1" and no <p:ph> - need to add placeholder element
            # Find: <p:nvPr userDrawn="1"/>
            # Replace with: <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
            content = re.sub(
                r'(<p:cNvSpPr txBox="1"><a:spLocks/></p:cNvSpPr><p:nvPr) userDrawn="1"(/>)',
                r'\1><p:ph type="body" idx="1"/></p:nvPr>',
                content
            )
            
            with open(layout3_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print("  Fixed White_Bullets_Photo layout")
        
        # Fix Gold_Bullets_Photo (slideLayout9.xml) 
        layout9_path = os.path.join(temp_dir, 'ppt', 'slideLayouts', 'slideLayout9.xml')
        if os.path.exists(layout9_path):
            with open(layout9_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            print("Fixing Gold_Bullets_Photo layout...")
            
            # Same fix for Gold layout
            content = re.sub(
                r'(<p:cNvSpPr txBox="1"><a:spLocks/></p:cNvSpPr><p:nvPr) userDrawn="1"(/>)',
                r'\1><p:ph type="body" idx="1"/></p:nvPr>',
                content
            )
            
            with open(layout9_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print("  Fixed Gold_Bullets_Photo layout")
        
        # Repack the template
        with zipfile.ZipFile(fixed_template_path, 'w', zipfile.ZIP_DEFLATED) as zip_write:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zip_write.write(file_path, arc_name)
    
    print(f"\nPlaceholder-fixed template saved as: {fixed_template_path}")
    return fixed_template_path

if __name__ == "__main__":
    fix_bullets_photo_placeholders()