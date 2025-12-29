#!/usr/bin/env python3
"""
Fix the Gold master slide bodyStyle to use explicit Helvetica Neue fonts
"""

import shutil
import zipfile
import os
import tempfile
import re

def fix_gold_master_bodyStyle():
    template_path = '4734_template.potx'
    fixed_template_path = '4734_template_master_fix.potx'
    
    # Work with a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract template
        with zipfile.ZipFile(template_path, 'r') as zip_read:
            zip_read.extractall(temp_dir)
        
        # Fix the Gold master slide (slideMaster2.xml)
        master_path = os.path.join(temp_dir, 'ppt', 'slideMasters', 'slideMaster2.xml')
        
        if os.path.exists(master_path):
            with open(master_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            print("Fixing Gold master slide bodyStyle...")
            
            # Replace theme font references with explicit Helvetica Neue in bodyStyle
            content = re.sub(
                r'<a:latin typeface="\+mn-lt"/>',
                '<a:latin typeface="Helvetica Neue Light"/>',
                content
            )
            content = re.sub(
                r'<a:ea typeface="\+mn-ea"/>',
                '<a:ea typeface="Helvetica Neue Light"/>',
                content
            )
            
            # Remove problematic panose attributes from master
            content = re.sub(r' panose="[^"]*"', '', content)
            content = re.sub(r' pitchFamily="[^"]*"', '', content)
            content = re.sub(r' charset="[^"]*"', '', content)
            
            with open(master_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print("  Fixed Gold master slide bodyStyle")
        
        # Also ensure Gold layouts have empty lstStyle so they inherit from master
        layout_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts')
        
        for i in range(1, 15):
            layout_file = f'slideLayout{i}.xml'
            layout_path = os.path.join(layout_dir, layout_file)
            
            if os.path.exists(layout_path):
                with open(layout_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Check if this is a Gold layout (except Gold_Quote which works)
                if 'Gold_' in content and 'Gold_Quote' not in content:
                    print(f"Processing Gold layout: {layout_file}")
                    
                    # Ensure body placeholder has empty lstStyle to inherit from master
                    # Find body placeholders and ensure they have empty lstStyle
                    content = re.sub(
                        r'(<p:nvPr><p:ph[^>]*idx="[^"]*"[^>]*></p:ph></p:nvPr>.*?<a:bodyPr[^>]*>.*?</a:bodyPr>)<a:lstStyle>.*?</a:lstStyle>',
                        r'\1<a:lstStyle/>',
                        content,
                        flags=re.DOTALL
                    )
                    
                    with open(layout_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    
                    print(f"  Ensured {layout_file} inherits from master bodyStyle")
        
        # Repack the template
        with zipfile.ZipFile(fixed_template_path, 'w', zipfile.ZIP_DEFLATED) as zip_write:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zip_write.write(file_path, arc_name)
    
    print(f"\nMaster-fixed template saved as: {fixed_template_path}")
    return fixed_template_path

if __name__ == "__main__":
    fix_gold_master_bodyStyle()