#!/usr/bin/env python3
"""
Fix Gold layout backgrounds and font panose metadata issues
"""

import shutil
import zipfile
import os
import tempfile
import re

def fix_gold_template_issues():
    template_path = '4734_template.potx'
    fixed_template_path = '4734_template_final_fix.potx'
    
    # Work with a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract template
        with zipfile.ZipFile(template_path, 'r') as zip_read:
            zip_read.extractall(temp_dir)
        
        # Find and fix Gold layout files
        layout_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts')
        fixed_layouts = []
        
        for i in range(1, 15):
            layout_file = f'slideLayout{i}.xml'
            layout_path = os.path.join(layout_dir, layout_file)
            
            if os.path.exists(layout_path):
                with open(layout_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Check if this is a Gold layout
                if 'Gold_' in content:
                    print(f"Processing Gold layout: {layout_file}")
                    fixed_layouts.append(layout_file)
                    
                    try:
                        # 1. Remove hardcoded background (keep inheritance)
                        content = re.sub(r'<p:bg>.*?</p:bg>', '', content, flags=re.DOTALL)
                        
                        # 2. Fix problematic panose values for Helvetica Neue fonts
                        # Remove panose attributes that may be incorrect
                        content = re.sub(r' panose="[^"]*"', '', content)
                        content = re.sub(r' pitchFamily="[^"]*"', '', content)
                        content = re.sub(r' charset="[^"]*"', '', content)
                        
                        # Keep the Helvetica Neue typeface but remove problematic metadata
                        print(f"  Removed background and fixed font metadata in {layout_file}")
                        
                        with open(layout_path, 'w', encoding='utf-8') as f:
                            f.write(content)
                        
                    except Exception as e:
                        print(f"  Error processing {layout_file}: {e}")
        
        # Repack the template
        with zipfile.ZipFile(fixed_template_path, 'w', zipfile.ZIP_DEFLATED) as zip_write:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zip_write.write(file_path, arc_name)
    
    print(f"\nFinal fixed template saved as: {fixed_template_path}")
    print(f"Gold layouts processed: {fixed_layouts}")
    return fixed_template_path

if __name__ == "__main__":
    fix_gold_template_issues()