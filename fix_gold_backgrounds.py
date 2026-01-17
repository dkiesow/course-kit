#!/usr/bin/env python3
"""
Fix Gold layout backgrounds by removing hardcoded background fills
so they inherit from the master slide properly.
"""

import xml.etree.ElementTree as ET
import shutil
import zipfile
import os
import tempfile

def fix_gold_layout_backgrounds():
    template_path = 'templates/4734_template.potx'
    fixed_template_path = 'templates/4734_template_fixed.potx'
    
    # Create a backup
    shutil.copy(template_path, 'templates/4734_template_backup.potx')
    
    # Work with a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract template
        with zipfile.ZipFile(template_path, 'r') as zip_read:
            zip_read.extractall(temp_dir)
        
        # Find and fix Gold layout files
        layout_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts')
        gold_layouts = []
        
        for i in range(1, 15):
            layout_file = f'slideLayout{i}.xml'
            layout_path = os.path.join(layout_dir, layout_file)
            
            if os.path.exists(layout_path):
                with open(layout_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Check if this is a Gold layout
                if 'Gold_' in content:
                    print(f"Processing Gold layout: {layout_file}")
                    gold_layouts.append(layout_file)
                    
                    try:
                        # Parse XML
                        root = ET.fromstring(content)
                        
                        # Find and remove background elements
                        bg_elements = root.findall(".//{http://schemas.openxmlformats.org/presentationml/2006/main}bg")
                        
                        if bg_elements:
                            print(f"  Found {len(bg_elements)} background elements to remove")
                            for bg_elem in bg_elements:
                                parent = root.find(".//*[{http://schemas.openxmlformats.org/presentationml/2006/main}bg='.']/..")
                                if parent is not None:
                                    parent.remove(bg_elem)
                                else:
                                    # Try finding parent differently
                                    for elem in root.iter():
                                        if bg_elem in elem:
                                            elem.remove(bg_elem)
                                            break
                        
                        # Write back the fixed XML
                        # Need to preserve XML declaration and formatting
                        xml_str = ET.tostring(root, encoding='unicode')
                        xml_str = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml_str
                        
                        with open(layout_path, 'w', encoding='utf-8') as f:
                            f.write(xml_str)
                        
                        print(f"  Fixed {layout_file}")
                        
                    except ET.ParseError as e:
                        print(f"  Parse error in {layout_file}: {e}")
                    except Exception as e:
                        print(f"  Error processing {layout_file}: {e}")
        
        # Repack the template
        with zipfile.ZipFile(fixed_template_path, 'w', zipfile.ZIP_DEFLATED) as zip_write:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zip_write.write(file_path, arc_name)
    
    print(f"\nFixed template saved as: {fixed_template_path}")
    print(f"Gold layouts processed: {gold_layouts}")
    return fixed_template_path

if __name__ == "__main__":
    fix_gold_layout_backgrounds()