#!/usr/bin/env python3
"""
Check for XML validation issues in our template fixes
"""

import xml.etree.ElementTree as ET
import zipfile
import tempfile

def validate_template_xml():
    template_path = '4734_template.potx'
    
    print("=== VALIDATING TEMPLATE XML ===")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract template
        with zipfile.ZipFile(template_path, 'r') as zip_read:
            zip_read.extractall(temp_dir)
        
        # Check Gold layout files for XML validity
        import os
        layout_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts')
        
        for i in range(1, 15):
            layout_file = f'slideLayout{i}.xml'
            layout_path = os.path.join(layout_dir, layout_file)
            
            if os.path.exists(layout_path):
                with open(layout_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Check if this is a Gold layout
                if 'Gold_' in content:
                    print(f"\nValidating {layout_file}:")
                    
                    try:
                        # Try to parse as XML
                        root = ET.fromstring(content)
                        print(f"  ✓ Valid XML")
                        
                        # Check for empty elements that might cause issues
                        empty_elements = []
                        for elem in root.iter():
                            if elem.text is None and len(elem) == 0 and len(elem.attrib) == 0:
                                empty_elements.append(elem.tag)
                        
                        if empty_elements:
                            print(f"  ⚠️  Found {len(empty_elements)} potentially problematic empty elements")
                            for tag in set(empty_elements[:5]):  # Show unique tags, limit to 5
                                print(f"     - {tag}")
                        
                        # Check for malformed attributes
                        malformed = []
                        if '<a:latin/>' in content and 'typeface=""' in content:
                            malformed.append("Empty typeface attributes")
                        if '<a:defRPr/>' in content and 'sz=""' in content:
                            malformed.append("Empty size attributes")
                            
                        if malformed:
                            print(f"  ❌ Potential issues: {', '.join(malformed)}")
                        
                    except ET.ParseError as e:
                        print(f"  ❌ XML Parse Error: {e}")
                        # Show the problematic area
                        lines = content.split('\n')
                        error_line = getattr(e, 'lineno', 1) - 1
                        start = max(0, error_line - 2)
                        end = min(len(lines), error_line + 3)
                        print(f"     Context around line {error_line + 1}:")
                        for i in range(start, end):
                            marker = " >>> " if i == error_line else "     "
                            print(f"{marker}{i+1}: {lines[i]}")

if __name__ == "__main__":
    validate_template_xml()