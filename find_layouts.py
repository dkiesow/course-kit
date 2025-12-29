#!/usr/bin/env python3
"""
Find which slideLayout files correspond to Gold vs White layouts
"""

import xml.etree.ElementTree as ET
import os

layout_dir = "template_analysis/ppt/slideLayouts"

print("=== LAYOUT ANALYSIS ===")
for i in range(1, 15):
    layout_file = f"slideLayout{i}.xml"
    layout_path = os.path.join(layout_dir, layout_file)
    
    if os.path.exists(layout_path):
        with open(layout_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        try:
            root = ET.fromstring(content)
            
            # Find the layout name
            name_elem = root.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
            if name_elem is not None:
                name_attr = name_elem.get('name')
                if name_attr:
                    print(f"\n{layout_file}: {name_attr}")
                else:
                    print(f"\n{layout_file}: (no name)")
            else:
                print(f"\n{layout_file}: (no cSld element)")
                
            # Check for background properties
            bg_elements = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill")
            if bg_elements:
                print(f"  Has solidFill elements: {len(bg_elements)}")
                
            bg_ref_elements = root.findall(".//{http://schemas.openxmlformats.org/presentationml/2006/main}bgRef")
            if bg_ref_elements:
                print(f"  Has bgRef elements: {len(bg_ref_elements)}")
                for bg_ref in bg_ref_elements:
                    idx = bg_ref.get('idx')
                    print(f"    bgRef idx: {idx}")
            
        except ET.ParseError as e:
            print(f"\n{layout_file}: Parse error - {e}")
        except Exception as e:
            print(f"\n{layout_file}: Error - {e}")