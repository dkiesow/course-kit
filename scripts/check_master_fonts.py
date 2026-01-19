#!/usr/bin/env python3
"""
Check font definitions in master slides
"""

import xml.etree.ElementTree as ET
import os

def analyze_master_fonts():
    master_dir = "template_analysis/ppt/slideMasters"
    
    print("=== MASTER SLIDE FONT ANALYSIS ===")
    
    for i in range(1, 5):
        master_file = f"slideMaster{i}.xml"
        master_path = os.path.join(master_dir, master_file)
        
        if os.path.exists(master_path):
            with open(master_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            try:
                root = ET.fromstring(content)
                
                # Find layout names in this master
                layouts = []
                # This is a simplified check - we'd need to check the relationships
                if "Gold" in content:
                    print(f"\n{master_file} (likely Gold master):")
                elif "White" in content:
                    print(f"\n{master_file} (likely White master):")
                else:
                    print(f"\n{master_file}:")
                
                # Find all font references
                font_elements = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
                print(f"  Latin font elements: {len(font_elements)}")
                for j, font in enumerate(font_elements):
                    typeface = font.get('typeface', 'No typeface')
                    print(f"    {j+1}. {typeface}")
                
                # Check theme fonts specifically
                theme_fonts = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}theme")
                print(f"  Theme references: {len(theme_fonts)}")
                
            except ET.ParseError as e:
                print(f"  Parse error: {e}")
        else:
            print(f"\n{master_file}: Not found")

if __name__ == "__main__":
    analyze_master_fonts()