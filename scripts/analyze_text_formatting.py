#!/usr/bin/env python3
"""
Compare text formatting between Gold and White layouts
"""

import xml.etree.ElementTree as ET
import os

def analyze_text_formatting():
    layout_dir = "template_analysis/ppt/slideLayouts"
    
    # Compare White_Bullets vs Gold_Bullets
    white_bullets = os.path.join(layout_dir, "slideLayout1.xml")  # White_Bullets
    gold_bullets = os.path.join(layout_dir, "slideLayout7.xml")   # Gold_Bullets
    
    print("=== TEXT FORMATTING COMPARISON ===")
    
    for layout_name, layout_path in [("White_Bullets", white_bullets), ("Gold_Bullets", gold_bullets)]:
        print(f"\n{layout_name}:")
        
        if os.path.exists(layout_path):
            with open(layout_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            try:
                root = ET.fromstring(content)
                
                # Find all font references
                font_elements = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
                print(f"  Latin font elements: {len(font_elements)}")
                for i, font in enumerate(font_elements):
                    typeface = font.get('typeface', 'No typeface')
                    print(f"    {i+1}. {typeface}")
                
                # Find defRPr (default run properties) elements
                def_rpr_elements = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr")
                print(f"  Default run properties: {len(def_rpr_elements)}")
                for i, def_rpr in enumerate(def_rpr_elements):
                    print(f"    defRPr {i+1}:")
                    # Look for font children
                    latin_fonts = def_rpr.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
                    for j, latin in enumerate(latin_fonts):
                        typeface = latin.get('typeface', 'No typeface')
                        print(f"      Latin font: {typeface}")
                
                # Find lstStyle elements (list styles)
                lst_style_elements = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}lstStyle")
                print(f"  List style elements: {len(lst_style_elements)}")
                
                # Look for any explicit text formatting
                rpr_elements = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
                print(f"  Run property elements: {len(rpr_elements)}")
                
            except ET.ParseError as e:
                print(f"  Parse error: {e}")
        else:
            print(f"  File not found: {layout_path}")

if __name__ == "__main__":
    analyze_text_formatting()