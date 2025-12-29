#!/usr/bin/env python3
"""
Examine the raw slide XML to understand font differences between White and Gold
"""

import zipfile
import xml.etree.ElementTree as ET

# Examine both slides in the test file
def examine_slide_xml():
    test_file = 'debug_fonts_test.pptx'
    
    print("=== RAW SLIDE XML COMPARISON ===")
    
    with zipfile.ZipFile(test_file, 'r') as zip_read:
        # Get both slides
        slide1_xml = zip_read.read('ppt/slides/slide1.xml').decode('utf-8')
        slide2_xml = zip_read.read('ppt/slides/slide2.xml').decode('utf-8')
        
        for i, (slide_name, xml_content) in enumerate([
            ("White_Bullets (slide1)", slide1_xml),
            ("Gold_Bullets (slide2)", slide2_xml)
        ], 1):
            print(f"\n{slide_name}:")
            
            # Parse and look for font definitions
            root = ET.fromstring(xml_content)
            
            # Look for any font references
            latin_fonts = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
            print(f"  Latin font elements: {len(latin_fonts)}")
            for font in latin_fonts:
                typeface = font.get('typeface', 'No typeface')
                print(f"    {typeface}")
            
            # Look for theme font references
            theme_refs = 0
            for elem in root.iter():
                if elem.get('typeface') in ['+mn-lt', '+mj-lt']:
                    theme_refs += 1
                    print(f"    Theme font: {elem.get('typeface')}")
            
            # Look for defRPr elements
            def_rpr_elements = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr")
            print(f"  defRPr elements: {len(def_rpr_elements)}")
            
            # Check if there are any solidFill color specifications that might indicate inheritance issues
            solid_fills = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill")
            print(f"  solidFill elements: {len(solid_fills)}")
            
            print(f"  Full XML preview (first 500 chars):")
            print(f"    {xml_content[:500]}...")

if __name__ == "__main__":
    examine_slide_xml()