#!/usr/bin/env python3
"""
Check which master slide each layout is connected to.
"""

import zipfile
import os
from xml.etree import ElementTree as ET

template_path = '4734_template.potx'
temp_extract = 'temp_check_masters'

# Extract POTX
os.makedirs(temp_extract, exist_ok=True)
with zipfile.ZipFile(template_path, 'r') as zip_ref:
    zip_ref.extractall(temp_extract)

# Namespaces
ns = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

# Check each layout's relationship to master
layout_dir = os.path.join(temp_extract, 'ppt', 'slideLayouts')
layout_rels_dir = os.path.join(temp_extract, 'ppt', 'slideLayouts', '_rels')

layout_files = sorted([f for f in os.listdir(layout_dir) if f.startswith('slideLayout') and f.endswith('.xml')])

print("Layout -> Master Relationships:\n")

for layout_file in layout_files:
    layout_path = os.path.join(layout_dir, layout_file)
    
    # Parse layout to get name
    tree = ET.parse(layout_path)
    root = tree.getroot()
    cSld = root.find('.//p:cSld', ns)
    layout_name = cSld.get('name') if cSld is not None else 'Unknown'
    
    # Check the relationship file
    layout_num = layout_file.replace('slideLayout', '').replace('.xml', '')
    rels_file = os.path.join(layout_rels_dir, f'slideLayout{layout_num}.xml.rels')
    
    master_rel = None
    if os.path.exists(rels_file):
        rels_tree = ET.parse(rels_file)
        rels_root = rels_tree.getroot()
        
        # Look for slideMaster relationship
        for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            if 'slideMaster' in rel.get('Type', ''):
                master_rel = rel.get('Target')
                break
    
    print(f"{layout_file:20} -> {layout_name:30} -> {master_rel}")

# Cleanup
import shutil
shutil.rmtree(temp_extract)
