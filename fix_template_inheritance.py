#!/usr/bin/env python3
"""
Fix placeholder inheritance in POTX template.
The layout slides need to reference master placeholders in their XML.
"""

import zipfile
import shutil
import os
from xml.etree import ElementTree as ET

# Namespace map for PowerPoint XML
namespaces = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

# Register namespaces to preserve prefixes
for prefix, uri in namespaces.items():
    ET.register_namespace(prefix, uri)

template_path = 'templates/4734_template.potx'
backup_path = 'templates/4734_template_backup.potx'
temp_extract = 'temp_potx_extract'

# Backup original
if not os.path.exists(backup_path):
    shutil.copy(template_path, backup_path)
    print(f"Created backup: {backup_path}")

# Extract POTX
os.makedirs(temp_extract, exist_ok=True)
with zipfile.ZipFile(template_path, 'r') as zip_ref:
    zip_ref.extractall(temp_extract)
print(f"Extracted POTX to {temp_extract}")

# Find all layout XML files
layout_dir = os.path.join(temp_extract, 'ppt', 'slideLayouts')
layout_files = [f for f in os.listdir(layout_dir) if f.startswith('slideLayout') and f.endswith('.xml')]

print(f"\nFound {len(layout_files)} layout files")

# Body placeholder template (to add to layouts that need it)
# This references idx 1 (body) from the master
body_placeholder_xml = """
<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="3" name="Content Placeholder 2"/>
    <p:cNvSpPr>
      <a:spLocks noGrp="1"/>
    </p:cNvSpPr>
    <p:nvPr>
      <p:ph idx="1"/>
    </p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr lvl="0"/>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </p:txBody>
</p:sp>
"""

# Check each layout
for layout_file in layout_files:
    layout_path = os.path.join(layout_dir, layout_file)
    
    try:
        tree = ET.parse(layout_path)
        root = tree.getroot()
        
        # Find the cSld (common slide data) element
        cSld = root.find('.//p:cSld', namespaces)
        if cSld is None:
            print(f"  {layout_file}: No cSld found, skipping")
            continue
        
        # Find spTree (shape tree)
        spTree = cSld.find('.//p:spTree', namespaces)
        if spTree is None:
            print(f"  {layout_file}: No spTree found, skipping")
            continue
        
        # Check existing placeholders
        placeholders = []
        for sp in spTree.findall('.//p:sp', namespaces):
            nvPr = sp.find('.//p:nvPr', namespaces)
            if nvPr is not None:
                ph = nvPr.find('.//p:ph', namespaces)
                if ph is not None:
                    idx = ph.get('idx', '0')
                    ph_type = ph.get('type', 'unknown')
                    placeholders.append((idx, ph_type))
        
        print(f"\n  {layout_file}:")
        print(f"    Current placeholders: {placeholders}")
        
        # Check if body placeholder (idx=1) exists
        has_body = any(idx == '1' for idx, _ in placeholders)
        
        if not has_body:
            print(f"    ⚠️  Missing body placeholder (idx=1), adding it...")
            
            # Parse the body placeholder XML
            body_elem = ET.fromstring(body_placeholder_xml)
            
            # Add to spTree (after title, which is usually first sp element)
            sp_elements = spTree.findall('p:sp', namespaces)
            if sp_elements:
                # Insert after the first sp (title)
                insert_pos = list(spTree).index(sp_elements[0]) + 1
                spTree.insert(insert_pos, body_elem)
                print(f"    ✅ Added body placeholder")
                
                # Save the modified XML
                tree.write(layout_path, encoding='utf-8', xml_declaration=True)
            else:
                print(f"    ❌ No sp elements found to insert after")
        else:
            print(f"    ✅ Already has body placeholder")
            
    except Exception as e:
        print(f"  {layout_file}: Error - {e}")

# Repackage POTX
print(f"\nRepackaging POTX...")
if os.path.exists(template_path):
    os.remove(template_path)

with zipfile.ZipFile(template_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
    for root, dirs, files in os.walk(temp_extract):
        for file in files:
            file_path = os.path.join(root, file)
            arc_name = os.path.relpath(file_path, temp_extract)
            zip_out.write(file_path, arc_name)

print(f"✅ Created fixed POTX: {template_path}")

# Cleanup
shutil.rmtree(temp_extract)
print(f"Cleaned up temp directory")

print(f"\n{'='*60}")
print(f"Template fixed! Original backed up to: {backup_path}")
print(f"{'='*60}")
