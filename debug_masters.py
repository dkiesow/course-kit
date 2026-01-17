#!/usr/bin/env python3
"""
Debug the master slides to see differences between Gold and White layouts
"""

from pptx import Presentation
import shutil
import zipfile

# Patch POTX to PPTX
template_path = 'templates/4734_template.potx'
temp_path = 'temp_template.pptx'

shutil.copy(template_path, temp_path)

with zipfile.ZipFile(temp_path, 'r') as zip_read:
    content_types = zip_read.read('[Content_Types].xml').decode('utf-8')
    content_types = content_types.replace(
        'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
    )
    
    with zipfile.ZipFile(temp_path + '.new', 'w', zipfile.ZIP_DEFLATED) as zip_write:
        for item in zip_read.infolist():
            data = zip_read.read(item.filename)
            if item.filename == '[Content_Types].xml':
                zip_write.writestr(item, content_types.encode('utf-8'))
            else:
                zip_write.writestr(item, data)

shutil.move(temp_path + '.new', temp_path)

prs = Presentation(temp_path)

print("=== MASTER SLIDES ANALYSIS ===")
for i, master in enumerate(prs.slide_masters):
    print(f"\nMaster {i+1}:")
    print(f"  Master name: {getattr(master, 'name', 'No name')}")
    print(f"  Number of layouts: {len(master.slide_layouts)}")
    
    # Check if master has background
    if hasattr(master, 'background'):
        print(f"  Background type: {master.background.fill.type if hasattr(master.background, 'fill') else 'No fill'}")
    
    print("  Layouts:")
    for j, layout in enumerate(master.slide_layouts):
        print(f"    {j+1}. {layout.name}")
        
        # Check layout background
        if hasattr(layout, 'background'):
            bg_type = layout.background.fill.type if hasattr(layout.background, 'fill') else 'No fill'
            print(f"       Background: {bg_type}")
        
        # Check if it's Gold or White layout
        if 'Gold' in layout.name:
            print(f"       *** GOLD LAYOUT ***")
            print(f"       Layout ID: {getattr(layout, '_element', {}).get('id', 'No ID')}")
        elif 'White' in layout.name:
            print(f"       *** WHITE LAYOUT ***")
            print(f"       Layout ID: {getattr(layout, '_element', {}).get('id', 'No ID')}")

print("\n=== THEME INFORMATION ===")
for i, master in enumerate(prs.slide_masters):
    theme = getattr(master, 'theme', None)
    if theme:
        print(f"Master {i+1} theme: {theme}")
    else:
        print(f"Master {i+1}: No theme found")