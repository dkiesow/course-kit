#!/usr/bin/env python3
import zipfile
import xml.etree.ElementTree as ET
import sys

filepath = sys.argv[1] if len(sys.argv) > 1 else 'Week_Week One Wednesday_1_21_26.pptx'

try:
    with zipfile.ZipFile(filepath, 'r') as z:
        namelist = z.namelist()
        
        # Check relationships
        if 'ppt/_rels/presentation.xml.rels' in namelist:
            rels_content = z.read('ppt/_rels/presentation.xml.rels')
            rels_tree = ET.fromstring(rels_content)
            
            # Get all relationship IDs
            rel_ids = {}
            for rel in rels_tree.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_id = rel.get('Id')
                rel_target = rel.get('Target')
                rel_ids[rel_id] = rel_target
            
            print(f'Relationships found: {len(rel_ids)}')
            
            # Check if all referenced slides exist
            pres_content = z.read('ppt/presentation.xml')
            pres_tree = ET.fromstring(pres_content)
            
            # Get slide IDs from presentation
            slide_refs = []
            for sld in pres_tree.findall('.//{http://schemas.openxmlformats.org/presentationml/2006/main}sldId'):
                r_id = sld.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                slide_refs.append(r_id)
            
            print(f'Slide references in presentation.xml: {len(slide_refs)}')
            
            # Check if all exist
            missing = []
            for ref in slide_refs:
                if ref not in rel_ids:
                    missing.append(ref)
                else:
                    target = rel_ids[ref]
                    full_path = 'ppt/' + target
                    if full_path not in namelist:
                        print(f'Missing slide file: {full_path} (referenced by {ref})')
            
            if missing:
                print(f'Missing relationship IDs: {missing}')
            else:
                print('✓ All slide references have relationships')
            
            # Check for slide relationship files
            print('\nChecking slide relationship files...')
            for i in range(1, len(slide_refs) + 1):
                slide_rels_path = f'ppt/slides/_rels/slide{i}.xml.rels'
                if slide_rels_path not in namelist:
                    print(f'  Missing: {slide_rels_path}')
            
except Exception as e:
    print(f'Error: {e}')
    import traceback
    traceback.print_exc()
