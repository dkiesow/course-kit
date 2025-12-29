#!/usr/bin/env python3
"""
Add proper list style font definitions to Gold body placeholders
"""

import shutil
import zipfile
import os
import tempfile
import re

def add_body_lstyle_fonts():
    template_path = '4734_template.potx'
    fixed_template_path = '4734_template_lstyle_fix.potx'
    
    # Work with a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract template
        with zipfile.ZipFile(template_path, 'r') as zip_read:
            zip_read.extractall(temp_dir)
        
        # Find and fix Gold layout files
        layout_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts')
        fixed_layouts = []
        
        for i in range(1, 15):
            layout_file = f'slideLayout{i}.xml'
            layout_path = os.path.join(layout_dir, layout_file)
            
            if os.path.exists(layout_path):
                with open(layout_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Check if this is a Gold layout (except Gold_Quote which already works)
                if 'Gold_' in content and 'Gold_Quote' not in content:
                    print(f"Processing Gold layout: {layout_file}")
                    fixed_layouts.append(layout_file)
                    
                    try:
                        # Add proper lstStyle font definitions for body placeholders
                        # Find empty lstStyle tags and replace them with proper font definitions
                        body_lstyle = '''<a:lstStyle><a:lvl1pPr marL="342900" marR="0" indent="-342900" algn="l" defTabSz="1371600" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
<a:lnSpc><a:spcPct val="120000"/></a:lnSpc>
<a:spcBef><a:spcPts val="500"/></a:spcBef>
<a:spcAft><a:spcPts val="1500"/></a:spcAft>
<a:buClrTx/>
<a:buSzTx/>
<a:buFont typeface="Arial"/>
<a:buChar char="•"/>
<a:defRPr sz="3600" b="0" i="0" kern="1200">
<a:solidFill><a:prstClr val="black"/></a:solidFill>
<a:latin typeface="Helvetica Neue Light"/>
<a:ea typeface="Helvetica Neue Light"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl1pPr>
<a:lvl2pPr marL="1028700" marR="0" indent="-342900" algn="l" defTabSz="1371600" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>
<a:spcBef><a:spcPts val="750"/></a:spcBef>
<a:spcAft><a:spcPts val="0"/></a:spcAft>
<a:buClrTx/>
<a:buSzTx/>
<a:buFont typeface="Arial"/>
<a:buChar char="•"/>
<a:defRPr sz="3000" b="0" i="0" kern="1200">
<a:solidFill><a:prstClr val="black"/></a:solidFill>
<a:latin typeface="Helvetica Neue Light"/>
<a:ea typeface="Helvetica Neue Light"/>
<a:cs typeface="+mn-cs"/>
</a:defRPr>
</a:lvl2pPr></a:lstStyle>'''
                        
                        # Find body placeholders with empty lstStyle (those are the ones that need font definitions)
                        # Pattern: body placeholder followed by empty lstStyle
                        pattern = r'(<p:nvPr><p:ph[^>]*(?:idx|type)="[^"]*"[^>]*></p:ph></p:nvPr>.*?<a:bodyPr[^>]*>.*?</a:bodyPr>)<a:lstStyle/>'
                        
                        # Replace empty lstStyle with proper font definitions
                        content = re.sub(pattern, r'\1' + body_lstyle, content, flags=re.DOTALL)
                        
                        print(f"  Added list style fonts to {layout_file}")
                        
                        with open(layout_path, 'w', encoding='utf-8') as f:
                            f.write(content)
                        
                    except Exception as e:
                        print(f"  Error processing {layout_file}: {e}")
        
        # Repack the template
        with zipfile.ZipFile(fixed_template_path, 'w', zipfile.ZIP_DEFLATED) as zip_write:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zip_write.write(file_path, arc_name)
    
    print(f"\nList style fixed template saved as: {fixed_template_path}")
    print(f"Gold layouts processed: {fixed_layouts}")
    return fixed_template_path

if __name__ == "__main__":
    add_body_lstyle_fonts()