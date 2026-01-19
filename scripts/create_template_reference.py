#!/usr/bin/env python3
"""
Export one slide per template type to create a reference PPTX for master slide design.
"""

import sqlite3
import json
from datetime import datetime
import os

DB_PATH = 'presentations.db'

# Map database template_base to CSS class names
TEMPLATE_CLASS_MAP = {
    'bullets': 'template-bullets',
    'bullets-image': 'template-bullets-image',
    'bullets-image-split': 'template-bullets-image-split',
    'lines': 'template-lines',
    'quote': 'template-quote-headline',
    'gold-bullets': 'gold-bullets',
    'gold-bullets-image-split': 'gold-bullets-image-split',
    'gold-quote': 'gold-quote',
    'photo-centered': 'photo-centered',
    'closing': 'closing',
    'title': 'title'
}

def get_template_samples():
    """Get one slide for each template_base type"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Get distinct template_base values
    c.execute('''SELECT DISTINCT template_base FROM slides WHERE template_base IS NOT NULL ORDER BY template_base''')
    templates = [row[0] for row in c.fetchall()]
    
    slides_by_template = {}
    
    # For each template, get one representative slide
    for template in templates:
        c.execute('''SELECT id, deck_id, headline, paragraph, bullets, quote, quote_citation, 
                            image_path, template_base, is_title, slide_class
                     FROM slides 
                     WHERE template_base = ? AND is_title = 0
                     LIMIT 1''', (template,))
        row = c.fetchone()
        if row:
            slides_by_template[template] = row
    
    conn.close()
    return slides_by_template

def build_markdown(slides_dict):
    """Build markdown content for reference slides"""
    content = '''---
marp: true
theme: classroom
paginate: false
---

# Template Reference Guide

One example slide per template type for master slide design.

---

'''
    
    for template, slide_data in sorted(slides_dict.items()):
        slide_id, deck_id, headline, paragraph, bullets, quote, quote_citation, image_path, template_base, is_title, slide_class = slide_data
        
        content += f'\n---\n\n'
        content += f'<!-- _class: {slide_class} -->\n'
        
        if headline:
            content += f'# {headline}\n'
        
        if paragraph:
            content += f'\n{paragraph}\n'
        
        if bullets:
            try:
                bullet_list = json.loads(bullets)
                if bullet_list:
                    content += '\n'
                    for bullet in bullet_list:
                        if bullet.strip():
                            content += f'- {bullet}\n'
            except:
                pass
        
        if quote:
            content += f'\n> {quote}\n'
            if quote_citation:
                content += f'> — {quote_citation}\n'
        
        if image_path:
            if image_path.startswith('/assets/'):
                image_path = './assets/' + image_path[8:]
            elif not image_path.startswith('./'):
                image_path = './assets/' + image_path
            content += f'\n![Image]({image_path})\n'
    
    return content

def main():
    print("Fetching template samples...")
    slides = get_template_samples()
    
    if not slides:
        print("No slides found with template_base values")
        return
    
    print(f"Found {len(slides)} unique templates:")
    for template in sorted(slides.keys()):
        print(f"  - {template}")
    
    print("\nBuilding markdown...")
    markdown = build_markdown(slides)
    
    # Write to temp file
    temp_md = 'output/template_reference.md'
    os.makedirs('output', exist_ok=True)
    with open(temp_md, 'w') as f:
        f.write(markdown)
    print(f"Wrote markdown to {temp_md}")
    
    # Export to PPTX using Marp with editable flag
    temp_md = 'output/template_reference.md'
    output_pptx = 'output/Template_Reference_Styled.pptx'
    
    print(f"\nExporting to styled PPTX with Marp...")
    cmd = f'marp "{temp_md}" -o "{output_pptx}" --allow-local-files --pptx --theme presentation-styles.css --pptx-editable'
    
    print(f"Running: {cmd}")
    
    import subprocess
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    
    if result.returncode == 0:
        print(f"\n✅ Success! Created {output_pptx}")
        print("\nThis PPTX has:")
        print("- Full CSS styling preserved (backgrounds, colors, fonts, positioning)")
        print("- Editable text in PowerPoint")
        print("- One example per template type")
        print("\nYou can now:")
        print("1. Open the PPTX in PowerPoint")
        print("2. View → Slide Master to see/edit master layouts")
        print("3. Use as reference for understanding the style system")
    else:
        print(f"\n❌ Error: {result.stderr}")
        return False
    
    return True

if __name__ == '__main__':
    main()
