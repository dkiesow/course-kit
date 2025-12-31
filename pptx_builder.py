#!/usr/bin/env python3
"""
Custom PPTX builder using python-pptx to work with custom layout names.
This replaces pandoc for PPTX generation.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
import json
import os


def build_pptx_from_slides(slides_data, output_path, template_path, pptx_layouts_map, deck_info=None):
    """
    Build a PPTX file directly from slide data using custom layouts.
    
    Args:
        slides_data: List of tuples containing slide data from database
        output_path: Path where PPTX should be saved
        template_path: Path to POTX/PPTX template file  
        pptx_layouts_map: Dict mapping template_base to layout names
        deck_info: Dict with course_title, week, date for title slide
    """
    import zipfile
    import shutil
    
    # python-pptx doesn't support POTX files, so we need to convert it
    temp_pptx = 'temp_template.pptx'
    
    # Remove old temp file if it exists to ensure we use latest template
    if os.path.exists(temp_pptx):
        os.remove(temp_pptx)
    
    # Copy the template
    shutil.copy2(template_path, temp_pptx)
    
    # Modify the content type in [Content_Types].xml to make it a presentation
    try:
        # Read the ZIP
        with zipfile.ZipFile(temp_pptx, 'r') as zf:
            files = {name: zf.read(name) for name in zf.namelist()}
        
        # Modify content types
        if '[Content_Types].xml' in files:
            content_xml = files['[Content_Types].xml'].decode('utf-8')
            content_xml = content_xml.replace(
                'presentationml.template.main',
                'presentationml.presentation.main'
            )
            # Add GIF support if not already present
            if 'image/gif' not in content_xml:
                # Insert GIF content type after PNG
                content_xml = content_xml.replace(
                    '<Default Extension="png" ContentType="image/png"/>',
                    '<Default Extension="png" ContentType="image/png"/><Default Extension="gif" ContentType="image/gif"/>'
                )
                print("Added GIF content type support to template")
            files['[Content_Types].xml'] = content_xml.encode('utf-8')
        
        # Write back
        with zipfile.ZipFile(temp_pptx, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, data in files.items():
                zf.writestr(name, data)
    except Exception as e:
        print(f"Warning: Could not patch template: {e}")
    
    # Now load with python-pptx
    prs = Presentation(temp_pptx)
    
    # Build a map of layout names to layout objects
    # Must iterate through all slide_masters since prs.slide_layouts only returns first master's layouts
    layout_map = {}
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            layout_map[layout.name] = layout
    
    print(f"Available layouts: {list(layout_map.keys())}")
    
    # Process each slide
    for slide_data in slides_data:
        (slide_class, headline, paragraph, bullets, quote, quote_citation, 
         image_path, is_title, hide_headline, larger_image, fullscreen, template_base) = slide_data
        
        # Determine layout name directly from slide_class (no guessing)
        if is_title:
            layout_name = "Arches_Title"
        else:
            # Use slide_class if it exists, otherwise template_base
            template_key = slide_class if slide_class else template_base
            layout_name = pptx_layouts_map.get(template_key)
            
            # For photo-centered templates, use headline version if hide_headline is False
            if template_key == 'photo-centered' and not hide_headline:
                layout_name = "White_Photo_Headline"
            elif template_key == 'gold-photo-centered' and not hide_headline:
                layout_name = "Gold_Photo_Headline"
            # For bullets-image-top templates, use big photo version if larger_image is True
            elif template_key == 'bullets-image-top' and larger_image:
                layout_name = "White_Top_Bullets_Big_Photo"
            elif template_key == 'gold-bullets-image-top' and larger_image:
                layout_name = "Gold_Top_Bullets_Big_Photo"
            
            if not layout_name:
                print(f"Warning: No layout mapping for template '{template_key}'")
                layout_name = "White_Bullets"
        
        # Get the layout
        layout = layout_map.get(layout_name)
        if not layout:
            print(f"Warning: Layout '{layout_name}' not found, using first available layout")
            layout = prs.slide_layouts[0]
        
        # Add slide with the layout
        slide = prs.slides.add_slide(layout)
        
        # Populate placeholders based on slide type
        if is_title:
            populate_title_slide(slide, headline, paragraph, deck_info)
        elif template_base == "closing":
            populate_closing_slide(slide, headline, paragraph, image_path)
        elif template_base in ["photo-centered", "gold-photo-centered"]:
            populate_photo_slide(slide, image_path, headline, hide_headline)
        elif template_base in ["quote", "gold-quote"]:
            populate_quote_slide(slide, quote, quote_citation)
        elif template_base in ["bullets-image", "bullets-image-split", "bullets-image-top", "gold-bullets-image-split", "gold-bullets-image-top"]:
            # Image layouts - add image
            populate_content_slide(slide, headline, paragraph, bullets, image_path, hide_headline, slide_class)
        else:
            # Regular content - no image unless it's explicitly an image layout
            populate_content_slide(slide, headline, paragraph, bullets, None, hide_headline, slide_class)
    
    # Save the presentation
    prs.save(output_path)
    return True


def add_formatted_text_to_frame(text_frame, text):
    """Add text with markdown formatting to a text frame."""
    import re
    
    # Clear existing paragraphs but keep the first one for formatting
    if text_frame.paragraphs:
        text_frame.paragraphs[0].text = ""
        while len(text_frame.paragraphs) > 1:
            elem = text_frame.paragraphs[1]._element
            elem.getparent().remove(elem)
    
    # Split text into lines and process each line
    lines = text.split('\n')
    
    for line_idx, line in enumerate(lines):
        if line.strip():  # Only process non-empty lines
            # Get or create paragraph
            if line_idx == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            # Parse markdown formatting in the line
            parse_markdown_to_paragraph(p, line)
        elif line_idx > 0:  # Add empty line breaks (but not at the beginning)
            text_frame.add_paragraph()


def parse_markdown_to_paragraph(paragraph, text):
    """Parse markdown formatting and add runs to paragraph."""
    import re
    
    # Pattern to match **bold** and *italic* markdown
    pattern = r'(\*\*.*?\*\*|\*.*?\*|[^*]+|\*)'
    parts = re.findall(pattern, text)
    
    for part in parts:
        if not part:
            continue
            
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            # Bold text
            run = paragraph.add_run()
            run.text = part[2:-2]  # Remove ** markers
            run.font.bold = True
        elif part.startswith('*') and part.endswith('*') and len(part) > 2 and not part.startswith('**'):
            # Italic text
            run = paragraph.add_run()
            run.text = part[1:-1]  # Remove * markers
            run.font.italic = True
        else:
            # Regular text
            run = paragraph.add_run()
            run.text = part


def populate_title_slide(slide, headline, paragraph, deck_info=None):
    """Populate title slide with course name, week, and date."""
    # Get deck information
    course_title = deck_info.get('course_title', 'Journalism Innovation') if deck_info else 'Journalism Innovation'
    week = deck_info.get('week', '') if deck_info else ''
    date = deck_info.get('date', '') if deck_info else ''
    
    # Find the three placeholders (should be course title, week, date)
    placeholders = [shape for shape in slide.placeholders if shape.placeholder_format.type in [1, 2, 7]]  # Title, Body, Object
    
    if len(placeholders) >= 3:
        # Placeholder 1: Course Title
        if placeholders[0].has_text_frame:
            placeholders[0].text = course_title
        
        # Placeholder 2: Week
        if placeholders[1].has_text_frame:
            placeholders[1].text = week if week else 'Week'
            
        # Placeholder 3: Date  
        if placeholders[2].has_text_frame:
            placeholders[2].text = date if date else 'Date'
    else:
        # Fallback to old behavior if not enough placeholders
        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 1:  # Title
                if shape.has_text_frame:
                    shape.text = course_title
                break


def populate_quote_slide(slide, quote, quote_citation):
    """Populate quote slide with quote text and citation."""
    # Find text placeholders - skip title placeholder (idx 0), use content placeholder
    for shape in slide.shapes:
        if shape.has_text_frame:
            # Skip title placeholders
            if hasattr(shape, 'placeholder_format') and shape.placeholder_format.type == 1:  # 1 = title
                continue
            
            text_frame = shape.text_frame
            text_frame.word_wrap = True
            # Clear placeholder text but preserve formatting
            if text_frame.paragraphs:
                text_frame.paragraphs[0].text = ""
                while len(text_frame.paragraphs) > 1:
                    elem = text_frame.paragraphs[1]._element
                    elem.getparent().remove(elem)
            if quote:
                p = text_frame.paragraphs[0] if text_frame.paragraphs else text_frame.add_paragraph()
                # Parse markdown formatting for quote
                parse_markdown_to_paragraph(p, quote)
                # Explicitly disable bullets and remove hanging indent
                p._element.get_or_add_pPr().buNone = None
                from lxml import etree
                buNone = etree.SubElement(p._element.get_or_add_pPr(), '{http://schemas.openxmlformats.org/drawingml/2006/main}buNone')
                # Remove hanging indent - set both first line and left indent to 0
                p.level = 0
                pPr = p._element.get_or_add_pPr()
                pPr.set('indent', '0')
                pPr.set('marL', '0')
                
                if quote_citation:
                    # Strip leading bullet markers from citation
                    citation_text = quote_citation.lstrip('- ').strip()
                    # Add citation as a new paragraph
                    p = text_frame.add_paragraph()
                    # Disable bullets BEFORE setting text
                    buNone = etree.SubElement(p._element.get_or_add_pPr(), '{http://schemas.openxmlformats.org/drawingml/2006/main}buNone')
                    p.text = f"\n— {citation_text}"
                    p.font.italic = True
                break


def populate_closing_slide(slide, headline, paragraph, image_path):
    """Populate closing slide with name and optional image."""
    # Find title placeholder
    title_placeholder = None
    content_placeholder = None
    
    for shape in slide.placeholders:
        ph_type = shape.placeholder_format.type
        if ph_type == 1:  # Title
            title_placeholder = shape
        elif ph_type in [2, 7]:  # Body or Object
            content_placeholder = shape
    
    if title_placeholder:
        if title_placeholder.has_text_frame:
            text_frame = title_placeholder.text_frame
            if text_frame.paragraphs:
                text_frame.paragraphs[0].text = ""
                while len(text_frame.paragraphs) > 1:
                    elem = text_frame.paragraphs[1]._element
                    elem.getparent().remove(elem)
        if headline:
            # Split headline by newlines, use first line as main text
            lines = headline.split('\n')
            title_placeholder.text = lines[0]
    
    # Add paragraph if present
    if content_placeholder:
        if content_placeholder.has_text_frame:
            content_placeholder.text_frame.clear()
        if paragraph:
            p = content_placeholder.text_frame.paragraphs[0] if content_placeholder.text_frame.paragraphs else content_placeholder.text_frame.add_paragraph()
            parse_markdown_to_paragraph(p, paragraph)
    
    # Add image if present
    if image_path:
        add_image_to_slide(slide, image_path)


def populate_photo_slide(slide, image_path, headline=None, hide_headline=True):
    """Populate photo-centered slide with large image and optional headline."""
    # Add headline if not hidden
    if headline and not hide_headline:
        title_shape = None
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = shape.placeholder_format.type
                if ph_type == 1:  # Title placeholder
                    title_shape = shape
                    break
        
        if title_shape and title_shape.has_text_frame:
            tf = title_shape.text_frame
            tf.paragraphs[0].text = ""
            parse_markdown_to_paragraph(tf.paragraphs[0], headline)
    
    # Add image
    if image_path:
        add_image_to_slide(slide, image_path, centered=True)


def populate_content_slide(slide, headline, paragraph, bullets, image_path, hide_headline, slide_class=None):
    """Populate standard content slide with headline, text, bullets, and optional image."""
    # Determine if this is a text-only slide
    is_text_only = slide_class and 'lines' in slide_class
    
    # Find title placeholder
    if not hide_headline and headline:
        title_shape = None
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = shape.placeholder_format.type
                if ph_type == 1:  # Title placeholder
                    title_shape = shape
                    break
        
        if title_shape and title_shape.has_text_frame:
            # Clear and add formatted headline
            tf = title_shape.text_frame
            tf.paragraphs[0].text = ""
            parse_markdown_to_paragraph(tf.paragraphs[0], headline)
    
    # Find content placeholder for bullets or paragraph
    # Try multiple placeholder types (2=Body, 7=Object, 14=Content)
    content_shape = None
    for shape in slide.shapes:
        if shape.is_placeholder and shape.has_text_frame:
            ph_type = shape.placeholder_format.type
            if ph_type in [2, 7, 14]:  # Body, Object, or Content placeholder
                content_shape = shape
                break
    
    # If no specific content placeholder found, try any text placeholder that's not the title
    if not content_shape:
        for shape in slide.shapes:
            if shape.has_text_frame and not (shape.is_placeholder and shape.placeholder_format.type == 1):
                if hasattr(shape, 'text_frame'):
                    content_shape = shape
                    break
    
    if content_shape and content_shape.has_text_frame:
        text_frame = content_shape.text_frame
        text_frame.word_wrap = True
        
        # Clear placeholder text but preserve first paragraph's formatting
        if text_frame.paragraphs:
            # Clear text from first paragraph but keep the paragraph (preserves template formatting)
            text_frame.paragraphs[0].text = ""
            # Remove extra paragraphs
            while len(text_frame.paragraphs) > 1:
                elem = text_frame.paragraphs[1]._element
                elem.getparent().remove(elem)
        
        # For text-only slides (like template-lines), prioritize paragraph content
        if is_text_only and paragraph:
            add_formatted_text_to_frame(text_frame, paragraph)
            print(f"    Adding formatted paragraph text for text-only slide: {paragraph[:50]}...")
        # For bullet slides, handle paragraph first, then bullets
        elif not is_text_only:
            paragraph_added = False
            
            # Add paragraph text first if present
            if paragraph:
                p = text_frame.paragraphs[0]
                parse_markdown_to_paragraph(p, paragraph)
                # Explicitly disable bullets for paragraph text
                from lxml import etree
                buNone = etree.SubElement(p._element.get_or_add_pPr(), '{http://schemas.openxmlformats.org/drawingml/2006/main}buNone')
                paragraph_added = True
                print(f"    Adding paragraph text: {paragraph[:50]}...")
            
            # Add bullets if present
            if bullets:
                try:
                    bullet_list = json.loads(bullets)
                    if bullet_list:  # Only add if list is not empty
                        print(f"    Adding {len(bullet_list)} bullets")
                        for i, bullet in enumerate(bullet_list):
                            if i == 0 and not paragraph_added:
                                # Use first paragraph if no paragraph text was added
                                p = text_frame.paragraphs[0]
                            else:
                                p = text_frame.add_paragraph()
                            
                            # Detect indent level from markdown-style -- prefix OR leading spaces/tabs
                            indent_level = 0
                            stripped_bullet = bullet.lstrip()
                            
                            # Check for markdown -- prefix for indentation
                            if stripped_bullet.startswith('--'):
                                # Count consecutive - characters starting from 2 (-- = level 1)
                                dash_count = len(stripped_bullet) - len(stripped_bullet.lstrip('-'))
                                indent_level = max(0, dash_count - 1)  # -- is level 1, --- is level 2, etc.
                                # Remove the - markers and any following space
                                stripped_bullet = stripped_bullet.lstrip('-').lstrip()
                            else:
                                # Fall back to space/tab detection
                                leading_whitespace = len(bullet) - len(stripped_bullet)
                                
                                # Calculate level: 2 spaces or 1 tab = 1 level
                                if '\t' in bullet[:leading_whitespace]:
                                    indent_level = bullet[:leading_whitespace].count('\t')
                                else:
                                    indent_level = leading_whitespace // 2
                            
                            # Cap at level 4 (PowerPoint supports up to 9 but 4 is reasonable)
                            indent_level = min(indent_level, 4)
                            
                            # Parse markdown formatting for bullet (use stripped text)
                            parse_markdown_to_paragraph(p, stripped_bullet)
                            p.level = indent_level
                except:
                    # If JSON parsing fails, treat as plain text
                    if not paragraph_added:
                        text_frame.text = bullets
                    else:
                        p = text_frame.add_paragraph()
                        p.text = bullets
        
        # Add paragraph if present (and no bullets handled above)
        elif paragraph:
            p = text_frame.paragraphs[0] if text_frame.paragraphs else text_frame.add_paragraph()
            parse_markdown_to_paragraph(p, paragraph)
            # Explicitly disable bullets for paragraph text
            from lxml import etree
            buNone = etree.SubElement(p._element.get_or_add_pPr(), '{http://schemas.openxmlformats.org/drawingml/2006/main}buNone')
    else:
        # Debug: print placeholder info
        print(f"Warning: No content placeholder found for slide. Available placeholders:")
        for shape in slide.shapes:
            if shape.is_placeholder:
                print(f"  Type: {shape.placeholder_format.type}, Has text frame: {shape.has_text_frame}")
    
    # Add image if present
    if image_path:
        add_image_to_slide(slide, image_path)


def add_image_to_slide(slide, image_path, centered=False):
    """Add an image to a slide, either in a picture placeholder or positioned."""
    from PIL import Image
    
    # Clean up image path
    if image_path.startswith('/assets/'):
        image_path = 'assets' + image_path[7:]
    elif image_path.startswith('assets/'):
        pass  # Already correct
    else:
        image_path = 'assets/' + image_path
    
    # Check if file exists
    if not os.path.exists(image_path):
        print(f"Warning: Image not found: {image_path}")
        return
    
    # Get image dimensions - handle GIFs specially to preserve animation
    if image_path.lower().endswith('.gif'):
        # For GIFs, use a simpler approach to avoid PIL processing that might strip animation
        try:
            from PIL import Image
            with Image.open(image_path) as img:
                img_width, img_height = img.size
                # Don't process the image further - let python-pptx handle it directly
        except Exception as e:
            print(f"Warning: Could not read GIF dimensions: {e}")
            # Use default dimensions if we can't read the file
            img_width, img_height = 800, 600
    else:
        # For other formats, use normal PIL processing
        from PIL import Image
        with Image.open(image_path) as img:
            img_width, img_height = img.size
    
    # Try to find a picture placeholder
    picture_placeholder = None
    for shape in slide.shapes:
        if shape.is_placeholder:
            ph_type = shape.placeholder_format.type
            if ph_type == 18:  # Picture placeholder
                picture_placeholder = shape
                break
    
    if picture_placeholder:
        # Get placeholder dimensions
        ph_width = picture_placeholder.width
        ph_height = picture_placeholder.height
        ph_left = picture_placeholder.left
        ph_top = picture_placeholder.top
        
        # Use the image dimensions already determined above
        img_aspect = img_width / img_height
        ph_aspect = ph_width / ph_height
        
        # Calculate size to fit within placeholder while maintaining aspect ratio
        if img_aspect > ph_aspect:
            # Image is wider - constrain by width
            new_width = ph_width
            new_height = int(ph_width / img_aspect)
        else:
            # Image is taller - constrain by height
            new_height = ph_height
            new_width = int(ph_height * img_aspect)
        
        # Center within placeholder
        left = ph_left + (ph_width - new_width) // 2
        top = ph_top + (ph_height - new_height) // 2
        
        # Remove the placeholder and add image in its place
        sp = picture_placeholder.element
        sp.getparent().remove(sp)
        slide.shapes.add_picture(image_path, left, top, width=new_width, height=new_height)
    else:
        # No placeholder - add at native size, centered on slide
        # Use the dimensions already determined above (img_width, img_height)
        img_width_px, img_height_px = img_width, img_height
        # Convert pixels to EMUs (English Metric Units): 1 inch = 914400 EMUs, assume 96 DPI
        dpi = 96
        img_width_emu = int(img_width_px * 914400 / dpi)
        img_height_emu = int(img_height_px * 914400 / dpi)
        
        # Get slide dimensions (standard is 10" x 7.5")
        slide_width = 9144000  # 10 inches in EMUs
        slide_height = 6858000  # 7.5 inches in EMUs
        
        # Center the image
        left = (slide_width - img_width_emu) // 2
        top = (slide_height - img_height_emu) // 2
        
        # Ensure image doesn't exceed slide bounds
        if img_width_emu > slide_width or img_height_emu > slide_height:
            # Scale down to fit
            scale = min(slide_width / img_width_emu, slide_height / img_height_emu) * 0.9
            img_width_emu = int(img_width_emu * scale)
            img_height_emu = int(img_height_emu * scale)
            left = (slide_width - img_width_emu) // 2
            top = (slide_height - img_height_emu) // 2
        
        slide.shapes.add_picture(image_path, left, top, width=img_width_emu, height=img_height_emu)
