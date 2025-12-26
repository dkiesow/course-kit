from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
import sqlite3
import json
from datetime import datetime
import os
import subprocess

app = Flask(__name__)
DB_PATH = 'presentations.db'

@app.route('/assets/<path:filename>')
def serve_asset(filename):
    return send_from_directory('assets', filename)

@app.route('/<path:filename>')
def serve_html(filename):
    if filename.endswith('.html'):
        return send_from_directory('.', filename)
    if filename.endswith('.css'):
        return send_from_directory('.', filename)
    return '', 404

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Presentations table
    c.execute('''CREATE TABLE IF NOT EXISTS presentations
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT NOT NULL,
                  front_matter TEXT,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    
    # Decks table (groups of slides separated by title slides)
    c.execute('''CREATE TABLE IF NOT EXISTS decks
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  presentation_id INTEGER,
                  week TEXT,
                  date TEXT,
                  order_index INTEGER,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  FOREIGN KEY (presentation_id) REFERENCES presentations(id) ON DELETE CASCADE)''')
    
    # Slides table
    c.execute('''CREATE TABLE IF NOT EXISTS slides
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  deck_id INTEGER,
                  slide_class TEXT,
                  headline TEXT,
                  paragraph TEXT,
                  bullets TEXT,
                  quote TEXT,
                  quote_citation TEXT,
                  image_path TEXT,
                  order_index INTEGER,
                  is_title BOOLEAN DEFAULT 0,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  FOREIGN KEY (deck_id) REFERENCES decks(id) ON DELETE CASCADE)''')
    
    conn.commit()
    conn.close()

init_db()

@app.route('/')
def index():
    return render_template('editor.html')

@app.route('/api/presentations', methods=['GET', 'POST'])
def presentations():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    if request.method == 'GET':
        c.execute('SELECT id, name, created_at, updated_at FROM presentations ORDER BY updated_at DESC')
        presentations = [{'id': row[0], 'name': row[1], 'created_at': row[2], 'updated_at': row[3]} 
                        for row in c.fetchall()]
        conn.close()
        return jsonify(presentations)
    
    elif request.method == 'POST':
        data = request.json
        c.execute('INSERT INTO presentations (name, front_matter) VALUES (?, ?)',
                 (data.get('name', 'New Presentation'), data.get('front_matter', '')))
        presentation_id = c.lastrowid
        conn.commit()
        conn.close()
        return jsonify({'id': presentation_id})

@app.route('/api/presentations/<int:presentation_id>', methods=['GET', 'PUT', 'DELETE'])
def presentation(presentation_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    if request.method == 'GET':
        c.execute('SELECT * FROM presentations WHERE id = ?', (presentation_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return jsonify({'error': 'Not found'}), 404
        
        # Get all decks with their slides
        c.execute('SELECT id, week, date, order_index FROM decks WHERE presentation_id = ? ORDER BY order_index', 
                 (presentation_id,))
        decks = []
        for deck_row in c.fetchall():
            deck_id = deck_row[0]
            c.execute('''SELECT id, slide_class, headline, paragraph, bullets, quote, quote_citation, 
                        image_path, order_index, is_title FROM slides 
                        WHERE deck_id = ? ORDER BY order_index''', (deck_id,))
            slides = [{'id': s[0], 'slideClass': s[1], 'headline': s[2], 'paragraph': s[3], 
                      'bullets': json.loads(s[4]) if s[4] else [], 'quote': s[5], 
                      'quoteCitation': s[6], 'imagePath': s[7], 'orderIndex': s[8], 'isTitle': bool(s[9])}
                     for s in c.fetchall()]
            decks.append({'id': deck_id, 'week': deck_row[1], 'date': deck_row[2], 
                         'orderIndex': deck_row[3], 'slides': slides})
        
        result = {
            'id': row[0],
            'name': row[1],
            'frontMatter': row[2],
            'decks': decks
        }
        conn.close()
        return jsonify(result)
    
    elif request.method == 'PUT':
        data = request.json
        c.execute('UPDATE presentations SET name = ?, front_matter = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                 (data.get('name'), data.get('frontMatter'), presentation_id))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    
    elif request.method == 'DELETE':
        c.execute('DELETE FROM presentations WHERE id = ?', (presentation_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})

@app.route('/api/decks', methods=['POST'])
def create_deck():
    data = request.json
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Get max order_index
    c.execute('SELECT MAX(order_index) FROM decks WHERE presentation_id = ?', 
             (data['presentation_id'],))
    max_order = c.fetchone()[0] or -1
    
    c.execute('INSERT INTO decks (presentation_id, week, date, order_index) VALUES (?, ?, ?, ?)',
             (data['presentation_id'], data.get('week', ''), data.get('date', ''), max_order + 1))
    deck_id = c.lastrowid
    
    # Automatically create a title slide for this deck
    c.execute('''INSERT INTO slides (deck_id, slide_class, headline, paragraph, bullets, 
                quote, quote_citation, image_path, order_index, is_title) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
             (deck_id, 'title', '', '', '[]', '', '', '', 0, True))
    
    conn.commit()
    conn.close()
    return jsonify({'id': deck_id})

@app.route('/api/decks/<int:deck_id>', methods=['PUT', 'DELETE'])
def deck(deck_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    if request.method == 'PUT':
        data = request.json
        c.execute('UPDATE decks SET week = ?, date = ? WHERE id = ?',
                 (data.get('week'), data.get('date'), deck_id))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    
    elif request.method == 'DELETE':
        c.execute('DELETE FROM decks WHERE id = ?', (deck_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})

@app.route('/api/slides', methods=['POST'])
def create_slide():
    data = request.json
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Get max order_index
    c.execute('SELECT MAX(order_index) FROM slides WHERE deck_id = ?', (data['deck_id'],))
    max_order = c.fetchone()[0] or -1
    
    c.execute('''INSERT INTO slides (deck_id, slide_class, headline, paragraph, bullets, 
                quote, quote_citation, image_path, order_index, is_title) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
             (data['deck_id'], data.get('class', ''), data.get('headline', ''),
              data.get('paragraph', ''), json.dumps(data.get('bullets', [])),
              data.get('quote', ''), data.get('quoteCitation', ''),
              data.get('imagePath', ''), max_order + 1, data.get('isTitle', False)))
    slide_id = c.lastrowid
    conn.commit()
    conn.close()
    return jsonify({'id': slide_id})

@app.route('/api/slides/<int:slide_id>', methods=['PUT', 'DELETE'])
def slide(slide_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    if request.method == 'PUT':
        data = request.json
        
        # Build dynamic UPDATE query based on provided fields
        update_fields = []
        update_values = []
        
        if 'slideClass' in data:
            update_fields.append('slide_class = ?')
            update_values.append(data.get('slideClass'))
        if 'headline' in data:
            update_fields.append('headline = ?')
            update_values.append(data.get('headline'))
        if 'paragraph' in data:
            update_fields.append('paragraph = ?')
            update_values.append(data.get('paragraph'))
        if 'bullets' in data:
            update_fields.append('bullets = ?')
            update_values.append(json.dumps(data.get('bullets', [])))
        if 'quote' in data:
            update_fields.append('quote = ?')
            update_values.append(data.get('quote'))
        if 'quoteCitation' in data:
            update_fields.append('quote_citation = ?')
            update_values.append(data.get('quoteCitation'))
        if 'imagePath' in data:
            update_fields.append('image_path = ?')
            update_values.append(data.get('imagePath'))
        if 'deck_id' in data:
            update_fields.append('deck_id = ?')
            update_values.append(data.get('deck_id'))
        
        if update_fields:
            update_values.append(slide_id)
            c.execute(f"UPDATE slides SET {', '.join(update_fields)} WHERE id = ?", tuple(update_values))
            conn.commit()
        
        conn.close()
        return jsonify({'success': True})
    
    elif request.method == 'DELETE':
        c.execute('DELETE FROM slides WHERE id = ?', (slide_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})

@app.route('/api/slides/reorder', methods=['POST'])
def reorder_slides():
    data = request.json
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    for item in data['slides']:
        c.execute('UPDATE slides SET order_index = ? WHERE id = ?',
                 (item['orderIndex'], item['id']))
    
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/decks/<int:deck_id>/export', methods=['GET'])
def export_deck(deck_id):
    """Export a specific deck to PDF or PPTX format"""
    try:
        format_type = request.args.get('format', 'pdf').lower()
        
        if format_type not in ['pdf', 'pptx']:
            return jsonify({'error': 'Invalid format. Use pdf or pptx'}), 400
        
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        
        # Get deck info
        c.execute('SELECT week, date, presentation_id FROM decks WHERE id = ?', (deck_id,))
        deck_row = c.fetchone()
        if not deck_row:
            conn.close()
            return jsonify({'error': 'Deck not found'}), 404
        
        week, date, presentation_id = deck_row
        
        # Get presentation front matter
        c.execute('SELECT front_matter FROM presentations WHERE id = ?', (presentation_id,))
        pres_row = c.fetchone()
        
        # Build markdown content for this deck only
        content = '''---
marp: true
theme: classroom
paginate: true
COURSE_TITLE: "Journalism Innovation"
---

'''
        
        # Get all slides for this deck
        c.execute('''SELECT slide_class, headline, paragraph, bullets, quote, quote_citation, 
                            image_path, is_title
                     FROM slides 
                     WHERE deck_id = ?
                     ORDER BY order_index''', (deck_id,))
        
        slides = c.fetchall()
        conn.close()
        
        for slide in slides:
            slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title = slide
            
            # Add slide separator
            content += '\n---\n\n'
            
            # Add class
            if slide_class:
                content += f'<!-- _class: {slide_class} -->\n'
            
            # Determine template type
            is_quote_template = slide_class and 'quote' in slide_class
            is_image_template = slide_class and 'image' in slide_class
            is_text_only = slide_class and 'lines' in slide_class
            
            if is_quote_template:
                # Quote templates: only export quote and citation
                if quote:
                    content += f'> {quote}\n'
                    if quote_citation:
                        content += f'>\n> {quote_citation}\n'
            else:
                # Non-quote templates: export headline, paragraph, bullets
                if headline:
                    content += f'# {headline}\n'
                
                if is_text_only:
                    # Text-only template: show paragraph, skip bullets
                    if paragraph:
                        content += f'\n{process_paragraph_linebreaks(paragraph)}\n'
                else:
                    # Bullet templates: show paragraph (if any) then bullets
                    if paragraph:
                        content += f'\n{process_paragraph_linebreaks(paragraph)}\n'
                    
                    if bullets:
                        try:
                            bullet_list = json.loads(bullets)
                            content += '\n'
                            for bullet in bullet_list:
                                if bullet.strip():
                                    content += f'- {bullet}\n'
                        except json.JSONDecodeError:
                            pass
                
                # Add image if present
                if image_path and is_image_template:
                    # Convert absolute web paths to relative filesystem paths
                    # /assets/cover.png -> ../assets/cover.png (relative from output/ directory)
                    if image_path.startswith('/assets/'):
                        image_path = '../assets/' + image_path[8:]  # Remove '/assets/' and add '../assets/'
                    elif image_path.startswith('assets/'):
                        image_path = '../' + image_path  # Add '../' prefix
                    content += f'\n![Image]({image_path})\n'
        
        # Write temporary markdown file
        # Sanitize filename - replace slashes and other problematic characters
        safe_week = str(week).replace('/', '-')
        safe_date = str(date).replace('/', '-').replace(' ', '_')
        
        temp_md = f'output/deck_{deck_id}_temp.md'
        output_file = f'output/Week_{safe_week}_{safe_date}.{format_type}'
        download_name = f'Week_{safe_week}_{safe_date}.{format_type}'
        
        with open(temp_md, 'w') as f:
            f.write(content)
        
        # Run Marp CLI to convert
        if format_type == 'pdf':
            cmd = f'marp "{temp_md}" -o "{output_file}" --allow-local-files --pdf --theme presentation-styles.css'
        else:  # pptx
            cmd = f'marp "{temp_md}" -o "{output_file}" --allow-local-files --pptx --theme presentation-styles.css'
        
        print(f"Running command: {cmd}")
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        print(f"Return code: {result.returncode}")
        print(f"Stdout: {result.stdout}")
        print(f"Stderr: {result.stderr}")
        
        # Clean up temp file
        if os.path.exists(temp_md):
            os.remove(temp_md)
        
        if result.returncode != 0:
            print(f"Marp error: {result.stderr}")
            return jsonify({'error': f'Marp conversion failed: {result.stderr}'}), 500
        
        # Check if output file was created
        if not os.path.exists(output_file):
            return jsonify({'error': f'Output file was not created: {output_file}'}), 500
        
        # Send file
        return send_file(
            output_file,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/pdf' if format_type == 'pdf' else 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        print(f"Export error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

def process_paragraph_linebreaks(text):
    """Convert single linebreaks to <br> tags for proper Markdown rendering"""
    if not text:
        return text
    # Split into paragraphs (separated by blank lines)
    paragraphs = text.split('\n\n')
    processed_paragraphs = []
    
    for para in paragraphs:
        # Within each paragraph, join lines with <br>
        lines = [line.strip() for line in para.split('\n') if line.strip()]
        if lines:
            processed_paragraphs.append('<br>'.join(lines))
    
    # Join paragraphs with double newlines to create spacing
    return '\n\n'.join(processed_paragraphs)


@app.route('/api/presentations/<int:presentation_id>/auto-export', methods=['POST'])
def auto_export_presentation(presentation_id):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    
    # Get presentation
    c.execute('SELECT * FROM presentations WHERE id = ?', (presentation_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Not found'}), 404
    
    # Build proper front matter with Marp configuration
    content = '''---
marp: true
theme: classroom
paginate: true
COURSE_TITLE: "Journalism Innovation"
---

'''
    
    # Get all decks and slides
    c.execute('SELECT id, week, date FROM decks WHERE presentation_id = ? ORDER BY order_index', 
             (presentation_id,))
    decks = c.fetchall()
    
    slides_markdown = []
    
    for deck in decks:
        deck_id = deck[0]
        week = deck[1]
        date = deck[2]
        
        c.execute('''SELECT slide_class, headline, paragraph, bullets, quote, quote_citation, 
                    image_path, is_title FROM slides WHERE deck_id = ? ORDER BY order_index''', 
                 (deck_id,))
        
        for slide in c.fetchall():
            slide_class = slide[0]
            headline = slide[1]
            paragraph = slide[2]
            bullets = json.loads(slide[3]) if slide[3] else []
            quote = slide[4]
            quote_citation = slide[5]
            image_path = slide[6]
            is_title = slide[7]
            
            slide_md = ''
            
            if is_title:
                # Title slide with week/date variables
                slide_md = f'<!--\nWEEK: "{week}"\nDATE: "{date}"\n_class: title\n-->\n'
                slide_md += f'# {{{{COURSE_TITLE}}}}\n'
                slide_md += f'## {{{{WEEK}}}}\n'
                slide_md += f'{{{{DATE}}}}\n'
            else:
                slide_md = f'<!-- _class: {slide_class} -->\n'
                
                # Determine what fields to export based on template type
                is_quote_template = slide_class and 'quote' in slide_class
                is_image_template = slide_class and 'image' in slide_class
                is_text_only = slide_class and 'lines' in slide_class
                
                if is_quote_template:
                    # Quote templates: only export quote and citation
                    if quote:
                        slide_md += f'> {quote}\n'
                        if quote_citation:
                            slide_md += f'>\n> {quote_citation}\n'
                else:
                    # Non-quote templates: export headline, paragraph, bullets
                    if headline:
                        slide_md += f'# {headline}\n'
                    
                    if is_text_only:
                        # Text-only template: show paragraph, skip bullets
                        if paragraph:
                            slide_md += f'\n{process_paragraph_linebreaks(paragraph)}\n'
                    else:
                        # Bullet templates: show paragraph (if any) then bullets
                        if paragraph:
                            slide_md += f'\n{process_paragraph_linebreaks(paragraph)}\n'
                        
                        if bullets:
                            slide_md += '\n'
                            if slide_class and '2col' in slide_class:
                                midpoint = (len(bullets) + 1) // 2
                                slide_md += '<div>\n\n'
                                for bullet in bullets[:midpoint]:
                                    slide_md += f'- {bullet}\n'
                                slide_md += '\n</div>\n<div>\n\n'
                                for bullet in bullets[midpoint:]:
                                    slide_md += f'- {bullet}\n'
                                slide_md += '\n</div>\n'
                            else:
                                for bullet in bullets:
                                    slide_md += f'- {bullet}\n'
                    
                    # Only export image if template supports images
                    if is_image_template and image_path:
                        slide_md += f'\n![Image]({image_path})\n'
            
            slides_markdown.append(slide_md)
    
    content += '\n\n---\n\n'.join(slides_markdown)
    
    # Save to file
    filename = f"{row[1].replace(' ', '_')}.md"
    filepath = os.path.join(os.path.dirname(__file__), filename)
    with open(filepath, 'w') as f:
        f.write(content)
    
    conn.close()
    return jsonify({'success': True, 'file': filename})

@app.route('/api/presentations/<int:presentation_id>/preview', methods=['POST'])
def preview_presentation(presentation_id):
    # Use auto_export_presentation to ensure consistency
    result = auto_export_presentation(presentation_id)
    if isinstance(result, tuple):  # Error response
        return result
    
    # Now export specifically to presentation.md for build process
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    
    c.execute('SELECT * FROM presentations WHERE id = ?', (presentation_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Not found'}), 404
    
    # Get the content that was just exported
    filename = f"{row[1].replace(' ', '_')}.md"
    source_filepath = os.path.join(os.path.dirname(__file__), filename)
    
    # Read and copy to presentation.md
    if os.path.exists(source_filepath):
        with open(source_filepath, 'r') as f:
            content = f.read()
        
        presentation_path = os.path.join(os.path.dirname(__file__), 'presentation.md')
        with open(presentation_path, 'w') as f:
            f.write(content)
    
    conn.close()
    
    # Run build.js to preprocess
    try:
        result = subprocess.run(['npm', 'run', 'build'], 
                              capture_output=True, 
                              text=True, 
                              cwd=os.path.dirname(__file__))
        if result.returncode != 0:
            return jsonify({'error': 'Build failed', 'output': result.stderr}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
    return jsonify({'success': True, 'file': 'output/presentation.html'})

@app.route('/api/presentations/<int:presentation_id>/export', methods=['GET'])
def export_presentation(presentation_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    c.execute('SELECT name, front_matter FROM presentations WHERE id = ?', (presentation_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Not found'}), 404
    
    # Build proper front matter with Marp configuration
    content = '''---
marp: true
theme: classroom
paginate: true
COURSE_TITLE: "Journalism Innovation"
---

'''
    
    # Get all decks and slides
    c.execute('SELECT id, week, date FROM decks WHERE presentation_id = ? ORDER BY order_index', 
             (presentation_id,))
    decks = c.fetchall()
    
    slides_markdown = []
    
    for deck in decks:
        deck_id = deck[0]
        week = deck[1]
        date = deck[2]
        
        c.execute('''SELECT slide_class, headline, paragraph, bullets, quote, quote_citation, 
                    image_path, is_title FROM slides WHERE deck_id = ? ORDER BY order_index''', 
                 (deck_id,))
        
        for slide in c.fetchall():
            slide_class = slide[0]
            headline = slide[1]
            paragraph = slide[2]
            bullets = json.loads(slide[3]) if slide[3] else []
            quote = slide[4]
            quote_citation = slide[5]
            image_path = slide[6]
            is_title = slide[7]
            
            slide_md = ''
            
            if is_title:
                # Title slide with week/date variables
                slide_md = f'<!--\nWEEK: "{week}"\nDATE: "{date}"\n_class: title\n-->\n'
                slide_md += f'# {{{{COURSE_TITLE}}}}\n'
                slide_md += f'## {{{{WEEK}}}}\n'
                slide_md += f'{{{{DATE}}}}\n'
            else:
                slide_md = f'<!-- _class: {slide_class} -->\n'
                
                # Determine what fields to export based on template type
                is_quote_template = slide_class and 'quote' in slide_class
                is_image_template = slide_class and 'image' in slide_class
                is_text_only = slide_class and 'lines' in slide_class
                
                if is_quote_template:
                    # Quote templates: only export quote and citation
                    if quote:
                        slide_md += f'> {quote}\n'
                        if quote_citation:
                            slide_md += f'>\n> {quote_citation}\n'
                else:
                    # Non-quote templates: export headline, paragraph, bullets
                    if headline:
                        slide_md += f'# {headline}\n'
                    
                    if is_text_only:
                        # Text-only template: show paragraph, skip bullets
                        if paragraph:
                            slide_md += f'\n{process_paragraph_linebreaks(paragraph)}\n'
                    else:
                        # Bullet templates: show paragraph (if any) then bullets
                        if paragraph:
                            slide_md += f'\n{process_paragraph_linebreaks(paragraph)}\n'
                        
                        if bullets:
                            slide_md += '\n'
                            if slide_class and '2col' in slide_class:
                                midpoint = (len(bullets) + 1) // 2
                                slide_md += '<div>\n\n'
                                for bullet in bullets[:midpoint]:
                                    slide_md += f'- {bullet}\n'
                                slide_md += '\n</div>\n<div>\n\n'
                                for bullet in bullets[midpoint:]:
                                    slide_md += f'- {bullet}\n'
                                slide_md += '\n</div>\n'
                            else:
                                for bullet in bullets:
                                    slide_md += f'- {bullet}\n'
                    
                    # Only export image if template supports images
                    if is_image_template and image_path:
                        slide_md += f'\n![Image]({image_path})\n'
            
            slides_markdown.append(slide_md)
    
    content += '\n\n---\n\n'.join(slides_markdown)
    
    conn.close()
    
    # Save to file
    filename = f"{row[0].replace(' ', '_')}.md"
    filepath = f"/tmp/{filename}"
    with open(filepath, 'w') as f:
        f.write(content)
    
    return send_file(filepath, as_attachment=True, download_name=filename)

@app.route('/api/presentations/import', methods=['POST'])
def import_presentation():
    data = request.json
    markdown_content = data.get('content', '')
    filename = data.get('filename', 'Imported Presentation')
    
    # Parse markdown content
    slides_raw = markdown_content.split('---')
    
    conn = sqlite3.connect('presentations.db')
    c = conn.cursor()
    
    # Create presentation
    front_matter = slides_raw[0].strip() if slides_raw else ''
    name = filename.replace('.md', '')
    
    c.execute('INSERT INTO presentations (name, front_matter) VALUES (?, ?)',
              (name, front_matter))
    presentation_id = c.lastrowid
    
    # Parse slides and organize into decks
    current_deck_id = None
    deck_order = 0
    slide_order = 0
    
    for i, slide_content in enumerate(slides_raw[1:], 1):  # Skip front matter
        lines = slide_content.strip().split('\n')
        if not lines:
            continue
        
        # Parse slide class
        slide_class = ''
        content_start = 0
        for idx, line in enumerate(lines):
            stripped = line.strip()
            # Handle both formats: "class: template" and "<!-- _class: template -->"
            if stripped.startswith('class:'):
                slide_class = stripped.replace('class:', '').strip()
                content_start = idx + 1
                break
            elif stripped.startswith('<!--') and '_class:' in stripped:
                # Extract class from <!-- _class: gold-quote-headline -->
                import re
                match = re.search(r'_class:\s*([\w-]+)', stripped)
                if match:
                    slide_class = match.group(1)
                content_start = idx + 1
                break
        
        # Check if title slide
        is_title = slide_class == 'title'
        
        # Extract content
        headline = ''
        paragraph = ''
        bullets = []
        quote = ''
        quote_citation = ''
        image_path = ''
        week = ''
        date = ''
        
        # Parse in order: headline, then paragraph, then bullets/quote
        j = content_start
        headline_found = False
        paragraph_lines = []
        quote_lines = []
        in_quote = False
        
        while j < len(lines):
            line = lines[j].strip()
            
            # Skip empty lines (but continue quote if already in one)
            if not line:
                if in_quote:
                    quote_lines.append('')  # Preserve empty lines in quotes
                j += 1
                continue
            
            # Extract headline
            if line.startswith('# '):
                headline = line[2:]
                headline_found = True
                j += 1
                continue
            
            # Title slide specific fields
            if line.startswith('WEEK:'):
                week = line.replace('WEEK:', '').strip()
                j += 1
                continue
            elif line.startswith('DATE:'):
                date = line.replace('DATE:', '').strip()
                j += 1
                continue
            
            # Bullets
            if line.startswith('- '):
                bullets.append(line[2:])
                j += 1
                continue
            
            # Quote (can be multiline, may include citation)
            if line.startswith('> '):
                content = line[2:]
                # Check if this is a citation line (starts with em dash)
                if content.startswith('—'):
                    quote_citation = content[1:].strip()
                    in_quote = False
                else:
                    in_quote = True
                    quote_lines.append(content)
                j += 1
                continue
            elif line.startswith('>'):
                content = line[1:].strip()
                if content:  # Non-empty line after >
                    if content.startswith('—'):
                        quote_citation = content[1:].strip()
                        in_quote = False
                    else:
                        in_quote = True
                        quote_lines.append(content)
                else:
                    # Empty quote line (just >)
                    if in_quote:
                        quote_lines.append('')
                j += 1
                continue
            
            # Citation not in quote block (marks end of quote)
            if line.startswith('—'):
                quote_citation = line[1:].strip()
                in_quote = False
                j += 1
                continue
            
            # Image
            if line.startswith('!['):
                import re
                match = re.search(r'\((.+?)\)', line)
                if match:
                    image_path = match.group(1)
                j += 1
                continue
            
            # Paragraph text (after headline, before bullets/quote)
            if headline_found and not bullets and not in_quote and not quote_lines:
                paragraph_lines.append(line)
            
            j += 1
        
        paragraph = '\n'.join(paragraph_lines).strip()
        quote = '\n'.join(quote_lines).strip()
        
        # If title slide, create new deck
        if is_title:
            c.execute('INSERT INTO decks (presentation_id, week, date, order_index) VALUES (?, ?, ?, ?)',
                     (presentation_id, week or f'Week {deck_order + 1}', date or 'TBD', deck_order))
            current_deck_id = c.lastrowid
            deck_order += 1
            slide_order = 0
            
            # Insert title slide
            c.execute('''INSERT INTO slides 
                        (deck_id, slide_class, headline, order_index, is_title) 
                        VALUES (?, ?, ?, ?, ?)''',
                     (current_deck_id, slide_class, headline, slide_order, 1))
        else:
            # Regular slide - need a deck
            if current_deck_id is None:
                # Create default deck if none exists
                c.execute('INSERT INTO decks (presentation_id, week, date, order_index) VALUES (?, ?, ?, ?)',
                         (presentation_id, 'Week 1', 'TBD', deck_order))
                current_deck_id = c.lastrowid
                deck_order += 1
                slide_order = 0
            
            # Insert slide
            c.execute('''INSERT INTO slides 
                        (deck_id, slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, order_index, is_title) 
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                     (current_deck_id, slide_class, headline, paragraph, 
                      json.dumps(bullets), quote, quote_citation, image_path, slide_order, 0))
        
        slide_order += 1
    
    conn.commit()
    conn.close()
    
    return jsonify({'id': presentation_id, 'name': name}), 201

if __name__ == '__main__':
    app.run(debug=True, port=5001)
