from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
import sqlite3
import json
from datetime import datetime
import os
import subprocess
from pptx_builder import build_pptx_from_slides

app = Flask(__name__)
DB_PATH = 'presentations.db'

def substitute_assignment_variables(text, conn=None):
    """Replace {assignment:name} with 'Assignment Name - Due: Date'"""
    if not text or '{assignment:' not in text:
        return text
    
    # Create connection if not provided
    close_conn = False
    if not conn:
        conn = sqlite3.connect(DB_PATH)
        close_conn = True
    
    c = conn.cursor()
    
    import re
    # Find all {assignment:xxx} patterns
    pattern = r'\{assignment:([^}]+)\}'
    matches = re.findall(pattern, text)
    
    for assignment_name in matches:
        # Look up the assignment
        c.execute('SELECT name, due_date FROM assignments WHERE name = ?', (assignment_name,))
        row = c.fetchone()
        
        if row:
            # Parse date and format as "Month Day" (e.g., "May 7")
            due_date = row[1]  # Format: YYYY-MM-DD
            try:
                date_obj = datetime.strptime(due_date, '%Y-%m-%d')
                formatted_date = date_obj.strftime('%B %-d')  # "May 7"
            except:
                formatted_date = due_date  # Fallback to original if parsing fails
            
            replacement = formatted_date
            text = text.replace(f'{{assignment:{assignment_name}}}', replacement)
    
    if close_conn:
        conn.close()
    
    return text

def substitute_slide_content(content_dict, conn):
    """Apply assignment variable substitution to all slide content fields"""
    if 'headline' in content_dict and content_dict['headline']:
        content_dict['headline'] = substitute_assignment_variables(content_dict['headline'], conn)
    if 'paragraph' in content_dict and content_dict['paragraph']:
        content_dict['paragraph'] = substitute_assignment_variables(content_dict['paragraph'], conn)
    if 'quote' in content_dict and content_dict['quote']:
        content_dict['quote'] = substitute_assignment_variables(content_dict['quote'], conn)
    if 'quote_citation' in content_dict and content_dict['quote_citation']:
        content_dict['quote_citation'] = substitute_assignment_variables(content_dict['quote_citation'], conn)
    if 'bullets' in content_dict and content_dict['bullets']:
        content_dict['bullets'] = [substitute_assignment_variables(b, conn) for b in content_dict['bullets']]
    return content_dict

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
                  notes TEXT,
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
                  hide_headline BOOLEAN DEFAULT 0,
                  larger_image BOOLEAN DEFAULT 0,
                  has_bullets BOOLEAN DEFAULT 1,
                  has_image BOOLEAN DEFAULT 0,
                  has_quote BOOLEAN DEFAULT 0,
                  is_gold BOOLEAN DEFAULT 0,
                  is_two_column BOOLEAN DEFAULT 0,
                  is_photo_centered BOOLEAN DEFAULT 0,
                  template_base TEXT,
                  module TEXT,
                  master_slide_id INTEGER,
                  fullscreen BOOLEAN DEFAULT 0,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  FOREIGN KEY (deck_id) REFERENCES decks(id) ON DELETE CASCADE)''')

    # Assignments table
    c.execute('''CREATE TABLE IF NOT EXISTS assignments
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  semester TEXT,
                  name TEXT NOT NULL,
                  due_date TEXT NOT NULL,
                  description TEXT,
                  points INTEGER DEFAULT 0,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    def ensure_column(table, column, definition):
        c.execute(f"PRAGMA table_info({table})")
        columns = [row[1] for row in c.fetchall()]
        if column not in columns:
            c.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

    # Backfill new columns for existing databases
    ensure_column('decks', 'notes', 'TEXT')
    ensure_column('slides', 'hide_headline', 'INTEGER DEFAULT 0')
    ensure_column('slides', 'larger_image', 'INTEGER DEFAULT 0')
    ensure_column('slides', 'has_bullets', 'INTEGER DEFAULT 1')
    ensure_column('slides', 'has_image', 'INTEGER DEFAULT 0')
    ensure_column('slides', 'has_quote', 'INTEGER DEFAULT 0')
    ensure_column('slides', 'is_gold', 'INTEGER DEFAULT 0')
    ensure_column('slides', 'is_two_column', 'INTEGER DEFAULT 0')
    ensure_column('slides', 'is_photo_centered', 'INTEGER DEFAULT 0')
    ensure_column('slides', 'template_base', 'TEXT')
    ensure_column('slides', 'module', 'TEXT')
    ensure_column('slides', 'master_slide_id', 'INTEGER')
    ensure_column('slides', 'fullscreen', 'INTEGER DEFAULT 0')
    
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
        c.execute('SELECT id, week, date, order_index, notes FROM decks WHERE presentation_id = ? ORDER BY order_index', 
                 (presentation_id,))
        decks = []
        for deck_row in c.fetchall():
            deck_id = deck_row[0]
            c.execute('''SELECT id, slide_class, headline, paragraph, bullets, quote, quote_citation, 
                        image_path, order_index, is_title, hide_headline, larger_image, has_bullets, has_image, 
                        has_quote, is_gold, is_two_column, is_photo_centered, template_base, module, master_slide_id, fullscreen, is_draft FROM slides 
                        WHERE deck_id = ? ORDER BY order_index''', (deck_id,))
            slides = [{'id': s[0], 'slideClass': s[1], 'headline': s[2], 'paragraph': s[3], 
                      'bullets': json.loads(s[4]) if s[4] else [], 'quote': s[5], 
                      'quoteCitation': s[6], 'imagePath': s[7], 'orderIndex': s[8], 'isTitle': bool(s[9]),
                      'hideHeadline': bool(s[10]), 'largerImage': bool(s[11]), 'hasBullets': bool(s[12]), 'hasImage': bool(s[13]),
                      'hasQuote': bool(s[14]), 'isGold': bool(s[15]), 'isTwoColumn': bool(s[16]),
                      'isPhotoCentered': bool(s[17]), 'templateBase': s[18], 'module': s[19], 'masterSlideId': s[20], 'fullscreen': bool(s[21]), 'isDraft': bool(s[22])}
                     for s in c.fetchall()]
            decks.append({'id': deck_id, 'week': deck_row[1], 'date': deck_row[2], 
                         'orderIndex': deck_row[3], 'notes': deck_row[4], 'slides': slides})
        
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
    
    # Automatically create slides for this deck:
    # 1. Title slide
    c.execute('''INSERT INTO slides (deck_id, slide_class, headline, paragraph, bullets, 
                quote, quote_citation, image_path, order_index, is_title, template_base) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
             (deck_id, 'title', '', '', '[]', '', '', '', 0, True, 'title'))
    
    # 2-5. Four headline/bullet slides
    for i in range(1, 5):
        c.execute('''INSERT INTO slides (deck_id, slide_class, headline, paragraph, bullets, 
                    quote, quote_citation, image_path, order_index, is_title, template_base, has_bullets) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                 (deck_id, 'template-bullets', '', '', '[""]', '', '', '', i, False, 'bullets', True))
    
    # 6. Closing slide with default content
    default_headline = """Damon Kiesow
Knight Chair in 
Journalism Innovation
216 Reynolds Journalism Institute
dkiesow@missouri.edu"""
    
    c.execute('''INSERT INTO slides (deck_id, slide_class, headline, paragraph, bullets, 
                quote, quote_citation, image_path, order_index, is_title, template_base) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
             (deck_id, 'closing', default_headline, 'Thank You', '[]', '', '', '', 5, False, 'closing'))
    
    conn.commit()
    conn.close()
    return jsonify({'id': deck_id})

@app.route('/api/decks/<int:deck_id>', methods=['PUT', 'DELETE'])
def deck(deck_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    if request.method == 'PUT':
        data = request.json
        c.execute('UPDATE decks SET week = ?, date = ?, notes = ? WHERE id = ?',
                 (data.get('week'), data.get('date'), data.get('notes'), deck_id))
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
    
    deck_id = data['deck_id']
    insert_after_slide_id = data.get('insert_after_slide_id')
    
    if insert_after_slide_id:
        # Get the order_index of the slide we're inserting after
        c.execute('SELECT order_index FROM slides WHERE id = ?', (insert_after_slide_id,))
        result = c.fetchone()
        if result:
            insert_after_order = result[0]
            # Shift all slides after this position down by 1
            c.execute('UPDATE slides SET order_index = order_index + 1 WHERE deck_id = ? AND order_index > ?',
                     (deck_id, insert_after_order))
            new_order_index = insert_after_order + 1
        else:
            # Fallback to end if slide not found
            c.execute('SELECT MAX(order_index) FROM slides WHERE deck_id = ?', (deck_id,))
            max_order = c.fetchone()[0] or -1
            new_order_index = max_order + 1
    else:
        # Insert at end
        c.execute('SELECT MAX(order_index) FROM slides WHERE deck_id = ?', (deck_id,))
        max_order = c.fetchone()[0] or -1
        new_order_index = max_order + 1
    
    c.execute('''INSERT INTO slides (deck_id, slide_class, headline, paragraph, bullets, 
                quote, quote_citation, image_path, order_index, is_title, is_draft) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
             (deck_id, data.get('class', ''), data.get('headline', ''),
              data.get('paragraph', ''), json.dumps(data.get('bullets', [])),
              data.get('quote', ''), data.get('quoteCitation', ''),
              data.get('imagePath', ''), new_order_index, data.get('isTitle', False), data.get('isDraft', False)))
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
        if 'hideHeadline' in data:
            update_fields.append('hide_headline = ?')
            update_values.append(1 if data.get('hideHeadline') else 0)
        if 'fullscreen' in data:
            update_fields.append('fullscreen = ?')
            update_values.append(1 if data.get('fullscreen') else 0)
        if 'hasBullets' in data:
            update_fields.append('has_bullets = ?')
            update_values.append(1 if data.get('hasBullets') else 0)
        if 'hasImage' in data:
            update_fields.append('has_image = ?')
            update_values.append(1 if data.get('hasImage') else 0)
        if 'hasQuote' in data:
            update_fields.append('has_quote = ?')
            update_values.append(1 if data.get('hasQuote') else 0)
        if 'isGold' in data:
            update_fields.append('is_gold = ?')
            update_values.append(1 if data.get('isGold') else 0)
        if 'isTwoColumn' in data:
            update_fields.append('is_two_column = ?')
            update_values.append(1 if data.get('isTwoColumn') else 0)
        if 'isPhotoCentered' in data:
            update_fields.append('is_photo_centered = ?')
            update_values.append(1 if data.get('isPhotoCentered') else 0)
        if 'templateBase' in data:
            update_fields.append('template_base = ?')
            update_values.append(data.get('templateBase'))
        if 'module' in data:
            update_fields.append('module = ?')
            update_values.append(data.get('module'))
        if 'deck_id' in data:
            update_fields.append('deck_id = ?')
            update_values.append(data.get('deck_id'))
        if 'isDraft' in data:
            update_fields.append('is_draft = ?')
            update_values.append(1 if data.get('isDraft') else 0)
        if 'largerImage' in data:
            update_fields.append('larger_image = ?')
            update_values.append(1 if data.get('largerImage') else 0)
        
        if update_fields:
            update_values.append(slide_id)
            c.execute(f"UPDATE slides SET {', '.join(update_fields)} WHERE id = ?", tuple(update_values))
            
            # Check if this is a master slide or instance and cascade changes
            c.execute('SELECT master_slide_id FROM slides WHERE id = ?', (slide_id,))
            result = c.fetchone()
            master_id = result[0] if result else None
            
            # Determine the master ID (either this slide or its master)
            cascade_id = master_id if master_id else slide_id
            
            # Build cascadable update (exclude deck_id and order_index)
            cascade_fields = []
            cascade_values = []
            non_cascade = ['deck_id']
            
            for i, field in enumerate(update_fields):
                field_name = field.split(' = ')[0]
                if field_name not in non_cascade:
                    cascade_fields.append(field)
                    cascade_values.append(update_values[i])
            
            if cascade_fields:
                # If editing an instance, update its master
                if master_id:
                    cascade_values_with_id = cascade_values + [cascade_id]
                    c.execute(f"UPDATE slides SET {', '.join(cascade_fields)} WHERE id = ?", 
                             tuple(cascade_values_with_id))
                
                # Update all instances of this master
                cascade_values_with_id = cascade_values + [cascade_id, slide_id]
                c.execute(f"UPDATE slides SET {', '.join(cascade_fields)} WHERE master_slide_id = ? AND id != ?", 
                         tuple(cascade_values_with_id))
            
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

@app.route('/api/modules', methods=['GET'])
def get_modules():
    """Get list of all distinct module names"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    c.execute('''SELECT DISTINCT module, COUNT(*) as slide_count 
                 FROM slides 
                 WHERE module IS NOT NULL AND module != '' 
                 GROUP BY module 
                 ORDER BY module''')
    
    modules = [{'name': row[0], 'slideCount': row[1]} for row in c.fetchall()]
    conn.close()
    return jsonify(modules)

@app.route('/api/decks/<int:deck_id>/insert-module', methods=['POST'])
def insert_module(deck_id):
    """Copy all slides from a module into a deck"""
    data = request.json
    module_name = data.get('moduleName')
    insert_after = data.get('insertAfter', None)  # Slide ID to insert after, or None for end
    
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Get all slides from the module (masters)
    c.execute('''SELECT id, slide_class, headline, paragraph, bullets, quote, quote_citation, 
                 image_path, hide_headline, larger_image, has_bullets, has_image, has_quote, is_gold, 
                 is_two_column, is_photo_centered, template_base, module, fullscreen 
                 FROM slides 
                 WHERE module = ? AND is_title = 0 AND master_slide_id IS NULL
                 ORDER BY order_index''', (module_name,))
    
    module_slides = c.fetchall()
    
    if not module_slides:
        conn.close()
        return jsonify({'error': 'Module not found'}), 404
    
    # Determine insertion point
    if insert_after:
        c.execute('SELECT order_index FROM slides WHERE id = ?', (insert_after,))
        result = c.fetchone()
        if result:
            start_order = result[0] + 1
        else:
            # If slide not found, append to end
            c.execute('SELECT MAX(order_index) FROM slides WHERE deck_id = ?', (deck_id,))
            max_order = c.fetchone()[0]
            start_order = (max_order or 0) + 1
    else:
        # Insert at end
        c.execute('SELECT MAX(order_index) FROM slides WHERE deck_id = ?', (deck_id,))
        max_order = c.fetchone()[0]
        start_order = (max_order or 0) + 1
    
    # Shift existing slides if needed
    if insert_after:
        c.execute('''UPDATE slides 
                     SET order_index = order_index + ? 
                     WHERE deck_id = ? AND order_index >= ?''',
                  (len(module_slides), deck_id, start_order))
    
    # Insert the module slides as linked instances
    inserted_ids = []
    for idx, slide_data in enumerate(module_slides):
        master_id = slide_data[0]
        c.execute('''INSERT INTO slides 
                     (deck_id, slide_class, headline, paragraph, bullets, quote, quote_citation, 
                      image_path, order_index, is_title, hide_headline, larger_image, has_bullets, has_image, 
                      has_quote, is_gold, is_two_column, is_photo_centered, template_base, module, master_slide_id, fullscreen)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (deck_id, slide_data[1], slide_data[2], slide_data[3], slide_data[4], 
                   slide_data[5], slide_data[6], slide_data[7], start_order + idx, 
                   slide_data[8], slide_data[9], slide_data[10], slide_data[11], slide_data[12],
                   slide_data[13], slide_data[14], slide_data[15], slide_data[16], slide_data[17], master_id, slide_data[18]))
        inserted_ids.append(c.lastrowid)
    
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'insertedCount': len(inserted_ids), 'insertedIds': inserted_ids})

@app.route('/api/assignments', methods=['GET'])
def get_assignments():
    """Get all assignments"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    c.execute('''SELECT id, name, due_date, description, points, created_at 
                 FROM assignments 
                 ORDER BY due_date''')
    
    assignments = [{
        'id': row[0],
        'name': row[1],
        'dueDate': row[2],
        'description': row[3],
        'points': row[4],
        'createdAt': row[5]
    } for row in c.fetchall()]
    
    conn.close()
    return jsonify(assignments)

@app.route('/api/assignments', methods=['POST'])
def create_assignment():
    """Create a new assignment"""
    data = request.json
    name = data.get('name')
    due_date = data.get('dueDate')
    description = data.get('description', '')
    points = data.get('points', 0)
    
    if not name or not due_date:
        return jsonify({'error': 'Name and due date are required'}), 400
    
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    c.execute('''INSERT INTO assignments (name, due_date, description, points)
                 VALUES (?, ?, ?, ?)''',
              (name, due_date, description, points))
    
    assignment_id = c.lastrowid
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'id': assignment_id})

@app.route('/api/assignments/<int:assignment_id>', methods=['PUT'])
def update_assignment(assignment_id):
    """Update an assignment"""
    data = request.json
    
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    update_fields = []
    update_values = []
    
    if 'name' in data:
        update_fields.append('name = ?')
        update_values.append(data['name'])
    if 'dueDate' in data:
        update_fields.append('due_date = ?')
        update_values.append(data['dueDate'])
    if 'description' in data:
        update_fields.append('description = ?')
        update_values.append(data['description'])
    if 'points' in data:
        update_fields.append('points = ?')
        update_values.append(data['points'])
    
    if update_fields:
        update_values.append(assignment_id)
        c.execute(f"UPDATE assignments SET {', '.join(update_fields)} WHERE id = ?",
                  tuple(update_values))
    
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/assignments/<int:assignment_id>', methods=['DELETE'])
def delete_assignment(assignment_id):
    """Delete an assignment"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    c.execute('DELETE FROM assignments WHERE id = ?', (assignment_id,))
    
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/assignments/import-csv', methods=['POST'])
def import_assignments_csv():
    """Import assignments from CSV (Semester, assignment name, due date)"""
    try:
        data = request.json
        csv_content = data.get('csvContent', '')
        
        if not csv_content:
            return jsonify({'error': 'No CSV content provided'}), 400
        
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        
        imported_count = 0
        errors = []
        
        # Parse CSV
        import csv
        from io import StringIO
        reader = csv.reader(StringIO(csv_content))
        
        # First pass: collect all semester codes and parse all rows
        rows_to_import = []
        semesters_in_csv = set()
        
        for row_num, row in enumerate(reader, 1):
            if len(row) < 3:
                errors.append(f"Row {row_num}: Not enough columns (need 3: semester, name, date)")
                continue
            
            semester = row[0].strip()
            name = row[1].strip()
            due_date_str = row[2].strip()
            
            semesters_in_csv.add(semester)
            
            # Parse date (format: M/D/YYYY)
            try:
                parts = due_date_str.split('/')
                if len(parts) == 3:
                    month, day, year = parts
                    # Convert to YYYY-MM-DD format
                    due_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                else:
                    errors.append(f"Row {row_num}: Invalid date format '{due_date_str}' (expected M/D/YYYY)")
                    continue
            except Exception as e:
                errors.append(f"Row {row_num}: Could not parse date '{due_date_str}': {str(e)}")
                continue
            
            # Store for later insertion
            rows_to_import.append((semester, name, due_date))
        
        # Delete all assignments for semesters present in the CSV
        if semesters_in_csv:
            placeholders = ','.join('?' * len(semesters_in_csv))
            c.execute(f'DELETE FROM assignments WHERE semester IN ({placeholders})', 
                     tuple(semesters_in_csv))
            deleted_count = c.rowcount
        else:
            deleted_count = 0
        
        # Insert all new assignments
        for semester, name, due_date in rows_to_import:
            try:
                c.execute('''INSERT INTO assignments (semester, name, due_date)
                             VALUES (?, ?, ?)''',
                          (semester, name, due_date))
                imported_count += 1
            except Exception as e:
                errors.append(f"Database error inserting '{name}': {str(e)}")
        
        conn.commit()
        conn.close()
        
        return jsonify({
            'success': True,
            'deleted': deleted_count,
            'imported': imported_count,
            'semesters': list(semesters_in_csv),
            'errors': errors
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/decks/<int:deck_id>/export', methods=['GET'])
def export_deck(deck_id):
    """Export a specific deck to PDF or PPTX format"""
    try:
        format_type = request.args.get('format', 'pdf').lower()
        
        if format_type not in ['pdf', 'pptx', 'odp']:
            return jsonify({'error': 'Invalid format. Use pdf, pptx, or odp'}), 400
        
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
        # Only add Marp frontmatter for PDF exports
        if format_type == 'pdf':
            content = '''---
marp: true
theme: classroom
paginate: true
COURSE_TITLE: "Journalism Innovation"
---

'''
        else:
            # For PPTX/ODP, start with empty content (no Marp directives)
            content = ''
        
        # Get all slides for this deck
        c.execute('''SELECT slide_class, headline, paragraph, bullets, quote, quote_citation, 
                            image_path, is_title, hide_headline, larger_image, fullscreen, template_base
                     FROM slides 
                     WHERE deck_id = ?
                     ORDER BY order_index''', (deck_id,))
        
        slides = c.fetchall()
        
        # Load PowerPoint layout mapping
        try:
            with open('pptx_layouts.json', 'r') as f:
                pptx_layouts = json.load(f)
        except:
            pptx_layouts = {}
        
        is_first_slide = True
        for slide in slides:
            slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, template_base = slide
            
            # Apply assignment variable substitution
            if headline:
                headline = substitute_assignment_variables(headline, conn)
            if paragraph:
                paragraph = substitute_assignment_variables(paragraph, conn)
            if bullets:
                try:
                    bullet_list = json.loads(bullets)
                    bullets = json.dumps([substitute_assignment_variables(b, conn) for b in bullet_list])
                except:
                    pass
            if quote:
                quote = substitute_assignment_variables(quote, conn)
            if quote_citation:
                quote_citation = substitute_assignment_variables(quote_citation, conn)
            
            # Add slide separator (skip for first slide if no frontmatter)
            if is_first_slide and format_type in ['pptx', 'odp']:
                # First slide in PPTX/ODP - no separator needed
                is_first_slide = False
            else:
                content += '\n---\n\n'
                is_first_slide = False
            
            # For PPTX/ODP export: add layout hint based on template_base
            pptx_layout = None
            if format_type in ['pptx', 'odp'] and template_base:
                pptx_layout = pptx_layouts.get(template_base)
            
            # Check if this is a title slide
            if is_title:
                # Title slide background + colorbar div (for PDF)
                if format_type == 'pdf':
                    content += f'<!-- _class: title -->\n'
                    content += f'![bg](../assets/title-background.jpg)\n'
                    content += f'<div class="colorbar"></div>\n'
                    content += f'# Journalism Innovation\n'
                    content += f'## {week}\n'
                    content += f'{date}\n'
                elif format_type in ['pptx', 'odp']:
                    # For PPTX, use Arches_Title layout
                    # Put everything in the div to ensure pandoc treats it as slide content
                    content += f'::: {{custom-style="Arches_Title"}}\n\n'
                    content += f'# Journalism Innovation\n\n'
                    content += f'## {week}\n\n'
                    content += f'{date}\n\n'
                    content += f':::\n'
                else:
                    # Fallback for other formats
                    content += f'# Journalism Innovation\n'
                    content += f'## {week}\n'
                    content += f'{date}\n'
                continue
            
            # Add class for PDF (Marp uses CSS classes)
            if slide_class and format_type == 'pdf':
                classes = [slide_class]
                if hide_headline:
                    classes.append('hide-headline')
                if fullscreen:
                    classes.append('fullscreen')
                content += f'<!-- _class: {" ".join(classes)} -->\n'
            
            # Determine template type FIRST before using these variables
            is_quote_template = slide_class and 'quote' in slide_class
            is_image_template = slide_class and ('image' in slide_class or slide_class == 'photo-centered')
            is_text_only = slide_class and 'lines' in slide_class
            is_photo_centered = slide_class == 'photo-centered'
            is_closing = slide_class == 'closing'
            
            # For PPTX/ODP with pandoc: headline and images go BEFORE the div, text content goes INSIDE
            if format_type in ['pptx', 'odp'] and pptx_layout:
                # Add headline at top level (becomes slide title) - but not for quotes or closing
                if headline and not is_closing and not is_quote_template:
                    content += f'# {headline}\n\n'
                
                # Add image at top level (before div) for photo-centered and image templates
                if image_path and (is_photo_centered or is_image_template):
                    # Convert image paths for PPTX/ODP
                    if image_path.startswith('/assets/'):
                        image_path = 'assets/' + image_path[8:]
                    elif image_path.startswith('assets/'):
                        pass  # Keep as-is
                    content += f'![Image]({image_path})\n\n'
                
                # Start custom-style div for text content only
                content += f'::: {{custom-style="{pptx_layout}"}}\n\n'
            
            if is_closing:
                # Closing slide: headline (contact info) and paragraph (thank you)
                if headline:
                    if format_type in ['pptx', 'odp']:
                        # For PPTX, put first line as title, rest in content div
                        lines = headline.split('\n')
                        if lines:
                            content += f'# {lines[0]}\n\n'
                            # Add custom-style div if pptx_layout exists
                            if pptx_layout and format_type in ['pptx', 'odp']:
                                content += f'::: {{custom-style="{pptx_layout}"}}\n\n'
                            # Add remaining headline lines
                            if len(lines) > 1:
                                content += '\n'.join(lines[1:]) + '\n\n'
                            # Add paragraph
                            if paragraph:
                                content += f'{paragraph}\n\n'
                            # Add logo
                            if format_type == 'pdf':
                                content += f'![width:400px](../assets/journalism_school_logo.png)\n'
                            # Close div for PPTX
                            if pptx_layout and format_type in ['pptx', 'odp']:
                                content += ':::\n'
                    else:
                        # For PDF, use span styling
                        lines = headline.split('\n')
                        if lines:
                            remaining_lines = '<br>'.join(lines[1:]) if len(lines) > 1 else ''
                            if remaining_lines:
                                content += f'# <span class="closing-name">{lines[0]}</span><br>{remaining_lines}\n'
                            else:
                                content += f'# <span class="closing-name">{lines[0]}</span>\n'
                        if paragraph:
                            content += f'\n{paragraph}\n'
                        # Add school logo at bottom
                        content += f'\n![width:400px](../assets/journalism_school_logo.png)\n'
            elif is_quote_template:
                # Quote templates: only export quote and citation
                if quote:
                    content += f'> {quote}\n'
                    if quote_citation:
                        content += f'>\n> {quote_citation}\n'
            elif is_photo_centered:
                # Photo centered: only headline and image (headline already added above for PPTX, image too)
                if headline and format_type == 'pdf':
                    content += f'# {headline}\n'
                
                # Add image if present (only for PDF, PPTX already added it above)
                if image_path and format_type == 'pdf':
                    # Convert absolute web paths to relative filesystem paths
                    if image_path.startswith('/assets/'):
                        image_path = '../assets/' + image_path[8:]
                    elif image_path.startswith('assets/'):
                        image_path = '../' + image_path
                    content += f'\n![Image]({image_path})\n'
            else:
                # Non-quote templates: export headline, paragraph, bullets
                # For PDF, add headline here; for PPTX it was already added above
                if headline and format_type == 'pdf':
                    content += f'# {headline}\n'
                
                if is_text_only:
                    # Text-only template: preserve line breaks as they are semantic
                    if paragraph:
                        content += f'\n{paragraph}\n'
                else:
                    # Bullet templates: show paragraph (if any) then bullets
                    if paragraph:
                        content += f'\n{paragraph}\n'
                    
                    if bullets:
                        try:
                            bullet_list = json.loads(bullets)
                            content += '\n'
                            for bullet in bullet_list:
                                if bullet.strip():
                                    content += f'- {bullet}\n'
                        except json.JSONDecodeError:
                            pass
                
                # Add image if present (only for PDF; PPTX already added it before the div)
                if image_path and is_image_template and format_type == 'pdf':
                    # Convert absolute web paths to relative filesystem paths
                    if image_path.startswith('/assets/'):
                        image_path = '../assets/' + image_path[8:]
                    elif image_path.startswith('assets/'):
                        image_path = '../' + image_path
                    content += f'\n![Image]({image_path})\n'
            
            # Close PPTX/ODP layout div
            if format_type in ['pptx', 'odp'] and pptx_layout:
                content += '\n:::\n'
        
        conn.close()
        
        # Write temporary markdown file
        # Sanitize filename - replace slashes and other problematic characters
        safe_week = str(week).replace('/', '-')
        safe_date = str(date).replace('/', '-').replace(' ', '_')
        
        temp_md = f'output/deck_{deck_id}_temp.md'
        output_file = f'output/Week_{safe_week}_{safe_date}.{format_type}'
        download_name = f'Week_{safe_week}_{safe_date}.{format_type}'
        
        with open(temp_md, 'w') as f:
            f.write(content)
        
        # Export logic by format
        if format_type == 'pdf':
            # Use Marp for PDF - renders with perfect styling
            cmd = f'marp "{temp_md}" -o "{output_file}" --allow-local-files --pdf --theme presentation-styles.css'
            print(f"Running command: {cmd}")
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            print(f"Return code: {result.returncode}")
            print(f"Stdout: {result.stdout}")
            print(f"Stderr: {result.stderr}")
        elif format_type == 'pptx':
            # Use python-pptx for direct PPTX generation with custom layouts
            print(f"Building PPTX with python-pptx using custom layouts")
            try:
                # Get slides data again for pptx_builder
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute('''SELECT slide_class, headline, paragraph, bullets, quote, quote_citation, 
                                    image_path, is_title, hide_headline, larger_image, fullscreen, template_base
                             FROM slides 
                             WHERE deck_id = ?
                             ORDER BY order_index''', (deck_id,))
                slides_data = c.fetchall()
                
                # Apply assignment variable substitution to slides_data
                processed_slides = []
                for slide in slides_data:
                    slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, template_base = slide
                    
                    # Substitute assignment variables
                    if headline:
                        headline = substitute_assignment_variables(headline, conn)
                    if paragraph:
                        paragraph = substitute_assignment_variables(paragraph, conn)
                    if bullets:
                        try:
                            bullet_list = json.loads(bullets)
                            bullets = json.dumps([substitute_assignment_variables(b, conn) for b in bullet_list])
                        except:
                            pass
                    if quote:
                        quote = substitute_assignment_variables(quote, conn)
                    if quote_citation:
                        quote_citation = substitute_assignment_variables(quote_citation, conn)
                    
                    processed_slides.append((slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, template_base))
                
                conn.close()
                
                # Build PPTX using custom layouts
                success = build_pptx_from_slides(
                    slides_data=processed_slides,
                    output_path=output_file,
                    template_path='4734_template.potx',
                    pptx_layouts_map=pptx_layouts,
                    deck_info={'week': week, 'date': date, 'course_title': 'Journalism Innovation'}
                )
                
                if success:
                    print(f"Successfully created PPTX: {output_file}")
                    result = type('obj', (object,), {'returncode': 0})()
                else:
                    print(f"Failed to create PPTX")
                    result = type('obj', (object,), {'returncode': 1})()
            except Exception as e:
                print(f"Error building PPTX: {e}")
                import traceback
                traceback.print_exc()
                result = type('obj', (object,), {'returncode': 1})()
        elif format_type == 'odp':
            # For ODP, first create PPTX with python-pptx then convert with LibreOffice
            temp_pptx = output_file.replace('.odp', '_temp.pptx')
            print(f"Building temporary PPTX for ODP conversion")
            try:
                # Get slides data for pptx_builder
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute('''SELECT slide_class, headline, paragraph, bullets, quote, quote_citation, 
                                    image_path, is_title, hide_headline, larger_image, fullscreen, template_base
                             FROM slides 
                             WHERE deck_id = ?
                             ORDER BY order_index''', (deck_id,))
                slides_data = c.fetchall()
                
                # Apply assignment variable substitution to slides_data
                processed_slides = []
                for slide in slides_data:
                    slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, template_base = slide
                    
                    # Substitute assignment variables
                    if headline:
                        headline = substitute_assignment_variables(headline, conn)
                    if paragraph:
                        paragraph = substitute_assignment_variables(paragraph, conn)
                    if bullets:
                        try:
                            bullet_list = json.loads(bullets)
                            bullets = json.dumps([substitute_assignment_variables(b, conn) for b in bullet_list])
                        except:
                            pass
                    if quote:
                        quote = substitute_assignment_variables(quote, conn)
                    if quote_citation:
                        quote_citation = substitute_assignment_variables(quote_citation, conn)
                    
                    processed_slides.append((slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, template_base))
                
                conn.close()
                
                # Build PPTX
                build_pptx_from_slides(
                    slides_data=processed_slides,
                    output_path=temp_pptx,
                    template_path='4734_template.potx',
                    pptx_layouts_map=pptx_layouts,
                    deck_info={'week': week, 'date': date, 'course_title': 'Journalism Innovation'}
                )
                result = type('obj', (object,), {'returncode': 0})()
            except Exception as e:
                print(f"Error building temp PPTX: {e}")
                result = type('obj', (object,), {'returncode': 1})()
        
        # For ODP, convert PPTX to ODP using LibreOffice  
        if format_type == 'odp' and result.returncode == 0:
            soffice_path = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
            odp_cmd = f'"{soffice_path}" --headless --convert-to odp --outdir output "{temp_pptx}"'
            print(f"Converting to ODP: {odp_cmd}")
            odp_result = subprocess.run(odp_cmd, shell=True, capture_output=True, text=True)
            print(f"ODP conversion return code: {odp_result.returncode}")
            
            # Clean up temp PPTX
            if os.path.exists(temp_pptx):
                os.remove(temp_pptx)
            
            if odp_result.returncode != 0:
                return jsonify({'error': f'ODP conversion failed: {odp_result.stderr}'}), 500
        
        # Keep temp file for debugging - do not delete
        # if os.path.exists(temp_md):
        #     os.remove(temp_md)
        
        if result.returncode != 0:
            print(f"Conversion error: {result.stderr}")
            return jsonify({'error': f'Conversion failed: {result.stderr}'}), 500
        
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
    """Preserve line breaks in paragraph text"""
    if not text:
        return text
    # Return text as-is to preserve original formatting
    return text


def generate_presentation_markdown(presentation_id, deck_id=None):
    """Generate markdown content for a presentation or single deck"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    
    # Get presentation
    c.execute('SELECT * FROM presentations WHERE id = ?', (presentation_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return None, 'Presentation not found'
    
    # Build proper front matter with Marp configuration
    content = '''---
marp: true
theme: classroom
paginate: true
COURSE_TITLE: "Journalism Innovation"
---

'''
    
    # Get all decks and slides (or just one deck if deck_id specified)
    if deck_id:
        c.execute('SELECT id, week, date FROM decks WHERE id = ?', (deck_id,))
    else:
        c.execute('SELECT id, week, date FROM decks WHERE presentation_id = ? ORDER BY order_index', 
                 (presentation_id,))
    decks = c.fetchall()
    
    slides_markdown = []
    
    for deck in decks:
        deck_id = deck[0]
        week = deck[1]
        date = deck[2]
        
        c.execute('''SELECT slide_class, headline, paragraph, bullets, quote, quote_citation, 
                    image_path, is_title, hide_headline, larger_image, fullscreen FROM slides WHERE deck_id = ? ORDER BY order_index''', 
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
            hide_headline = slide[8]
            larger_image = slide[9]
            fullscreen = slide[10]
            
            # Apply assignment variable substitution
            if headline:
                headline = substitute_assignment_variables(headline, conn)
            if paragraph:
                paragraph = substitute_assignment_variables(paragraph, conn)
            if bullets:
                bullets = [substitute_assignment_variables(b, conn) for b in bullets]
            if quote:
                quote = substitute_assignment_variables(quote, conn)
            if quote_citation:
                quote_citation = substitute_assignment_variables(quote_citation, conn)
            
            slide_md = ''
            
            if is_title:
                # Title slide with week/date variables
                slide_md = f'<!--\nWEEK: "{week}"\nDATE: "{date}"\n_class: title\n-->\n'
                slide_md += f'# {{{{COURSE_TITLE}}}}\n'
                slide_md += f'## {{{{WEEK}}}}\n'
                slide_md += f'{{{{DATE}}}}\n'
            else:
                # Add class (append hide-headline and/or fullscreen class if needed)
                classes = [slide_class] if slide_class else []
                if hide_headline:
                    classes.append('hide-headline')
                if fullscreen:
                    classes.append('fullscreen')
                slide_md = f'<!-- _class: {" ".join(classes)} -->\n' if classes else ''
                
                # Determine what fields to export based on template type
                is_quote_template = slide_class and 'quote' in slide_class
                is_image_template = slide_class and ('image' in slide_class or slide_class == 'photo-centered')
                is_text_only = slide_class and 'lines' in slide_class
                is_photo_centered = slide_class == 'photo-centered'
                is_closing = slide_class == 'closing'
                
                if is_closing:
                    # Closing slide: headline (contact info) and paragraph (thank you)
                    if headline:
                        # Split headline into lines and wrap first line in span, join rest with <br>
                        lines = headline.split('\n')
                        if lines:
                            remaining_lines = '<br>'.join(lines[1:]) if len(lines) > 1 else ''
                            if remaining_lines:
                                slide_md += f'# <span class="closing-name">{lines[0]}</span><br>{remaining_lines}\n'
                            else:
                                slide_md += f'# <span class="closing-name">{lines[0]}</span>\n'
                    
                    if paragraph:
                        slide_md += f'\n{paragraph}\n'
                elif is_quote_template:
                    # Quote templates: only export quote and citation
                    if quote:
                        slide_md += f'> {quote}\n'
                        if quote_citation:
                            slide_md += f'>\n> {quote_citation}\n'
                elif is_photo_centered:
                    # Photo centered: only headline and image
                    if headline:
                        slide_md += f'# {headline}\n'
                    
                    # Add image if present
                    if image_path:
                        slide_md += f'\n![Image]({image_path})\n'
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
                    
                    # Only export image if template supports images (but not photo-centered, handled above)
                    if is_image_template and not is_photo_centered and image_path:
                        slide_md += f'\n![Image]({image_path})\n'
            
            slides_markdown.append(slide_md)
    
    content += '\n\n---\n\n'.join(slides_markdown)
    
    filename = f"{row[1].replace(' ', '_')}.md"
    conn.close()
    return content, filename

@app.route('/api/presentations/<int:presentation_id>/auto-export', methods=['POST'])
def auto_export_presentation(presentation_id):
    content, filename = generate_presentation_markdown(presentation_id)
    if content is None:
        return jsonify({'error': filename}), 404
    
    # Save to file
    filepath = os.path.join(os.path.dirname(__file__), filename)
    with open(filepath, 'w') as f:
        f.write(content)
    
    return jsonify({'success': True, 'file': filename})

@app.route('/api/presentations/<int:presentation_id>/preview', methods=['POST'])
def preview_presentation(presentation_id):
    # Generate markdown for full presentation
    content, _ = generate_presentation_markdown(presentation_id)
    if content is None:
        return jsonify({'error': 'Failed to generate markdown'}), 500
    
    # Write to presentation.md
    presentation_path = os.path.join(os.path.dirname(__file__), 'presentation.md')
    with open(presentation_path, 'w') as f:
        f.write(content)
    
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

@app.route('/api/decks/<int:deck_id>/preview', methods=['POST'])
def preview_deck(deck_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Get deck's presentation_id
    c.execute('SELECT presentation_id FROM decks WHERE id = ?', (deck_id,))
    deck_row = c.fetchone()
    conn.close()
    
    if not deck_row:
        return jsonify({'error': 'Deck not found'}), 404
    
    presentation_id = deck_row[0]
    
    # Generate markdown for just this deck
    content, _ = generate_presentation_markdown(presentation_id, deck_id=deck_id)
    if content is None:
        return jsonify({'error': 'Failed to generate markdown'}), 500
    
    # Write to presentation.md
    presentation_path = os.path.join(os.path.dirname(__file__), 'presentation.md')
    with open(presentation_path, 'w') as f:
        f.write(content)
    
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
                    image_path, is_title, hide_headline, larger_image, fullscreen FROM slides WHERE deck_id = ? ORDER BY order_index''', 
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
            hide_headline = slide[8]
            larger_image = slide[9]
            fullscreen = slide[10]
            
            # Apply assignment variable substitution
            if headline:
                headline = substitute_assignment_variables(headline, conn)
            if paragraph:
                paragraph = substitute_assignment_variables(paragraph, conn)
            if bullets:
                bullets = [substitute_assignment_variables(b, conn) for b in bullets]
            if quote:
                quote = substitute_assignment_variables(quote, conn)
            if quote_citation:
                quote_citation = substitute_assignment_variables(quote_citation, conn)
            
            slide_md = ''
            
            if is_title:
                # Title slide with week/date variables
                slide_md = f'<!--\nWEEK: "{week}"\nDATE: "{date}"\n_class: title\n-->\n'
                slide_md += f'# {{{{COURSE_TITLE}}}}\n'
                slide_md += f'## {{{{WEEK}}}}\n'
                slide_md += f'{{{{DATE}}}}\n'
            else:
                # Add class (append hide-headline and/or fullscreen class if needed)
                classes = [slide_class] if slide_class else []
                if hide_headline:
                    classes.append('hide-headline')
                if fullscreen:
                    classes.append('fullscreen')
                slide_md = f'<!-- _class: {" ".join(classes)} -->\n' if classes else ''
                
                # Determine what fields to export based on template type
                is_quote_template = slide_class and 'quote' in slide_class
                is_image_template = slide_class and ('image' in slide_class or slide_class == 'photo-centered')
                is_text_only = slide_class and 'lines' in slide_class
                is_photo_centered = slide_class == 'photo-centered'
                is_closing = slide_class == 'closing'
                
                if is_closing:
                    # Closing slide: headline (contact info) and paragraph (thank you)
                    if headline:
                        # Split headline into lines and wrap first line in span, join rest with <br>
                        lines = headline.split('\n')
                        if lines:
                            remaining_lines = '<br>'.join(lines[1:]) if len(lines) > 1 else ''
                            if remaining_lines:
                                slide_md += f'# <span class="closing-name">{lines[0]}</span><br>{remaining_lines}\n'
                            else:
                                slide_md += f'# <span class="closing-name">{lines[0]}</span>\n'
                    
                    if paragraph:
                        slide_md += f'\n{paragraph}\n'
                elif is_quote_template:
                    # Quote templates: only export quote and citation
                    if quote:
                        slide_md += f'> {quote}\n'
                        if quote_citation:
                            slide_md += f'>\n> {quote_citation}\n'
                elif is_photo_centered:
                    # Photo centered: only headline and image
                    if headline:
                        slide_md += f'# {headline}\n'
                    
                    # Add image if present
                    if image_path:
                        slide_md += f'\n![Image]({image_path})\n'
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
                    
                    # Only export image if template supports images (but not photo-centered, handled above)
                    if is_image_template and not is_photo_centered and image_path:
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
