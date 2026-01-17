import sqlite3
import app
import os


def test_put_updates_topics(tmp_path):
    db_path = tmp_path / "test_decks_topics.db"
    conn = sqlite3.connect(str(db_path))
    c = conn.cursor()

    c.execute('''
        CREATE TABLE decks (
            id INTEGER PRIMARY KEY,
            presentation_id INTEGER,
            week TEXT,
            date TEXT,
            topic1 TEXT,
            topic2 TEXT,
            notes TEXT,
            order_index INTEGER
        )
    ''')

    conn.commit()

    c.execute('INSERT INTO decks (id, presentation_id, week, date, topic1, topic2, notes, order_index) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
              (1, 1, 'Week 1', '2026-01-13', 'Old A', 'Old B', '', 0))
    conn.commit()
    conn.close()

    app.DB_PATH = str(db_path)

    client = app.app.test_client()
    resp = client.put('/api/decks/1', json={'week': 'Week 1', 'date': '2026-01-13', 'notes': '', 'topic1': 'New A', 'topic2': 'New B'})

    assert resp.status_code == 200

    conn = sqlite3.connect(str(db_path))
    c = conn.cursor()
    c.execute('SELECT topic1, topic2 FROM decks WHERE id = ?', (1,))
    row = c.fetchone()
    conn.close()

    assert row[0] == 'New A'
    assert row[1] == 'New B'


def test_roundtrip_topic_update_and_export(tmp_path):
    db_path = tmp_path / "test_roundtrip_topics.db"
    conn = sqlite3.connect(str(db_path))
    c = conn.cursor()

    # Create tables
    c.execute('''
        CREATE TABLE presentations (
            id INTEGER PRIMARY KEY,
            name TEXT,
            front_matter TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE decks (
            id INTEGER PRIMARY KEY,
            presentation_id INTEGER,
            week TEXT,
            date TEXT,
            topic1 TEXT,
            topic2 TEXT,
            notes TEXT,
            order_index INTEGER
        )
    ''')
    c.execute('''
        CREATE TABLE slides (
            id INTEGER PRIMARY KEY,
            deck_id INTEGER,
            slide_class TEXT,
            headline TEXT,
            paragraph TEXT,
            bullets TEXT,
            quote TEXT,
            quote_citation TEXT,
            image_path TEXT,
            is_title INTEGER,
            hide_headline INTEGER,
            larger_image INTEGER,
            fullscreen INTEGER,
            template_base TEXT,
            order_index INTEGER
        )
    ''')
    conn.commit()

    c.execute('INSERT INTO presentations (id, name) VALUES (?, ?)', (1, 'Test Presentation'))
    c.execute('INSERT INTO decks (id, presentation_id, week, date, topic1, topic2, notes, order_index) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
              (1, 1, 'Week 1', '2026-01-13', '', '', '', 0))
    c.execute('''INSERT INTO slides (id, deck_id, slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, template_base, order_index)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
              (1, 1, 'title', '', '', '[]', '', '', '', 1, 0, 0, 0, 'title', 0))
    conn.commit()
    conn.close()

    app.DB_PATH = str(db_path)
    client = app.app.test_client()

    # Update topics via API
    resp = client.put('/api/decks/1', json={'week': 'Week 1', 'date': '2026-01-13', 'notes': '', 'topic1': 'Round A', 'topic2': 'Round B'})
    assert resp.status_code == 200

    # Generate markdown and assert topics are NOT included on slides
    content, filename = app.generate_presentation_markdown(1, deck_id=1)
    assert 'Round A' not in content
    assert 'Round B' not in content

    # Build PPTX to verify topics appear there too
    conn = sqlite3.connect(str(db_path))
    c = conn.cursor()
    c.execute('SELECT slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, template_base FROM slides WHERE deck_id = ? ORDER BY order_index', (1,))
    slides_data = c.fetchall()
    c.execute('SELECT topic1, topic2 FROM decks WHERE id = ?', (1,))
    trow = c.fetchone()
    conn.close()

    # Process slides similar to export flow
    processed_slides = []
    for slide in slides_data:
        slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, template_base = slide
        if bullets:
            try:
                bullet_list = json.loads(bullets)
                bullets = json.dumps(bullet_list)
            except:
                pass
        processed_slides.append((slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, template_base))

    out = str(tmp_path / 'roundtrip.pptx')
    success = app.build_pptx_from_slides(
        slides_data=processed_slides,
        output_path=out,
        template_path='templates/4734_template.potx',
        pptx_layouts_map=__import__('json').load(open('pptx_layouts.json')),
        deck_info={'week': 'Week 1', 'date': '2026-01-13', 'course_title': 'Test Presentation'}
    )

    assert success
    assert os.path.exists(out)

    # Verify topics are NOT present in PPTX text
    prs = __import__('pptx').Presentation(out)
    texts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, 'has_text_frame', False):
                try:
                    if shape.text and shape.text.strip():
                        texts.append(shape.text.strip())
                except Exception:
                    pass
    combined = '\n'.join(texts)
    assert 'Round A' not in combined
    assert 'Round B' not in combined
