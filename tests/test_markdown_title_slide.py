import sqlite3
import tempfile
import os
import app


def test_title_slide_populates_variables(tmp_path):
    db_path = tmp_path / "test_presentations.db"
    conn = sqlite3.connect(str(db_path))
    c = conn.cursor()

    c.execute('''
        CREATE TABLE presentations (
            id INTEGER PRIMARY KEY,
            name TEXT
        )
    ''')

    c.execute('''
        CREATE TABLE decks (
            id INTEGER PRIMARY KEY,
            presentation_id INTEGER,
            week TEXT,
            date TEXT,
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
            order_index INTEGER
        )
    ''')

    conn.commit()

    c.execute('INSERT INTO presentations (id, name) VALUES (?, ?)', (1, 'Test Presentation'))
    c.execute('INSERT INTO decks (id, presentation_id, week, date, order_index) VALUES (?, ?, ?, ?, ?)', (1, 1, 'Week 1', '2026-01-13', 0))
    c.execute('''INSERT INTO slides (id, deck_id, slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, order_index)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
              (1, 1, 'title', '', '', '[]', '', '', '', 1, 0, 0, 0, 0))

    conn.commit()
    conn.close()

    app.DB_PATH = str(db_path)

    content, filename = app.generate_presentation_markdown(1, deck_id=1)
    assert content is not None

    assert '# Test Presentation' in content
    assert '## Week 1' in content
    assert '2026-01-13' in content
    # Ensure no unresolved placeholders remain
    assert '{{COURSE_TITLE}}' not in content
    assert '{{WEEK}}' not in content
    assert '{{DATE}}' not in content
