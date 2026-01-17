import sqlite3
import tempfile
import os
import app


def test_front_matter_includes_presentation_and_assignments(tmp_path):
    # Create a temporary SQLite DB and initialize minimal schema
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

    c.execute('''
        CREATE TABLE assignments (
            id INTEGER PRIMARY KEY,
            semester TEXT,
            name TEXT NOT NULL,
            due_date TEXT NOT NULL,
            description TEXT,
            points INTEGER DEFAULT 0
        )
    ''')

    conn.commit()

    # Insert test data
    c.execute('INSERT INTO presentations (id, name) VALUES (?, ?)', (1, 'Test Presentation'))
    c.execute('INSERT INTO decks (id, presentation_id, week, date, order_index) VALUES (?, ?, ?, ?, ?)', (1, 1, 'Week 1', '01/01/2026', 0))
    c.execute('''INSERT INTO slides (id, deck_id, slide_class, headline, paragraph, bullets, quote, quote_citation, image_path, is_title, hide_headline, larger_image, fullscreen, order_index)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
              (1, 1, 'bullets', 'Headline', '', '[]', '', '', '', 0, 0, 0, 0, 0))

    # Insert an assignment
    c.execute('INSERT INTO assignments (id, semester, name, due_date) VALUES (?, ?, ?, ?)', (1, 'Spring', 'Test Assignment', '2026-03-01'))

    conn.commit()
    conn.close()

    # Point the app to the test DB
    app.DB_PATH = str(db_path)

    content, filename = app.generate_presentation_markdown(1)
    assert content is not None

    # Check front matter contains presentation title and assignment info
    assert 'PRESENTATION_TITLE: "Test Presentation"' in content
    assert 'EXPORT_DATE:' in content
    assert 'ASSIGNMENTS:' in content
    assert 'Test Assignment' in content
    assert '2026-03-01' in content
