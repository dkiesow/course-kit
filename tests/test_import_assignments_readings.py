import sqlite3

from scripts import import_assignments


def test_readings_rows_update_decks(tmp_path):
    csv_path = tmp_path / 'assignments.csv'
    db_path = str(tmp_path / 'test_pres.db')

    csv_content = '''
Semester,Week,Assignment,Due Date,XLS Short
SP2026,4,Chapter One - The News Product Manager,,
SP2026,1,Missouri School of Journalism students and alums leading new generation of news product thinkers,,
'''.strip()
    csv_path.write_text(csv_content)

    # Create DB and decks
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE decks (
            id INTEGER PRIMARY KEY,
            presentation_id INTEGER,
            week TEXT,
            date TEXT,
            order_index INTEGER,
            reading_list TEXT
        )
    ''')
    c.execute('INSERT INTO decks (id, week, date) VALUES (?, ?, ?)', (1, '4', '2026-02-13'))
    c.execute('INSERT INTO decks (id, week, date) VALUES (?, ?, ?)', (2, '1', '2026-01-11'))
    conn.commit()
    conn.close()

    # Dry run: should detect readings to add but not modify DB
    ins, upd, skip = import_assignments.import_assignments(str(csv_path), db_path=db_path, backup=False, dry_run=True)
    assert ins == 0 and upd == 0

    # Real run: apply readings
    ins2, upd2, skip2 = import_assignments.import_assignments(str(csv_path), db_path=db_path, backup=False, dry_run=False)
    assert ins2 == 0

    # Verify DB updates
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('SELECT reading_list FROM decks WHERE id = 1')
    r1 = c.fetchone()[0]
    assert 'Chapter One - The News Product Manager' in r1
    c.execute('SELECT reading_list FROM decks WHERE id = 2')
    r2 = c.fetchone()[0]
    assert 'Missouri School of Journalism' in r2
    conn.close()


def test_assignment_row_with_due_date_and_no_short_is_inserted(tmp_path):
    csv_path = tmp_path / 'assignments2.csv'
    db_path = str(tmp_path / 'test_pres2.db')

    csv_content = '''
Semester,Week,Assignment,Due Date,XLS Short
SP2026,7,Unit One Quiz,2/13/26,Quiz One
'''.strip()
    csv_path.write_text(csv_content)

    ins, upd, skip = import_assignments.import_assignments(str(csv_path), db_path=db_path, backup=False, dry_run=False)
    assert ins == 1

    # Verify DB row inserted
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("SELECT semester, name, due_date, short FROM assignments WHERE semester = ? AND name = ?", ('SP2026', 'Unit One Quiz'))
    r = c.fetchone()
    conn.close()
    assert r and r[0] == 'SP2026' and r[1] == 'Unit One Quiz' and r[2] == '2026-02-13' and r[3] == 'Quiz One'


def test_readings_are_idempotent(tmp_path):
    csv_path = tmp_path / 'assignments3.csv'
    db_path = str(tmp_path / 'test_pres3.db')

    csv_content = '''
Semester,Week,Assignment,Due Date,XLS Short
SP2026,4,Chapter One - The News Product Manager,,
'''.strip()
    csv_path.write_text(csv_content)

    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE decks (
            id INTEGER PRIMARY KEY,
            presentation_id INTEGER,
            week TEXT,
            date TEXT,
            order_index INTEGER,
            reading_list TEXT
        )
    ''')
    c.execute('INSERT INTO decks (id, week, date) VALUES (?, ?, ?)', (1, '4', '2026-02-13'))
    conn.commit()
    conn.close()

    # Apply twice
    import_assignments.import_assignments(str(csv_path), db_path=db_path, backup=False, dry_run=False)
    import_assignments.import_assignments(str(csv_path), db_path=db_path, backup=False, dry_run=False)

    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('SELECT reading_list FROM decks WHERE id = 1')
    r = c.fetchone()[0]
    conn.close()
    # Should only contain the reading once
    assert r.count('Chapter One - The News Product Manager') == 1
