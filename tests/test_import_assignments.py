import sqlite3
import os
import tempfile
from scripts import import_assignments


def test_parse_date_handles_trailing_chars():
    assert import_assignments.parse_date('5/7/26%') == '2026-05-07'
    assert import_assignments.parse_date('02-13-2026') == '2026-02-13'


def test_import_insert_and_upsert(tmp_path):
    csv_path = tmp_path / 'assignments.csv'
    db_path = str(tmp_path / 'test_pres.db')

    csv_content = """
SP2026,Test Assignment,2/13/26,TEST
SP2026,Other Assignment,3/1/26,OTHER
""".strip()
    csv_path.write_text(csv_content)

    # Dry run first
    ins, upd, skip = import_assignments.import_assignments(str(csv_path), db_path=db_path, backup=False, dry_run=True)
    assert ins == 2 and upd == 0

    # Apply for real
    ins2, upd2, skip2 = import_assignments.import_assignments(str(csv_path), db_path=db_path, backup=False, dry_run=False)
    assert ins2 == 2

    # Now change a date and ensure upsert updates
    csv_content2 = """
SP2026,Test Assignment,3/10/26,TEST
""".strip()
    csv_path.write_text(csv_content2)

    ins3, upd3, skip3 = import_assignments.import_assignments(str(csv_path), db_path=db_path, backup=False, dry_run=False)
    assert upd3 == 1

    # Verify DB row updated and short field stored
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("SELECT due_date, short FROM assignments WHERE semester = ? AND name = ?", ('SP2026', 'Test Assignment'))
    r = c.fetchone()
    conn.close()
    assert r and r[0] == '2026-03-10' and r[1] == 'TEST'


def test_backup_creates_file(tmp_path):
    csv_path = tmp_path / 'assignments.csv'
    db_path = str(tmp_path / 'test_pres.db')
    csv_path.write_text('SP2026,A,2/13/26')

    # Create a dummy DB file
    open(db_path, 'wb').close()

    ins, upd, skip = import_assignments.import_assignments(str(csv_path), db_path=db_path, backup=True, dry_run=True)
    # A backup file should exist with .bak. in the name
    bak_files = [p for p in os.listdir(tmp_path) if p.startswith('test_pres.db.bak.')]
    assert bak_files
