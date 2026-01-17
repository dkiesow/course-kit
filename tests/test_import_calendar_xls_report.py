import sqlite3
import csv
from openpyxl import Workbook
from scripts.import_calendar_xls import import_calendar_xls


def test_report_csv_contains_predicted_links(tmp_path):
    db_path = tmp_path / 'test_report.db'
    conn = sqlite3.connect(str(db_path))
    c = conn.cursor()
    c.execute('''
        CREATE TABLE decks (
            id INTEGER PRIMARY KEY,
            date TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE assignments (
            id INTEGER PRIMARY KEY,
            semester TEXT,
            name TEXT NOT NULL,
            due_date TEXT NOT NULL,
            short TEXT
        )
    ''')
    c.execute('INSERT INTO decks (id, date) VALUES (?, ?)', (1, '2026-01-11'))
    # Insert assignment with short 'Report Proposal'
    c.execute('INSERT INTO assignments (id, semester, name, due_date, short) VALUES (?, ?, ?, ?, ?)', (1, 'SP2026', 'Report Proposal', '2026-02-13', 'Report Proposal'))
    conn.commit()
    conn.close()

    # Create an XLSX with an uppercase short code
    wb = Workbook()
    ws = wb.active
    ws.append(['Unit', 'Reading', 'MON', 'WED', 'Assignments'])
    ws.append(['Unit 1 - 1/11/26', '', '', '', 'REPORT PROPOSAL'])
    xls_path = tmp_path / 'cal_report.xlsx'
    wb.save(str(xls_path))

    report_path = tmp_path / 'report.csv'

    updated, linked, skipped = import_calendar_xls(str(xls_path), db_path=str(db_path), semester='SP2026', dry_run=True, link_assignments=True, report_csv=str(report_path))

    # Report should exist
    assert report_path.exists()
    # Read CSV and find the row for deck 1
    with open(str(report_path), newline='') as rf:
        rdr = csv.DictReader(rf)
        rows = list(rdr)
    assert len(rows) == 1
    r = rows[0]
    # predicted_assignment_ids should include '1'
    assert '1' in r['predicted_assignment_ids']
    assert r['deck_id'] == '1'
    assert r['date_iso'] == '2026-01-11'


def test_per_week_readings_aggregated(tmp_path):
    db_path = tmp_path / 'test_report2.db'
    conn = sqlite3.connect(str(db_path))
    c = conn.cursor()
    c.execute('''
        CREATE TABLE decks (
            id INTEGER PRIMARY KEY,
            date TEXT,
            reading_list TEXT
        )
    ''')
    # Two decks in same week
    c.execute('INSERT INTO decks (id, date) VALUES (?, ?)', (8, '2026-01-26'))
    c.execute('INSERT INTO decks (id, date) VALUES (?, ?)', (9, '2026-01-28'))
    conn.commit()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.append(['Unit', 'Reading', 'MON', 'WED', 'Assignments'])
    ws.append(['Unit 1 - 1/26 - 1/28', 'Chapter One - The News Product Manager', '', '', ''])
    xls_path = tmp_path / 'cal_report2.xlsx'
    wb.save(str(xls_path))

    report_path = tmp_path / 'report2.csv'
    updated, linked, skipped = import_calendar_xls(str(xls_path), db_path=str(db_path), semester='SP2026', dry_run=True, link_assignments=False, report_csv=str(report_path), apply_readings='per_week')

    assert report_path.exists()
    with open(str(report_path), newline='') as rf:
        rdr = csv.DictReader(rf)
        rows = list(rdr)
    # Find rows that include reading_list in updates
    reading_rows = [r for r in rows if '"reading_list"' in r['updates_json'] or 'reading_list' in r['updates_json']]
    # There should be exactly one aggregated reading update for the week
    assert len(reading_rows) == 1
    r = reading_rows[0]
    assert r['deck_id'] in ('8', '9')
    assert 'Chapter One - The News Product Manager' in r['updates_json']