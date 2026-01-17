#!/usr/bin/env python3
"""Import assignments CSV into presentations.db assignments table.

CSV format (no header expected): semester, name, due_date
Example: SP2026,Final Report Proposal,2/13/26

Features:
- date normalization (M/D/YY or M/D/YYYY)
- --backup to create a timestamped copy of the DB before modifying it
- --dry-run to preview changes
- upsert behavior by (semester, name) by default
"""

from __future__ import annotations

import argparse
import csv
import datetime
import os
import re
import shutil
import sqlite3
import sys
import uuid
from typing import List, Tuple, Optional


def parse_date(s: str) -> Optional[str]:
    """Parse a date string like '2/13/26' or '02-13-2026' and return 'YYYY-MM-DD'."""
    if not s:
        return None
    s = s.strip()
    # Extract the first date-like token
    m = re.search(r"(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})", s)
    if not m:
        raise ValueError(f"No date found in {s!r}")
    tok = m.group(1)
    tok = tok.replace('.', '/').replace('-', '/')
    for fmt in ("%m/%d/%y", "%m/%d/%Y"):
        try:
            dt = datetime.datetime.strptime(tok, fmt).date()
            return dt.isoformat()
        except ValueError:
            continue
    raise ValueError(f"Unknown date format: {s!r}")

def generate_assignment_uuid(semester: str, short: str) -> str:
    """Generate a stable UUID for an assignment based on semester and short code.
    
    This ensures that re-importing the same assignment from CSV won't create duplicates.
    Uses UUID5 with a namespace for deterministic generation.
    
    Args:
        semester: Semester code (e.g., 'SP2026')
        short: Short code for assignment (e.g., 'PROJ1')
    
    Returns:
        UUID string
    """
    # Use a custom namespace UUID for our assignments
    namespace = uuid.UUID('a1b2c3d4-e5f6-7890-abcd-ef1234567890')
    # Normalize inputs to ensure consistency
    key = f"{semester.upper().strip()}:{short.upper().strip()}"
    return str(uuid.uuid5(namespace, key))

def read_csv_rows(path: str) -> List[dict]:
    """
    Read CSV rows supporting two formats:
      - Legacy: semester, name, due_date[, short]
      - New header format: Semester,Week,Assignment,Due Date,XLS Short
    Returns a list of dicts with keys:
      - kind: 'assignment' or 'reading'
      - semester, week (may be ''), name
      - due_raw (if assignment), short (if assignment)
    """
    rows = []
    with open(path, newline='') as f:
        rdr = csv.reader(f)
        first = next(rdr, None)
        def _norm_cell(cell):
            if cell is None:
                return ''
            s = str(cell).strip()
            # remove BOM if present
            if s and s[0] == '\ufeff':
                s = s.lstrip('\ufeff').strip()
            return s.lower()
        if first and any((c is not None and str(c).strip() != '') for c in first):
            # Heuristic: detect header-style CSV if at least two known header tokens are present in the first row
            header_tokens = {'semester', 'week', 'assignment', 'due date', 'due_date', 'due', 'xls short', 'xls_short', 'short', 'uuid'}
            header_norms = [_norm_cell(h) for h in first]
            matches = 0
            for h in header_norms:
                for token in header_tokens:
                    # exact or underscore/space-normalized match
                    if h == token or h.replace('_', ' ') == token:
                        matches += 1
                        break
            if matches >= 2:
                header = header_norms
                idx_sem = header.index('semester') if 'semester' in header else None
                idx_week = header.index('week') if 'week' in header else None
                idx_assignment = header.index('assignment') if 'assignment' in header else None
                # 'due date' might appear as 'due date' or 'due_date'
                idx_due = None
                for alt in ('due date', 'due_date', 'due'):
                    if alt in header:
                        idx_due = header.index(alt)
                        break
                idx_short = None
                for name in ('xls short', 'xls_short', 'short'):
                    if name in header:
                        idx_short = header.index(name)
                        break
                idx_uuid = header.index('uuid') if 'uuid' in header else None
                for i, r in enumerate(rdr, start=2):
                    if not r or all((c is None or str(c).strip() == '') for c in r[:5]):
                        continue
                    def get(idx):
                        return str(r[idx]).strip() if idx is not None and idx < len(r) and r[idx] is not None else ''
                    semester = get(idx_sem) if idx_sem is not None else ''
                    week = get(idx_week) if idx_week is not None else ''
                    name = get(idx_assignment) if idx_assignment is not None else ''
                    due_raw = get(idx_due) if idx_due is not None else ''
                    short = get(idx_short) if idx_short is not None else ''
                    uuid_val = get(idx_uuid) if idx_uuid is not None else ''
                    if short == '':
                        short = None
                    if uuid_val == '':
                        uuid_val = None
                    if short is None and (due_raw == '' or due_raw is None):
                        rows.append({'kind': 'reading', 'semester': semester, 'week': week, 'name': name})
                    else:
                        rows.append({'kind': 'assignment', 'semester': semester, 'week': week, 'name': name, 'due_raw': due_raw, 'short': short, 'uuid': uuid_val})
            else:
                # Treat as legacy (no header)
                if first:
                    r0 = first
                    rdr = [r0] + list(rdr)
                for i, r in enumerate(rdr, start=1):
                    if not r:
                        continue
                    if len(r) < 3:
                        print(f"Skipping malformed row {i}: {r}")
                        continue
                    semester = str(r[0]).strip()
                    name = str(r[1]).strip()
                    due_raw = str(r[2]).strip()
                    short = str(r[3]).strip() if len(r) >= 4 and str(r[3]).strip() != '' else None
                    if short is None and due_raw == '':
                        rows.append({'kind': 'reading','semester':semester,'week':'','name':name})
                    else:
                        rows.append({'kind':'assignment','semester':semester,'week':'','name':name,'due_raw':due_raw,'short':short})
        else:
            # Legacy, no header: semester, name, due_date[, short]
            # Rewind reader to include first row as data
            if first:
                r0 = first
                rdr = [r0] + list(rdr)
            for i, r in enumerate(rdr, start=1):
                if not r:
                    continue
                if len(r) < 3:
                    print(f"Skipping malformed row {i}: {r}")
                    continue
                semester = str(r[0]).strip()
                name = str(r[1]).strip()
                due_raw = str(r[2]).strip()
                short = str(r[3]).strip() if len(r) >= 4 and str(r[3]).strip() != '' else None
                if short is None and due_raw == '':
                    rows.append({'kind':'reading','semester':semester,'week':'','name':name})
                else:
                    rows.append({'kind':'assignment','semester':semester,'week':'','name':name,'due_raw':due_raw,'short':short})
    return rows


def ensure_assignments_table(conn: sqlite3.Connection) -> None:
    c = conn.cursor()
    # Create table if missing, include `short` and `uuid` columns
    c.execute('''CREATE TABLE IF NOT EXISTS assignments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    semester TEXT,
                    name TEXT NOT NULL,
                    due_date TEXT NOT NULL,
                    short TEXT,
                    uuid TEXT UNIQUE,
                    description TEXT,
                    points INTEGER DEFAULT 0
                )''')
    # Add columns if older schema lacks them
    c.execute("PRAGMA table_info(assignments)")
    cols = [r[1] for r in c.fetchall()]
    if 'short' not in cols:
        try:
            c.execute('ALTER TABLE assignments ADD COLUMN short TEXT')
        except Exception:
            pass
    if 'uuid' not in cols:
        try:
            c.execute('ALTER TABLE assignments ADD COLUMN uuid TEXT UNIQUE')
        except Exception:
            pass
    conn.commit()


def find_deck_ids_by_week(conn: sqlite3.Connection, week_value: str) -> List[int]:
    c = conn.cursor()
    if not week_value or str(week_value).strip() == '':
        return []
    w = str(week_value).strip()
    candidates: List[int] = []
    # Exact match
    c.execute('SELECT id FROM decks WHERE week = ?', (w,))
    candidates += [r[0] for r in c.fetchall()]
    # 'Week {w}' match
    c.execute('SELECT id FROM decks WHERE week = ?', (f'Week {w}',))
    candidates += [r[0] for r in c.fetchall()]
    # If there's a numeric token, try loose LIKE match as a fallback
    m = re.search(r"(\d+)", w)
    if m:
        num = m.group(1)
        c.execute('SELECT id FROM decks WHERE week LIKE ?', (f"%{num}%",))
        candidates += [r[0] for r in c.fetchall()]
    # Deduplicate while preserving order
    seen = set()
    uniq = []
    for d in candidates:
        if d not in seen:
            seen.add(d)
            uniq.append(d)
    return uniq


def import_assignments(csv_path: str, db_path: str = 'presentations.db', backup: bool = False, dry_run: bool = True, upsert: bool = True) -> Tuple[int, int, int]:
    rows = read_csv_rows(csv_path)
    planned_assignments: List[Tuple[str, str, str, Optional[str], Optional[str]]] = []  # (sem, name, due, short, uuid)
    planned_readings: List[Tuple[str, str, str]] = []
    errors = []

    for r in rows:
        kind = r.get('kind')
        if kind == 'assignment':
            sem = r.get('semester', '')
            name = r.get('name', '')
            due_raw = r.get('due_raw', '') or ''
            short = r.get('short')
            csv_uuid = r.get('uuid')
            try:
                due = parse_date(due_raw) if due_raw else None
            except Exception as e:
                errors.append((sem, name, due_raw, str(e)))
                continue
            if not due:
                # Missing due date for assignment is not acceptable; skip and report
                print(f"Skipping assignment {sem} / {name}: missing or unparseable due date")
                errors.append((sem, name, due_raw, 'missing due date'))
                continue
            planned_assignments.append((sem, name, due, short, csv_uuid))
        elif kind == 'reading':
            planned_readings.append((r.get('semester', ''), r.get('week', ''), r.get('name', '')))
        else:
            errors.append(('unknown', r))

    if not planned_assignments and not planned_readings:
        print("No valid rows to import.")
        return 0, 0, len(errors)

    if backup and os.path.exists(db_path):
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        bak_path = f"{db_path}.bak.{ts}"
        shutil.copy2(db_path, bak_path)
        print(f"Backup created: {bak_path}")

    conn = sqlite3.connect(db_path)
    ensure_assignments_table(conn)
    c = conn.cursor()

    inserted = 0
    updated = 0
    skipped = 0

    # Process assignments
    for sem, name, due, short, csv_uuid in planned_assignments:
        existing = None
        # Prefer matching by UUID if provided in CSV
        if csv_uuid:
            c.execute('SELECT id, due_date, short, name, uuid FROM assignments WHERE uuid = ?', (csv_uuid,))
            existing = c.fetchone()
        # Then try short code
        if not existing and short:
            c.execute('SELECT id, due_date, short, name, uuid FROM assignments WHERE semester = ? AND short = ?', (sem, short))
            existing = c.fetchone()
        # Finally fallback to name
        if not existing:
            c.execute('SELECT id, due_date, short, name, uuid FROM assignments WHERE semester = ? AND name = ?', (sem, name))
            existing = c.fetchone()

        if existing:
            eid, cur_due, cur_short, cur_name, cur_uuid = existing
            changed = False
            if upsert:
                if cur_due != due:
                    print(f"Update: {sem} / {name}: {cur_due} -> {due}")
                    if not dry_run:
                        c.execute('UPDATE assignments SET due_date = ? WHERE id = ?', (due, eid))
                        conn.commit()
                    changed = True
                # Update short if provided and different
                if short and cur_short != short:
                    print(f"Update short: {sem} / {name}: {cur_short} -> {short}")
                    if not dry_run:
                        c.execute('UPDATE assignments SET short = ? WHERE id = ?', (short, eid))
                        conn.commit()
                    changed = True
                # Backfill UUID if provided in CSV but missing in DB
                if csv_uuid and not cur_uuid:
                    print(f"Backfill UUID: {sem} / {name}: {csv_uuid}")
                    if not dry_run:
                        c.execute('UPDATE assignments SET uuid = ? WHERE id = ?', (csv_uuid, eid))
                        conn.commit()
                    changed = True
                # Optionally update name if name differs
                if name and cur_name != name:
                    if not dry_run:
                        c.execute('UPDATE assignments SET name = ? WHERE id = ?', (name, eid))
                        conn.commit()
                    changed = True
            if changed:
                updated += 1
            else:
                skipped += 1
        else:
            # Generate UUID if not provided in CSV
            if not csv_uuid and short:
                csv_uuid = generate_assignment_uuid(sem, short)
            print(f"Insert: {sem} / {name} -> {due} (short={short}, uuid={csv_uuid})")
            if not dry_run:
                c.execute('INSERT INTO assignments (semester, name, due_date, short, uuid) VALUES (?, ?, ?, ?, ?)', (sem, name, due, short, csv_uuid))
                conn.commit()
            inserted += 1

    # Process readings (update decks.reading_list)
    reading_updates = 0
    reading_skipped = 0
    for sem, week, name in planned_readings:
        deck_ids = find_deck_ids_by_week(conn, week)
        if not deck_ids:
            print(f"No deck found for week {week}; skipping reading '{name}'")
            reading_skipped += 1
            continue
        for deck_id in deck_ids:
            c.execute('SELECT reading_list FROM decks WHERE id = ?', (deck_id,))
            cur = c.fetchone()
            cur_text = cur[0] if cur and cur[0] else ''
            if name.strip() in cur_text.splitlines():
                print(f"Reading '{name.strip()}' already present in deck {deck_id}; skipping")
                continue
            new_text = name.strip() if not cur_text else cur_text + '\n' + name.strip()
            print(f"Add reading to deck {deck_id}: '{name.strip()}'")
            if not dry_run:
                c.execute('UPDATE decks SET reading_list = ? WHERE id = ?', (new_text, deck_id))
                conn.commit()
            reading_updates += 1

    if dry_run:
        print(f"Dry run: {inserted} inserts, {updated} updates, {skipped} skipped, {len(errors)} errors; {reading_updates} reading(s) to add, {reading_skipped} reading(s) skipped")
    else:
        print(f"Applied: {inserted} inserted, {updated} updated, {skipped} skipped, {len(errors)} errors; {reading_updates} reading(s) added, {reading_skipped} reading(s) skipped")

    conn.close()
    return inserted, updated, skipped


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description='Import assignments CSV into presentations.db')
    parser.add_argument('--csv', '-c', required=True, help='Path to CSV file')
    parser.add_argument('--db', default='presentations.db', help='Path to SQLite DB')
    parser.add_argument('--backup', action='store_true', help='Create a timestamped backup of the DB before modifying')
    parser.add_argument('--dry-run', action='store_true', help='Do not modify DB, just show actions')
    parser.add_argument('--no-upsert', action='store_true', help="Do not update existing rows (insert-only)")
    args = parser.parse_args(argv)

    try:
        inserted, updated, skipped = import_assignments(args.csv, db_path=args.db, backup=args.backup, dry_run=args.dry_run, upsert=not args.no_upsert)
        return 0
    except Exception as e:
        print(f"Error: {e}")
        return 2


if __name__ == '__main__':
    sys.exit(main())
