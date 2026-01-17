#!/usr/bin/env python3
"""Sync readings and assignments from CSV to database (without touching deck/slide data)"""

import sqlite3
import csv
import sys

def sync_csv_data(csv_path, db_path='presentations.db'):
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    
    # Read CSV
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        next(reader)  # skip header
        
        csv_assignments = []
        for row in reader:
            if len(row) < 5:
                continue
            semester, week, assignment, due_date, xls_short = row[0], row[1], row[2], row[3], row[4]
            
            if assignment and due_date and assignment not in ['Holiday', 'No Class']:
                csv_assignments.append((assignment, due_date, xls_short or assignment, semester))
    
    # Update or create assignments
    updated = 0
    created = 0
    
    for name, due_date, short, semester in csv_assignments:
        # Check if assignment exists
        c.execute('SELECT id, due_date, short FROM assignments WHERE name = ?', (name,))
        row = c.fetchone()
        
        if row:
            aid, old_due, old_short = row
            # Update if changed
            if old_due != due_date or old_short != short:
                c.execute('UPDATE assignments SET due_date = ?, short = ? WHERE id = ?',
                         (due_date, short, aid))
                print(f'Updated: {name} - due:{due_date}, short:{short}')
                updated += 1
        else:
            # Create new assignment
            c.execute('INSERT INTO assignments (name, due_date, short, semester) VALUES (?, ?, ?, ?)',
                     (name, due_date, short, semester))
            print(f'Created: {name} - due:{due_date}, short:{short}')
            created += 1
    
    conn.commit()
    conn.close()
    
    print(f'\nSummary: {created} created, {updated} updated')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python sync_csv_data.py <csv_path>')
        sys.exit(1)
    
    sync_csv_data(sys.argv[1])
