#!/usr/bin/env python3
"""Remove duplicate assignments, keeping ones with Canvas linkage."""
import sqlite3

conn = sqlite3.connect('presentations.db')
c = conn.cursor()

# Find all assignments grouped by name
c.execute('''
SELECT name, GROUP_CONCAT(id || '|' || COALESCE(canvas_assignment_id, '') || '|' || due_date, ',') 
FROM assignments 
GROUP BY name 
HAVING COUNT(*) > 1
''')

duplicates = c.fetchall()
print(f'Found {len(duplicates)} duplicate assignment names:\n')

deleted_count = 0
for name, ids_info in duplicates:
    entries = []
    for entry in ids_info.split(','):
        parts = entry.split('|')
        aid = int(parts[0])
        canvas_id = parts[1] if len(parts) > 1 else ''
        due_date = parts[2] if len(parts) > 2 else ''
        entries.append((aid, canvas_id, due_date))
    
    # Keep the one with canvas_assignment_id, otherwise keep the first
    keep = None
    delete = []
    
    for aid, canvas_id, due_date in entries:
        if canvas_id:
            keep = (aid, canvas_id, due_date)
            break
    
    if not keep:
        keep = entries[0]
        delete = entries[1:]
    else:
        delete = [e for e in entries if e[0] != keep[0]]
    
    print(f'{name}:')
    print(f'  Keep: ID {keep[0]} (canvas: {keep[1] or "None"}, due: {keep[2]})')
    for aid, canvas_id, due_date in delete:
        print(f'  Delete: ID {aid} (canvas: {canvas_id or "None"}, due: {due_date})')
        c.execute('DELETE FROM assignments WHERE id = ?', (aid,))
        deleted_count += 1

conn.commit()
print(f'\nDeleted {deleted_count} duplicate assignments')
conn.close()
