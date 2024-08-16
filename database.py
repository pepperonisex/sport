import sqlite3
import os
from datetime import datetime

def setup_database(db_name="database.db", db_dir="db"):
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    full_db_dir = os.path.join(script_dir, db_dir)
    os.makedirs(full_db_dir, exist_ok=True)
    
    full_path = os.path.join(full_db_dir, db_name)
    
    conn = sqlite3.connect(full_path)
    
    with conn:
        conn.execute('''
        CREATE TABLE IF NOT EXISTS worksheet_info (
            id INTEGER PRIMARY KEY,
            worksheet_id TEXT NOT NULL,
            next_col TEXT NOT NULL,
            next_row INTEGER NOT NULL,
            last_updated TEXT DEFAULT (strftime('%Y-%m-%d', 'now')) 
        )
        ''')
    
    return conn


def add_worksheet_info(conn, worksheet_id, next_col, next_row):
    with conn:
        conn.execute('''
        INSERT INTO worksheet_info (worksheet_id, next_col, next_row, last_updated)
        VALUES (?, ?, ?, ?)
        ''', (worksheet_id, next_col, next_row, datetime.now().strftime('%Y-%m-%d')))

def update_worksheet_info(conn, worksheet_id, next_col, next_row):
    with conn:
        conn.execute('''
            UPDATE worksheet_info
            SET next_col = ?, next_row = ?, last_updated = strftime('%Y-%m-%d', 'now')
            WHERE worksheet_id = ?
        ''', (next_col, next_row, worksheet_id))

def get_worksheet_info_by_id(conn, worksheet_id):
    cursor = conn.cursor()
    cursor.execute('''
    SELECT id, worksheet_id, next_col, next_row, last_updated
    FROM worksheet_info
    WHERE worksheet_id = ?
    ''', (worksheet_id,))
    return cursor.fetchone()