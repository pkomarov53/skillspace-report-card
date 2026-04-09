import sqlite3

DB_NAME = 'data/academy_database.db'

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS utp_modules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            utp_name TEXT,
            module_name TEXT,
            hours TEXT,
            control_form TEXT
        )
    ''')
    conn.commit()
    conn.close()