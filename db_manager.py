import os
import sqlite3

def setup_database(db_folder="data", db_name="pcm_database.db"):
    """Check if SQLite DB exists, if not create it."""
    if not os.path.exists(db_folder):
        os.makedirs(db_folder)

    db_path = os.path.join(db_folder, db_name)

    if not os.path.exists(db_path):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pcm (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                type_of_report TEXT,
                created_at TEXT,
                modified_at TEXT,
                report_blob BLOB
            )
        """)
        conn.commit()
        conn.close()

    return db_path


def insert_report(db_path, report_type, created_at, modified_at, report_blob):
    """Insert report into database."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO pcm (type_of_report, created_at, modified_at, report_blob)
        VALUES (?, ?, ?, ?)
    """, (report_type, created_at, modified_at, report_blob))
    conn.commit()
    conn.close()
