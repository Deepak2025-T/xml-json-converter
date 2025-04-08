import psycopg2
from datetime import datetime

DB_CONFIG = {
    "dbname": "your_db",
    "user": "your_user",
    "password": "your_password",
    "host": "localhost",
    "port": "5432"
}

def get_connection():
    return psycopg2.connect(**DB_CONFIG)

def initialize_db():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS file_metadata (
            id SERIAL PRIMARY KEY,
            file_name TEXT,
            file_type TEXT,
            source TEXT,
            url TEXT,
            storage_path TEXT,
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)
    conn.commit()
    cur.close()
    conn.close()

def save_file_record(file_name, file_type, source, storage_path, url=None):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO file_metadata (file_name, file_type, source, url, storage_path, created_at)
        VALUES (%s, %s, %s, %s, %s, %s)
    """, (file_name, file_type, source, url, storage_path, datetime.now()))
    conn.commit()
    cur.close()
    conn.close()
