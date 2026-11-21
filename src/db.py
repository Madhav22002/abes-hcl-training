"""
Database Access Layer for SQLite (data/app.db)
Ensures robust absolute pathing and foreign key enforcement.
"""

import sqlite3
import pandas as pd
from pathlib import Path
from src.config import DB_PATH

def get_connection():
    """Establish and return SQLite connection with PRAGMA foreign_keys = ON."""
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database not found at {DB_PATH}. Please run src/generate_data.py first.")
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def execute_query(query: str, params: tuple = ()):
    """Execute arbitrary query and return a pandas DataFrame."""
    conn = get_connection()
    try:
        df = pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()
    return df

def get_table_data(table_name: str):
    """Retrieve full table safely."""
    allowed = ["campaigns", "channels", "audience_segments", "creatives", "daily_performance"]
    if table_name not in allowed:
        raise ValueError(f"Table {table_name} is not permitted.")
    return execute_query(f"SELECT * FROM {table_name}")

def get_analysis_view():
    """Retrieve full joined analysis view."""
    return execute_query("SELECT * FROM v_campaign_performance_details ORDER BY report_date ASC")

def verify_db_integrity():
    """Run SQLite integrity check and foreign key check."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("PRAGMA integrity_check;")
    integrity = cur.fetchone()[0]
    cur.execute("PRAGMA foreign_key_check;")
    fk_violations = cur.fetchall()
    conn.close()
    return {
        "integrity": integrity,
        "fk_violations_count": len(fk_violations),
        "status": "PASS" if integrity == "ok" and len(fk_violations) == 0 else "FAIL"
    }
