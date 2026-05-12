import sqlite3
from pathlib import Path
from datetime import datetime

import os
DB_PATH = Path("/tmp/receipts.db") if os.environ.get("VERCEL") else Path(__file__).parent / "receipts.db"


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS receipts (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                voucher_no       TEXT    UNIQUE NOT NULL,
                store_name       TEXT,
                payment_datetime TEXT,
                total_amount     REAL,
                category         TEXT DEFAULT '기타',
                filename         TEXT,
                created_at       TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS items (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                receipt_id  INTEGER NOT NULL,
                name        TEXT,
                quantity    REAL,
                unit_price  REAL,
                FOREIGN KEY (receipt_id) REFERENCES receipts(id) ON DELETE CASCADE
            )
        """)


def next_voucher_no():
    year = datetime.now().year
    with get_conn() as conn:
        row = conn.execute(
            "SELECT COUNT(*) AS cnt FROM receipts WHERE voucher_no LIKE ?",
            (f"{year}-%",),
        ).fetchone()
        return f"{year}-{row['cnt'] + 1:04d}"
