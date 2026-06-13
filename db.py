import os
import sqlite3

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'csupd.db')


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def ensure_migrations():
    """Безопасно применяет недостающие изменения схемы к существующей БД."""
    conn = sqlite3.connect(DB_PATH)
    try:
        cols = {row[1] for row in conn.execute("PRAGMA table_info(stock)")}
        if 'max_quantity' not in cols:
            conn.execute("ALTER TABLE stock ADD COLUMN max_quantity INTEGER")
            conn.execute(
                "UPDATE stock SET max_quantity = ("
                "  SELECT monthly_plan FROM parts WHERE parts.id = stock.part_id"
                ") WHERE max_quantity IS NULL"
            )
            conn.commit()
    finally:
        conn.close()
