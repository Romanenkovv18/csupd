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
        # max_quantity в stock
        cols = {row[1] for row in conn.execute("PRAGMA table_info(stock)")}
        if 'max_quantity' not in cols:
            conn.execute("ALTER TABLE stock ADD COLUMN max_quantity INTEGER")
            conn.execute(
                "UPDATE stock SET max_quantity = ("
                "  SELECT monthly_plan FROM parts WHERE parts.id = stock.part_id"
                ") WHERE max_quantity IS NULL"
            )
            conn.commit()

        # category в parts
        part_cols = {row[1] for row in conn.execute("PRAGMA table_info(parts)")}
        if 'category' not in part_cols:
            conn.execute("ALTER TABLE parts ADD COLUMN category TEXT")
            conn.execute("""
                UPDATE parts SET category = CASE
                  WHEN name LIKE '%кольцо%' OR name LIKE '%кольца%' THEN 'Кольца'
                  WHEN name LIKE '%шайба%' OR name LIKE '%шайбы%' THEN 'Шайбы'
                  WHEN name LIKE '%прокладк%' THEN 'Прокладки'
                  WHEN name LIKE '%замок%' OR name LIKE '%замки%' THEN 'Замки'
                  WHEN name LIKE '%шплинт%' THEN 'Шплинты'
                  WHEN name LIKE '%манжет%' THEN 'Манжеты'
                  WHEN name LIKE '%контровк%' THEN 'Контровки'
                  WHEN name LIKE '%болт%' THEN 'Болты'
                  WHEN name LIKE '%гайк%' THEN 'Гайки'
                  WHEN name LIKE '%шпилька%' OR name LIKE '%шпильки%' THEN 'Шпильки'
                  ELSE 'Прочее'
                END
                WHERE category IS NULL
            """)
            conn.commit()

        # audit_log
        conn.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY,
                table_name TEXT,
                record_id INTEGER,
                action TEXT,
                old_value TEXT,
                new_value TEXT,
                changed_by TEXT,
                changed_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                ip_address TEXT
            )
        """)

        # is_cancelled в movements
        mov_cols = {row[1] for row in conn.execute("PRAGMA table_info(movements)")}
        if 'is_cancelled' not in mov_cols:
            conn.execute("ALTER TABLE movements ADD COLUMN is_cancelled INTEGER DEFAULT 0")
            conn.execute("ALTER TABLE movements ADD COLUMN cancel_reason TEXT")
            conn.execute("ALTER TABLE movements ADD COLUMN cancelled_by TEXT")
            conn.execute("ALTER TABLE movements ADD COLUMN cancelled_at DATETIME")
            conn.commit()

        # updated_by / updated_at в stock
        if 'updated_by' not in {row[1] for row in conn.execute("PRAGMA table_info(stock)")}:
            conn.execute("ALTER TABLE stock ADD COLUMN updated_by TEXT")
            conn.commit()

        # updated_by / updated_at в parts
        if 'updated_by' not in {row[1] for row in conn.execute("PRAGMA table_info(parts)")}:
            conn.execute("ALTER TABLE parts ADD COLUMN updated_by TEXT")
            conn.execute("ALTER TABLE parts ADD COLUMN updated_at DATETIME")
            conn.commit()

    finally:
        conn.close()
