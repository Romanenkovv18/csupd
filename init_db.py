"""
Инициализация БД: создание таблиц, загрузка деталей из CSV, тестовые остатки.
Запустить один раз: python init_db.py
"""

import csv
import math
import os

from werkzeug.security import generate_password_hash

from db import DB_PATH, get_db

SCHEMA = """
CREATE TABLE IF NOT EXISTS parts (
    id              INTEGER PRIMARY KEY,
    article         TEXT NOT NULL,
    name            TEXT NOT NULL,
    engine_type     TEXT,
    unit            TEXT DEFAULT 'шт',
    monthly_plan    INTEGER DEFAULT 0,
    yellow_threshold INTEGER,
    red_threshold   INTEGER
);

CREATE TABLE IF NOT EXISTS storage_cells (
    id           INTEGER PRIMARY KEY,
    cell_code    TEXT UNIQUE NOT NULL,
    shelf_number TEXT,
    level_number TEXT,
    cell_number  TEXT,
    notes        TEXT
);

CREATE TABLE IF NOT EXISTS stock (
    id         INTEGER PRIMARY KEY,
    part_id    INTEGER REFERENCES parts(id),
    cell_id    INTEGER REFERENCES storage_cells(id),
    quantity   INTEGER DEFAULT 0,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS movements (
    id             INTEGER PRIMARY KEY,
    part_id        INTEGER REFERENCES parts(id),
    cell_id        INTEGER REFERENCES storage_cells(id),
    operation_type TEXT,
    quantity       INTEGER,
    engine_number  TEXT,
    assembly_stage TEXT,
    operator       TEXT,
    created_at     DATETIME DEFAULT CURRENT_TIMESTAMP,
    notes          TEXT
);

CREATE TABLE IF NOT EXISTS replenishment_requests (
    id           INTEGER PRIMARY KEY,
    part_id      INTEGER REFERENCES parts(id),
    quantity     INTEGER,
    status       TEXT DEFAULT 'Создана',
    created_by   TEXT,
    created_at   DATETIME DEFAULT CURRENT_TIMESTAMP,
    completed_at DATETIME
);

CREATE TABLE IF NOT EXISTS users (
    id            INTEGER PRIMARY KEY,
    username      TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    full_name     TEXT,
    role          TEXT DEFAULT 'sborshik'
);
"""


def parse_cell_code(cell_code):
    parts = cell_code.split('-')
    if len(parts) == 3:
        return parts[0], parts[1], parts[2]
    return cell_code, '', ''


def calc_thresholds(monthly_plan):
    """
    Плановый дневной расход = monthly_plan / 22 рабочих дня
    Красный порог  = 1 день (срок поставки)
    Жёлтый порог   = 4 дня  (срок + 3 дня страхового запаса)
    """
    if monthly_plan == 0:
        return 0, 0
    daily = monthly_plan / 22
    red = math.ceil(daily * 1)
    yellow = math.ceil(daily * 4)
    return red, yellow


def test_quantity(i, red_threshold, yellow_threshold, monthly_plan):
    """
    Первые 3 позиции → красный, следующие 8 → жёлтый, остальные → зелёный.
    """
    if i < 3:  # красный: строго ниже красного порога
        return max(0, red_threshold - 1)
    elif i < 11:  # жёлтый: между красным и жёлтым порогами
        return red_threshold + max(1, (yellow_threshold - red_threshold) // 2)
    else:  # зелёный: выше жёлтого порога
        return yellow_threshold + max(1, monthly_plan // 4)


def main():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print(f"Удалён старый файл: {DB_PATH}")

    conn = get_db()
    conn.executescript(SCHEMA)

    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'parts_list.csv')
    with open(csv_path, encoding='utf-8') as f:
        rows = list(csv.DictReader(f))

    red_count = yellow_count = green_count = 0

    for i, row in enumerate(rows):
        monthly_plan = int(row['monthly_plan'])
        red_thr, yellow_thr = calc_thresholds(monthly_plan)

        cur = conn.execute(
            "INSERT INTO parts (article, name, engine_type, unit, monthly_plan, yellow_threshold, red_threshold) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (row['article'], row['name'], row['engine_type'], row['unit'],
             monthly_plan, yellow_thr, red_thr),
        )
        part_id = cur.lastrowid

        cell_code = row['cell_code']
        shelf, level, cell = parse_cell_code(cell_code)
        conn.execute(
            "INSERT OR IGNORE INTO storage_cells (cell_code, shelf_number, level_number, cell_number) "
            "VALUES (?, ?, ?, ?)",
            (cell_code, shelf, level, cell),
        )
        cell_id = conn.execute(
            "SELECT id FROM storage_cells WHERE cell_code = ?", (cell_code,)
        ).fetchone()[0]

        qty = test_quantity(i, red_thr, yellow_thr, monthly_plan)
        conn.execute(
            "INSERT INTO stock (part_id, cell_id, quantity) VALUES (?, ?, ?)",
            (part_id, cell_id, qty),
        )

        # Считаем для отчёта
        if qty < red_thr:
            red_count += 1
        elif qty < yellow_thr:
            yellow_count += 1
        else:
            green_count += 1

    # Демо-пользователи
    demo_users = [
        ('sborshik',   '1234',  'Иванов Иван Иванович',        'sborshik'),
        ('kladovshik', '1234',  'Петрова Мария Сергеевна',     'kladovshik'),
        ('master',     '1234',  'Сидоров Алексей Николаевич',  'master'),
        ('admin',      'admin', 'Администратор системы',        'admin'),
    ]
    for username, password, full_name, role in demo_users:
        conn.execute(
            "INSERT INTO users (username, password_hash, full_name, role) VALUES (?, ?, ?, ?)",
            (username, generate_password_hash(password), full_name, role),
        )

    # Демо-движения (исторические данные за последние 2 недели)
    # (part_id, cell_id, op, qty, engine, stage, operator, days_ago, notes)
    demo_movements = [
        (4,  4,  'расход', 15, '117-03-4215', 'Сборка двигателя',    'Иванов И.И.',    1,  None),
        (6,  6,  'расход', 20, '117-03-4215', 'Сборка узла',         'Иванов И.И.',    1,  None),
        (8,  8,  'расход', 10, '117-03-4216', 'Сборка двигателя',    'Петров С.В.',    2,  None),
        (1,  1,  'приход', 60, None,           None,                  'Петрова М.С.',   2,  'Накладная №45'),
        (2,  2,  'расход',  1, '117-03-4214', 'Входной контроль',    'Иванов И.И.',    3,  None),
        (7,  7,  'приход', 30, None,           None,                  'Петрова М.С.',   3,  'Накладная №44'),
        (5,  5,  'расход', 10, '117-03-4213', 'Сборка двигателя',    'Сидоров А.Н.',   5,  None),
        (10, 10, 'расход',  1, '117-03-4213', 'Дефектация',          'Сидоров А.Н.',   5,  None),
        (11, 11, 'приход', 20, None,           None,                  'Петрова М.С.',   7,  'Накладная №43'),
        (3,  3,  'расход',  1, '117-03-4212', 'Разборка',            'Иванов И.И.',    7,  None),
        (12, 12, 'расход',  2, '117-03-4212', 'Сборка узла',         'Петров С.В.',    9,  None),
        (9,  9,  'приход', 20, None,           None,                  'Петрова М.С.',  10,  'Накладная №42'),
        (13, 13, 'расход',  2, '117-03-4211', 'Стендовые испытания', 'Сидоров А.Н.',  10,  None),
        (6,  6,  'приход',200, None,           None,                  'Петрова М.С.',  12,  'Накладная №41'),
        (4,  4,  'расход', 20, '117-03-4210', 'Сборка двигателя',    'Иванов И.И.',   14,  None),
    ]
    for dm in demo_movements:
        pid, cid, op, qty, eng, stage, oper, days, notes = dm
        conn.execute(
            "INSERT INTO movements "
            "(part_id, cell_id, operation_type, quantity, engine_number, "
            "assembly_stage, operator, notes, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now', ?))",
            (pid, cid, op, qty, eng, stage, oper, notes, f'-{days} days'),
        )

    conn.commit()
    conn.close()

    print(f"\nБД создана: {DB_PATH}")
    print(f"Деталей загружено : {len(rows)}")
    print(f"  [RED]    Красных  : {red_count}")
    print(f"  [YELLOW] Жёлтых   : {yellow_count}")
    print(f"  [GREEN]  Зелёных  : {green_count}")
    print(f"\nПользователи:")
    for u, p, _, r in demo_users:
        print(f"  {u} / {p}  ({r})")


if __name__ == '__main__':
    main()
