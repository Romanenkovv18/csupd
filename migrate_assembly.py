"""
Миграция: добавляет таблицы модуля сборки в существующую БД.
Безопасно — не трогает parts, stock, movements, users.
Запуск: python migrate_assembly.py
"""
import os
from db import DB_PATH, get_db

ASSEMBLY_SCHEMA = """
CREATE TABLE IF NOT EXISTS assembly_stages (
    id           INTEGER PRIMARY KEY,
    engine_type  TEXT NOT NULL,
    stage_number INTEGER NOT NULL,
    stage_name   TEXT NOT NULL,
    description  TEXT
);

CREATE TABLE IF NOT EXISTS stage_parts (
    id       INTEGER PRIMARY KEY,
    stage_id INTEGER REFERENCES assembly_stages(id),
    part_id  INTEGER REFERENCES parts(id),
    quantity INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS engine_assemblies (
    id            INTEGER PRIMARY KEY,
    engine_number TEXT UNIQUE NOT NULL,
    engine_type   TEXT NOT NULL,
    current_stage INTEGER DEFAULT 1,
    status        TEXT DEFAULT 'В работе',
    assembler     TEXT NOT NULL,
    started_at    DATETIME DEFAULT CURRENT_TIMESTAMP,
    completed_at  DATETIME
);

CREATE TABLE IF NOT EXISTS assembly_history (
    id           INTEGER PRIMARY KEY,
    assembly_id  INTEGER REFERENCES engine_assemblies(id),
    stage_id     INTEGER REFERENCES assembly_stages(id),
    completed_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    assembler    TEXT,
    is_cancelled INTEGER DEFAULT 0
);
"""

# 7 операций сборки ТВ3-117
STAGES_TV3 = [
    (1, 'Входное устройство',   'Проверка и установка деталей входного устройства'),
    (2, 'Сборка компрессора',   'Установка лопаток и деталей компрессора'),
    (3, 'Камера сгорания',      'Монтаж камеры сгорания'),
    (4, 'Турбина',              'Установка рабочих лопаток турбины'),
    (5, 'Свободная турбина',    'Сборка и установка свободной турбины'),
    (6, 'Масляная система',     'Монтаж масляных уплотнений и трубопроводов'),
    (7, 'Финальная сборка',     'Контровка, затяжка и финальный контроль'),
]

# Детали для каждой операции: {stage_number: [(part_id, qty), ...]}
# part_id соответствует порядковому номеру строки в parts_list.csv
STAGE_PARTS = {
    1: [(2, 2),  (13, 1)],           # Кольцо уплотнительное ×2, Прокладка ×1
    2: [(1, 3),  (8, 11), (3, 1)],   # Замок ×3, Шайба контровочная ×11, Кольцо стопорное ×1
    3: [(4, 11), (14, 1)],           # Шайба контровочная двойная ×11, Прокладка регулировочная ×1
    4: [(35, 4), (37, 21)],          # Замок контровочный ×4, Шайба стопорная двусторонняя ×21
    5: [(8, 15), (39, 6)],           # Шайба контровочная ×15, Кольцо ×6
    6: [(10, 2), (11, 1), (13, 2)],  # Манжета ×2, Кольцо уплотнительное ×1, Прокладка ×2
    7: [(51, 30), (22, 75)],         # Шплинт ×30, Шайба замковая ×75
}

# 3 тестовые сборки на разных этапах
TEST_ASSEMBLIES = [
    ('ТВ3-117 №А-2024-001', 'ТВ3-117', 1, 'В работе', 'sborshik'),  # только начал
    ('ТВ3-117 №А-2024-002', 'ТВ3-117', 3, 'В работе', 'sborshik'),  # выполнил 2 операции
    ('ТВ3-117 №А-2024-003', 'ТВ3-117', 7, 'В работе', 'sborshik'),  # финальная операция
]


def get_stage_id(conn, engine_type, stage_number):
    row = conn.execute(
        "SELECT id FROM assembly_stages WHERE engine_type=? AND stage_number=?",
        (engine_type, stage_number)
    ).fetchone()
    return row[0] if row else None


def main():
    if not os.path.exists(DB_PATH):
        print(f"БД не найдена: {DB_PATH}")
        print("Сначала запустите: python init_db.py")
        return

    conn = get_db()
    conn.executescript(ASSEMBLY_SCHEMA)

    # Сброс и переиизагрузка данных (идемпотентно)
    conn.execute("DELETE FROM assembly_history")
    conn.execute("DELETE FROM engine_assemblies")
    conn.execute("DELETE FROM stage_parts")
    conn.execute("DELETE FROM assembly_stages")

    # ── Операции ТВ3-117 ─────────────────────────────────────────────────────
    for num, name, desc in STAGES_TV3:
        conn.execute(
            "INSERT INTO assembly_stages (engine_type, stage_number, stage_name, description) "
            "VALUES (?, ?, ?, ?)",
            ('ТВ3-117', num, name, desc),
        )

    # ── Детали операций ──────────────────────────────────────────────────────
    part_errors = []
    for stage_num, parts in STAGE_PARTS.items():
        stage_id = get_stage_id(conn, 'ТВ3-117', stage_num)
        for part_id, qty in parts:
            row = conn.execute("SELECT name FROM parts WHERE id=?", (part_id,)).fetchone()
            if not row:
                part_errors.append(f"  part_id={part_id} не найден (операция {stage_num})")
                continue
            conn.execute(
                "INSERT INTO stage_parts (stage_id, part_id, quantity) VALUES (?, ?, ?)",
                (stage_id, part_id, qty),
            )

    # ── Тестовые сборки ──────────────────────────────────────────────────────
    for eng_num, eng_type, cur_stage, status, assembler in TEST_ASSEMBLIES:
        conn.execute(
            "INSERT INTO engine_assemblies "
            "(engine_number, engine_type, current_stage, status, assembler, started_at) "
            "VALUES (?, ?, ?, ?, ?, datetime('now', '-3 days'))",
            (eng_num, eng_type, cur_stage, status, assembler),
        )

    # ── История для сборки №2 (выполнены операции 1 и 2) ────────────────────
    asm2_id = conn.execute(
        "SELECT id FROM engine_assemblies WHERE engine_number='ТВ3-117 №А-2024-002'"
    ).fetchone()[0]
    for sn in [1, 2]:
        stage_id = get_stage_id(conn, 'ТВ3-117', sn)
        conn.execute(
            "INSERT INTO assembly_history (assembly_id, stage_id, assembler, completed_at) "
            "VALUES (?, ?, ?, datetime('now', ?))",
            (asm2_id, stage_id, 'sborshik', f'-{3 - sn} days'),
        )

    # ── История для сборки №3 (выполнены операции 1–6) ───────────────────────
    asm3_id = conn.execute(
        "SELECT id FROM engine_assemblies WHERE engine_number='ТВ3-117 №А-2024-003'"
    ).fetchone()[0]
    for sn in range(1, 7):
        stage_id = get_stage_id(conn, 'ТВ3-117', sn)
        conn.execute(
            "INSERT INTO assembly_history (assembly_id, stage_id, assembler, completed_at) "
            "VALUES (?, ?, ?, datetime('now', ?))",
            (asm3_id, stage_id, 'sborshik', f'-{7 - sn} days'),
        )

    conn.commit()
    conn.close()

    # ── Отчёт ────────────────────────────────────────────────────────────────
    print(f"\nМиграция выполнена: {DB_PATH}")
    if part_errors:
        print("ПРЕДУПРЕЖДЕНИЯ:")
        for e in part_errors:
            print(e)
    print(f"\nОпераций ТВ3-117 : {len(STAGES_TV3)}")
    for num, name, _ in STAGES_TV3:
        parts_count = len(STAGE_PARTS.get(num, []))
        print(f"  {num}. {name} ({parts_count} деталей)")
    print(f"\nТестовые сборки ({len(TEST_ASSEMBLIES)}):")
    for eng, _, cur, status, _ in TEST_ASSEMBLIES:
        print(f"  {eng}  —  операция {cur}/{len(STAGES_TV3)}  [{status}]")


if __name__ == '__main__':
    main()
