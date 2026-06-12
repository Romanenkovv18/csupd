import math
import os
from datetime import datetime
from functools import wraps
from io import BytesIO

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from flask import (Flask, flash, redirect, render_template,
                   request, send_file, session, url_for)
from werkzeug.security import check_password_hash

from db import get_db

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'csupd-dev-secret-2024')

# ── Константы ────────────────────────────────────────────────────────────────

ASSEMBLY_STAGES = [
    'Входной контроль', 'Разборка', 'Дефектация',
    'Сборка узла', 'Сборка двигателя',
    'Стендовые испытания', 'Выходной контроль',
]

ROLE_LEVEL = {'sborshik': 1, 'kladovshik': 2, 'master': 3, 'admin': 4}
ROLE_NAMES = {
    'sborshik':   'Сборщик',
    'kladovshik': 'Кладовщик',
    'master':     'Мастер',
    'admin':      'Администратор',
}

# ── Вспомогательные функции авторизации ──────────────────────────────────────

def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login', next=request.full_path))
        return f(*args, **kwargs)
    return wrapper


def role_required(min_role):
    """Пропускает только пользователей с уровнем роли >= min_role."""
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if 'user_id' not in session:
                return redirect(url_for('login', next=request.full_path))
            if ROLE_LEVEL.get(session.get('role'), 0) < ROLE_LEVEL.get(min_role, 99):
                flash('Недостаточно прав для этого действия.', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return wrapper
    return decorator


@app.context_processor
def inject_user():
    """Вставляет current_user и ROLE_LEVEL во все шаблоны."""
    u = None
    if 'user_id' in session:
        u = {
            'id':        session['user_id'],
            'username':  session['username'],
            'role':      session['role'],
            'full_name': session['full_name'],
            'role_name': ROLE_NAMES.get(session['role'], session['role']),
            'level':     ROLE_LEVEL.get(session['role'], 0),
        }
    return {'current_user': u, 'ROLE_LEVEL': ROLE_LEVEL}


# ── Авторизация ───────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))

    next_url = request.args.get('next', '')

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        next_url = request.form.get('next', '')

        db = get_db()
        user = db.execute(
            'SELECT * FROM users WHERE username = ?', (username,)
        ).fetchone()
        db.close()

        if user and check_password_hash(user['password_hash'], password):
            session.clear()
            session['user_id']   = user['id']
            session['username']  = user['username']
            session['role']      = user['role']
            session['full_name'] = user['full_name'] or user['username']
            flash(f'Добро пожаловать, {session["full_name"]}!', 'success')
            return redirect(next_url or url_for('dashboard'))

        flash('Неверный логин или пароль.', 'danger')

    return render_template('login.html', next=next_url)


@app.route('/logout')
def logout():
    name = session.get('full_name', '')
    session.clear()
    flash(f'Вы вышли из системы{(", " + name) if name else ""}.', 'success')
    return redirect(url_for('login'))


# ── Дашборд ───────────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def dashboard():
    q = request.args.get('q', '').strip()
    db = get_db()

    sql = '''
        SELECT p.id, p.article, p.name, p.engine_type, p.unit,
               p.monthly_plan, p.red_threshold, p.yellow_threshold,
               s.quantity, sc.cell_code,
               CASE
                 WHEN s.quantity < p.red_threshold   THEN 'red'
                 WHEN s.quantity < p.yellow_threshold THEN 'yellow'
                 ELSE 'green'
               END AS status,
               CASE
                 WHEN s.quantity < p.red_threshold   THEN 1
                 WHEN s.quantity < p.yellow_threshold THEN 2
                 ELSE 3
               END AS sort_order
        FROM parts p
        JOIN stock s  ON s.part_id = p.id
        JOIN storage_cells sc ON sc.id = s.cell_id
    '''
    params = []
    if q:
        sql += ' WHERE (p.name LIKE ? OR p.article LIKE ?)'
        params = [f'%{q}%', f'%{q}%']
    sql += ' ORDER BY sort_order, p.name'

    parts = db.execute(sql, params).fetchall()

    red_count    = sum(1 for p in parts if p['status'] == 'red')
    yellow_count = sum(1 for p in parts if p['status'] == 'yellow')
    green_count  = sum(1 for p in parts if p['status'] == 'green')

    movements = db.execute('''
        SELECT m.operation_type, m.quantity, m.operator, m.created_at,
               p.name AS part_name, p.unit
        FROM movements m
        JOIN parts p ON m.part_id = p.id
        ORDER BY m.created_at DESC
        LIMIT 10
    ''').fetchall()

    db.close()
    return render_template('dashboard.html',
        parts=parts,
        red_count=red_count,
        yellow_count=yellow_count,
        green_count=green_count,
        movements=movements,
        q=q,
    )


# ── Расход ────────────────────────────────────────────────────────────────────

@app.route('/raskhod', methods=['GET', 'POST'])
@login_required
def raskhod():
    db = get_db()

    if request.method == 'POST':
        part_id        = request.form.get('part_id', '').strip()
        quantity_raw   = request.form.get('quantity', '').strip()
        engine_number  = request.form.get('engine_number', '').strip()
        assembly_stage = request.form.get('assembly_stage', '').strip()
        operator       = request.form.get('operator', '').strip()

        errors = []
        if not part_id:       errors.append('Выберите деталь.')
        if not engine_number: errors.append('Укажите номер двигателя.')
        if not assembly_stage:errors.append('Выберите этап сборки.')
        if not operator:      errors.append('Укажите исполнителя.')
        if not quantity_raw or not quantity_raw.isdigit() or int(quantity_raw) < 1:
            errors.append('Количество должно быть целым числом больше нуля.')

        if errors:
            for e in errors:
                flash(e, 'danger')
        else:
            quantity = int(quantity_raw)
            stock = db.execute(
                'SELECT s.id, s.quantity, s.cell_id FROM stock s WHERE s.part_id = ?',
                (part_id,)
            ).fetchone()

            if not stock:
                flash('Деталь не найдена в остатках.', 'danger')
            elif quantity > stock['quantity']:
                flash(
                    f'Недостаточно на складе. '
                    f'Запрошено: {quantity}, в наличии: {stock["quantity"]} шт.',
                    'danger'
                )
            else:
                db.execute(
                    'UPDATE stock SET quantity = quantity - ?, updated_at = CURRENT_TIMESTAMP '
                    'WHERE id = ?',
                    (quantity, stock['id'])
                )
                db.execute(
                    'INSERT INTO movements '
                    '(part_id, cell_id, operation_type, quantity, '
                    'engine_number, assembly_stage, operator) '
                    'VALUES (?, ?, ?, ?, ?, ?, ?)',
                    (part_id, stock['cell_id'], 'расход',
                     quantity, engine_number, assembly_stage, operator)
                )
                db.commit()
                part = db.execute(
                    'SELECT name, article FROM parts WHERE id = ?', (part_id,)
                ).fetchone()
                flash(
                    f'Расход зафиксирован: {part["name"]} ({part["article"]}) — '
                    f'{quantity} шт. · двигатель {engine_number}',
                    'success'
                )
                db.close()
                return redirect(url_for('dashboard'))

    parts = db.execute('''
        SELECT p.id, p.article, p.name, p.unit, s.quantity
        FROM parts p
        JOIN stock s ON s.part_id = p.id
        ORDER BY p.name
    ''').fetchall()

    db.close()
    return render_template('raskhod.html',
        parts=parts,
        assembly_stages=ASSEMBLY_STAGES,
        form=request.form,
        default_operator=session.get('full_name', ''),
    )


# ── Приход ────────────────────────────────────────────────────────────────────

@app.route('/prikhod', methods=['GET', 'POST'])
@role_required('kladovshik')
def prikhod():
    db = get_db()

    if request.method == 'POST':
        part_id      = request.form.get('part_id', '').strip()
        quantity_raw = request.form.get('quantity', '').strip()
        cell_id      = request.form.get('cell_id', '').strip()
        operator     = request.form.get('operator', '').strip()
        notes        = request.form.get('notes', '').strip()

        errors = []
        if not part_id:  errors.append('Выберите деталь.')
        if not operator: errors.append('Укажите исполнителя.')
        if not quantity_raw or not quantity_raw.isdigit() or int(quantity_raw) < 1:
            errors.append('Количество должно быть целым числом больше нуля.')

        if errors:
            for e in errors:
                flash(e, 'danger')
        else:
            quantity = int(quantity_raw)
            stock = db.execute(
                'SELECT s.id, s.quantity, s.cell_id FROM stock s WHERE s.part_id = ?',
                (part_id,)
            ).fetchone()

            target_cell_id = int(cell_id) if cell_id else (stock['cell_id'] if stock else None)

            if stock:
                db.execute(
                    'UPDATE stock SET quantity = quantity + ?, updated_at = CURRENT_TIMESTAMP '
                    'WHERE id = ?',
                    (quantity, stock['id'])
                )
            else:
                db.execute(
                    'INSERT INTO stock (part_id, cell_id, quantity) VALUES (?, ?, ?)',
                    (part_id, target_cell_id, quantity)
                )

            db.execute(
                'INSERT INTO movements '
                '(part_id, cell_id, operation_type, quantity, operator, notes) '
                'VALUES (?, ?, ?, ?, ?, ?)',
                (part_id, target_cell_id, 'приход', quantity, operator, notes or None)
            )
            db.commit()
            part = db.execute(
                'SELECT name, article FROM parts WHERE id = ?', (part_id,)
            ).fetchone()
            flash(
                f'Приход зафиксирован: {part["name"]} ({part["article"]}) — {quantity} шт.',
                'success'
            )
            db.close()
            return redirect(url_for('dashboard'))

    parts = db.execute('''
        SELECT p.id, p.article, p.name, p.unit,
               s.quantity, s.cell_id, sc.cell_code
        FROM parts p
        JOIN stock s  ON s.part_id = p.id
        JOIN storage_cells sc ON sc.id = s.cell_id
        ORDER BY p.name
    ''').fetchall()

    cells = db.execute(
        'SELECT id, cell_code FROM storage_cells ORDER BY cell_code'
    ).fetchall()

    db.close()
    return render_template('prikhod.html',
        parts=parts,
        cells=cells,
        form=request.form,
        default_operator=session.get('full_name', ''),
    )


# ── Заявки на пополнение ──────────────────────────────────────────────────────

@app.route('/zayavki')
@role_required('kladovshik')
def zayavki():
    status_filter = request.args.get('status', '')
    db = get_db()

    counts = {r['status']: r['cnt'] for r in db.execute(
        'SELECT status, COUNT(*) AS cnt FROM replenishment_requests GROUP BY status'
    ).fetchall()}

    sql = '''
        SELECT rr.id, rr.quantity, rr.status, rr.created_by,
               rr.created_at, rr.completed_at,
               p.name AS part_name, p.article, p.unit,
               p.monthly_plan, s.quantity AS stock_qty
        FROM replenishment_requests rr
        JOIN parts p ON rr.part_id = p.id
        JOIN stock s ON s.part_id = p.id
    '''
    params = []
    if status_filter:
        sql += ' WHERE rr.status = ?'
        params.append(status_filter)
    sql += ' ORDER BY rr.created_at DESC'

    requests_list = db.execute(sql, params).fetchall()
    db.close()

    return render_template('zayavki.html',
        requests_list=requests_list,
        counts=counts,
        status_filter=status_filter,
    )


@app.route('/zayavki/create', methods=['POST'])
@role_required('kladovshik')
def zayavki_create():
    part_id    = request.form.get('part_id', '').strip()
    quantity   = request.form.get('quantity', '').strip()
    created_by = request.form.get('created_by', '').strip()

    if not part_id or not quantity or not quantity.isdigit() or int(quantity) < 1:
        flash('Некорректные данные для создания заявки.', 'danger')
        return redirect(url_for('dashboard'))

    db = get_db()
    db.execute(
        'INSERT INTO replenishment_requests (part_id, quantity, status, created_by) '
        'VALUES (?, ?, ?, ?)',
        (part_id, int(quantity), 'Создана',
         created_by or session.get('full_name', 'Неизвестно'))
    )
    db.commit()
    part = db.execute('SELECT name, article FROM parts WHERE id = ?', (part_id,)).fetchone()
    db.close()

    flash(f'Заявка создана: {part["name"]} ({part["article"]}) — {quantity} шт.', 'success')
    return redirect(request.form.get('next', url_for('dashboard')))


@app.route('/zayavki/<int:req_id>/update', methods=['POST'])
@role_required('kladovshik')
def zayavki_update(req_id):
    new_status = request.form.get('status', '')
    if new_status not in ('В работе', 'Выполнена'):
        flash('Недопустимый статус.', 'danger')
        return redirect(url_for('zayavki'))

    db = get_db()
    if new_status == 'Выполнена':
        db.execute(
            'UPDATE replenishment_requests '
            'SET status = ?, completed_at = CURRENT_TIMESTAMP WHERE id = ?',
            (new_status, req_id)
        )
    else:
        db.execute(
            'UPDATE replenishment_requests SET status = ? WHERE id = ?',
            (new_status, req_id)
        )
    db.commit()
    db.close()

    flash(f'Статус заявки #{req_id} обновлён: {new_status}', 'success')
    return redirect(url_for('zayavki'))


# ── Журнал операций ───────────────────────────────────────────────────────────

@app.route('/journal')
@login_required
def journal():
    q         = request.args.get('q', '').strip()
    op_type   = request.args.get('op_type', '')
    date_from = request.args.get('date_from', '')
    date_to   = request.args.get('date_to', '')
    page      = max(1, int(request.args.get('page', 1) or 1))
    per_page  = 50

    db = get_db()

    conditions = ['1=1']
    params     = []

    if q:
        conditions.append(
            '(p.name LIKE ? OR p.article LIKE ? OR m.operator LIKE ? OR m.engine_number LIKE ?)'
        )
        params += [f'%{q}%'] * 4
    if op_type in ('приход', 'расход'):
        conditions.append('m.operation_type = ?')
        params.append(op_type)
    if date_from:
        conditions.append('DATE(m.created_at) >= ?')
        params.append(date_from)
    if date_to:
        conditions.append('DATE(m.created_at) <= ?')
        params.append(date_to)

    where = ' AND '.join(conditions)

    total = db.execute(
        f'SELECT COUNT(*) FROM movements m JOIN parts p ON m.part_id = p.id WHERE {where}',
        params,
    ).fetchone()[0]

    movements = db.execute(f'''
        SELECT m.id, m.operation_type, m.quantity, m.engine_number,
               m.assembly_stage, m.operator, m.created_at, m.notes,
               p.name AS part_name, p.article, p.unit,
               sc.cell_code
        FROM movements m
        JOIN parts p ON m.part_id = p.id
        LEFT JOIN storage_cells sc ON sc.id = m.cell_id
        WHERE {where}
        ORDER BY m.created_at DESC
        LIMIT ? OFFSET ?
    ''', params + [per_page, (page - 1) * per_page]).fetchall()

    db.close()

    return render_template('journal.html',
        movements=movements,
        total=total,
        page=page,
        total_pages=max(1, math.ceil(total / per_page)),
        per_page=per_page,
        q=q,
        op_type=op_type,
        date_from=date_from,
        date_to=date_to,
    )


# ── Экспорт в Excel ───────────────────────────────────────────────────────────

def _xl_style(bold=False, bg=None, color='000000', wrap=False, align='left'):
    """Возвращает именованный стиль для ячейки openpyxl."""
    thin = Side(style='thin', color='D0D0D0')
    return {
        'font':      Font(bold=bold, color=color, size=10),
        'fill':      PatternFill('solid', fgColor=bg) if bg else PatternFill(),
        'border':    Border(left=thin, right=thin, top=thin, bottom=thin),
        'alignment': Alignment(wrap_text=wrap, horizontal=align, vertical='center'),
    }


def _apply(cell, style):
    for attr, val in style.items():
        setattr(cell, attr, val)


def _xl_send(wb, filename_prefix):
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(
        buf,
        as_attachment=True,
        download_name=f'{filename_prefix}_{ts}.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/journal/export')
@login_required
def journal_export():
    q         = request.args.get('q', '').strip()
    op_type   = request.args.get('op_type', '')
    date_from = request.args.get('date_from', '')
    date_to   = request.args.get('date_to', '')

    db = get_db()
    conditions, params = ['1=1'], []

    if q:
        conditions.append(
            '(p.name LIKE ? OR p.article LIKE ? OR m.operator LIKE ? OR m.engine_number LIKE ?)'
        )
        params += [f'%{q}%'] * 4
    if op_type in ('приход', 'расход'):
        conditions.append('m.operation_type = ?')
        params.append(op_type)
    if date_from:
        conditions.append('DATE(m.created_at) >= ?')
        params.append(date_from)
    if date_to:
        conditions.append('DATE(m.created_at) <= ?')
        params.append(date_to)

    where = ' AND '.join(conditions)
    rows = db.execute(f'''
        SELECT m.created_at, m.operation_type, p.article, p.name, m.quantity, p.unit,
               sc.cell_code, m.engine_number, m.assembly_stage, m.operator, m.notes
        FROM movements m
        JOIN parts p ON m.part_id = p.id
        LEFT JOIN storage_cells sc ON sc.id = m.cell_id
        WHERE {where}
        ORDER BY m.created_at DESC
    ''', params).fetchall()
    db.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Журнал операций'

    # Заголовки
    headers = [
        '№', 'Дата', 'Время', 'Тип операции', 'Артикул', 'Наименование',
        'Кол-во', 'Ед.изм.', 'Ячейка', 'Двигатель', 'Этап сборки',
        'Исполнитель', 'Примечание',
    ]
    col_w = [5, 12, 8, 13, 18, 32, 8, 7, 11, 16, 22, 22, 28]

    hdr_style = _xl_style(bold=True, bg='1A3A6E', color='FFFFFF', align='center')
    for col, (h, w) in enumerate(zip(headers, col_w), 1):
        cell = ws.cell(row=1, column=col, value=h)
        _apply(cell, hdr_style)
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = 'A2'

    # Стили строк
    s_in  = _xl_style(bg='EAF7EF')   # приход — светло-зелёный
    s_out = _xl_style(bg='FDEEEE')   # расход — светло-красный
    s_def = _xl_style()

    for idx, row in enumerate(rows, 1):
        dt_str = row[0] or ''
        op     = row[1] or ''
        style  = s_in if op == 'приход' else (s_out if op == 'расход' else s_def)

        values = [
            idx,
            dt_str[:10],
            dt_str[11:16],
            op,
            row[2], row[3],
            row[4], row[5],
            row[6] or '',
            row[7] or '',
            row[8] or '',
            row[9] or '',
            row[10] or '',
        ]
        r = idx + 1
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=r, column=col, value=val)
            _apply(cell, style)
        ws.row_dimensions[r].height = 16

    return _xl_send(wb, 'journal')


@app.route('/export/stock')
@login_required
def stock_export():
    db = get_db()
    rows = db.execute('''
        SELECT p.article, p.name, p.engine_type, p.unit,
               s.quantity, p.red_threshold, p.yellow_threshold,
               p.monthly_plan, sc.cell_code,
               CASE
                 WHEN s.quantity < p.red_threshold   THEN 'Критично'
                 WHEN s.quantity < p.yellow_threshold THEN 'Внимание'
                 ELSE 'Норма'
               END AS status
        FROM parts p
        JOIN stock s  ON s.part_id = p.id
        JOIN storage_cells sc ON sc.id = s.cell_id
        ORDER BY
          CASE WHEN s.quantity < p.red_threshold   THEN 1
               WHEN s.quantity < p.yellow_threshold THEN 2
               ELSE 3 END,
          p.name
    ''').fetchall()
    db.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Остатки деталей'

    headers = [
        'Артикул', 'Наименование', 'Двигатель', 'Ед.изм.',
        'Остаток', 'Порог красный', 'Порог жёлтый',
        'План/мес', 'Ячейка', 'Статус',
    ]
    col_w = [18, 32, 10, 7, 9, 14, 14, 10, 12, 10]

    hdr_style = _xl_style(bold=True, bg='1A3A6E', color='FFFFFF', align='center')
    for col, (h, w) in enumerate(zip(headers, col_w), 1):
        cell = ws.cell(row=1, column=col, value=h)
        _apply(cell, hdr_style)
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = 'A2'

    s_red    = _xl_style(bg='FDEEEE')
    s_yellow = _xl_style(bg='FFF9E6')
    s_green  = _xl_style()

    for idx, row in enumerate(rows, 1):
        status = row['status']
        style  = s_red if status == 'Критично' else (s_yellow if status == 'Внимание' else s_green)
        values = [
            row['article'], row['name'], row['engine_type'], row['unit'],
            row['quantity'], row['red_threshold'], row['yellow_threshold'],
            row['monthly_plan'], row['cell_code'], status,
        ]
        r = idx + 1
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=r, column=col, value=val)
            _apply(cell, style)
        ws.row_dimensions[r].height = 16

    return _xl_send(wb, 'stock')


# ── Запуск ────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'csupd.db')
    if not os.path.exists(db_path):
        from init_db import main as init_db
        init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)
