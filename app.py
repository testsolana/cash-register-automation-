from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import sqlite3
import uuid
from functools import wraps
from datetime import datetime
from io import BytesIO
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-me-in-production')


@app.template_filter('money')
def format_money(value):
    """Formats a number with a thin-space thousands separator and two decimal places:
    10000 -> '10 000.00' (a thin space, not a regular one), -12345.6 -> '-12 345.60'."""
    try:
        value = float(value)
    except (TypeError, ValueError):
        return value
    thin_space = '\u2009'  # thin space — narrower than a regular one, but still a visible separator
    formatted = f'{abs(value):,.2f}'.replace(',', thin_space)
    return ('-' if value < 0 else '') + formatted

DB_PATH = os.path.join(os.path.dirname(__file__), 'data', 'kasa.db')

# --- Users ---
# User1  -> view only (viewer)
# user2  -> view + edit (editor)
USERS = {
    'User1': {'password': 'password', 'role': 'viewer'},
    'user2': {'password': 'password', 'role': 'editor'},
}

CURRENCIES = ['UAH', 'USD', 'EUR']
CURRENCY_LABELS = {'UAH': 'Hryvnia (UAH)', 'USD': 'US Dollar (USD)', 'EUR': 'Euro (EUR)'}
CURRENCY_SYMBOLS = {'UAH': '₴', 'USD': '$', 'EUR': '€'}
TYPE_LABELS = {'prihid': 'Income', 'vydatok': 'Expense'}

OBJECTS_COUNT = 10  # number of sub-registers (objects)


def get_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()

    conn.execute('''
        CREATE TABLE IF NOT EXISTS objects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL
        )
    ''')

    conn.execute('''
        CREATE TABLE IF NOT EXISTS operations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            object_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            type TEXT NOT NULL,          -- 'prihid' (income) or 'vydatok' (expense)
            currency TEXT NOT NULL,      -- 'UAH' / 'USD' / 'EUR'
            amount REAL NOT NULL,
            rate REAL NOT NULL,          -- exchange rate at the time of the operation (usually 1 for UAH)
            description TEXT NOT NULL,   -- purpose of the operation, required
            created_by TEXT,
            created_at TEXT
        )
    ''')

    # --- Counterparties directory (who income is from / who expense is to) ---
    conn.execute('''
        CREATE TABLE IF NOT EXISTS contragents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            code TEXT,          -- company/tax ID (optional)
            phone TEXT,
            note TEXT,
            created_at TEXT
        )
    ''')

    # Migration: add the contragent_id column to an existing operations table,
    # in case we're upgrading from an older database version without this field.
    op_columns = [r['name'] for r in conn.execute('PRAGMA table_info(operations)').fetchall()]
    if 'contragent_id' not in op_columns:
        conn.execute('ALTER TABLE operations ADD COLUMN contragent_id INTEGER')

    # Migration: fields for transferring funds between cash registers.
    # transfer_object_id — id of the other cash register involved in the transfer (source/destination).
    # transfer_group_id  — shared identifier for the pair of operations (an expense in
    #                      one register + income in the other), so they can be edited/
    #                      deleted together and distinguished from regular operations.
    if 'transfer_object_id' not in op_columns:
        conn.execute('ALTER TABLE operations ADD COLUMN transfer_object_id INTEGER')
    if 'transfer_group_id' not in op_columns:
        conn.execute('ALTER TABLE operations ADD COLUMN transfer_group_id TEXT')

    # Populate 10 objects (cash registers), if they don't exist yet
    count = conn.execute('SELECT COUNT(*) c FROM objects').fetchone()['c']
    if count == 0:
        for i in range(1, OBJECTS_COUNT + 1):
            conn.execute('INSERT INTO objects (name) VALUES (?)', (f'Register {i}',))

    conn.commit()
    conn.close()


# Initialize the database as soon as the module loads
# (needed both for gunicorn and for a plain run)
init_db()


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'username' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


def editor_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if session.get('role') != 'editor':
            flash('You do not have permission to do this', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated


def get_objects():
    conn = get_db()
    objects = conn.execute('SELECT * FROM objects ORDER BY id').fetchall()
    conn.close()
    return objects


def get_current_object(object_id):
    conn = get_db()
    obj = conn.execute('SELECT * FROM objects WHERE id = ?', (object_id,)).fetchone()
    conn.close()
    return obj


# --- Counterparties ("from whom" / "to whom" directory) ---

def get_contragents(search=None):
    conn = get_db()
    if search:
        rows = conn.execute(
            'SELECT * FROM contragents WHERE name LIKE ? ORDER BY name COLLATE NOCASE',
            (f'%{search}%',)
        ).fetchall()
    else:
        rows = conn.execute('SELECT * FROM contragents ORDER BY name COLLATE NOCASE').fetchall()
    conn.close()
    return rows


def get_contragent(contragent_id):
    conn = get_db()
    row = conn.execute('SELECT * FROM contragents WHERE id = ?', (contragent_id,)).fetchone()
    conn.close()
    return row


def contragent_usage_count(contragent_id):
    conn = get_db()
    c = conn.execute(
        'SELECT COUNT(*) c FROM operations WHERE contragent_id = ?', (contragent_id,)
    ).fetchone()['c']
    conn.close()
    return c


# Virtual object "All registers" — id is always 0, not stored in the objects table.
TOTAL_OBJECT_ID = 0
TOTAL_OBJECT = {'id': TOTAL_OBJECT_ID, 'name': 'All registers'}


def get_tabs():
    """List of tabs for the top menu: 'All registers' first, then the real registers."""
    return [TOTAL_OBJECT] + [dict(o) for o in get_objects()]


# --- Operations log filters (currency, type, date range) ---

def parse_filters(source):
    """source — request.args or request.form (a MultiDict with .get/.getlist)."""
    active = bool(source.get('filtered'))
    if active:
        fcur = [c for c in source.getlist('fcur') if c in CURRENCIES]
        ftype = [t for t in source.getlist('ftype') if t in TYPE_LABELS]
        date_from = (source.get('date_from') or '').strip()
        date_to = (source.get('date_to') or '').strip()
        fctr_raw = source.get('fctr')
        try:
            fctr = int(fctr_raw) if fctr_raw else None
        except (TypeError, ValueError):
            fctr = None
    else:
        fcur = list(CURRENCIES)
        ftype = list(TYPE_LABELS.keys())
        date_from = ''
        date_to = ''
        fctr = None
    return {
        'active': active,
        'fcur': fcur,
        'ftype': ftype,
        'fctr': fctr,
        'date_from': date_from,
        'date_to': date_to,
    }


def filters_query_args(filters):
    """Builds a dict for url_for(**...) to pass the current filters in the query string."""
    qs = {}
    if filters['active']:
        qs['filtered'] = '1'
        if filters['fcur']:
            qs['fcur'] = filters['fcur']
        if filters['ftype']:
            qs['ftype'] = filters['ftype']
        if filters['fctr']:
            qs['fctr'] = filters['fctr']
        if filters['date_from']:
            qs['date_from'] = filters['date_from']
        if filters['date_to']:
            qs['date_to'] = filters['date_to']
    return qs


def index_redirect(object_id, **extra):
    """Redirect to index while keeping the active filters (passed as hidden
    fields in POST-action forms and read here from request.form)."""
    qs = filters_query_args(parse_filters(request.form))
    qs.update(extra)
    return redirect(url_for('index', object_id=object_id, **qs))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'username' in session:
        return redirect(url_for('index'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        user = USERS.get(username)
        if user and user['password'] == password:
            session['username'] = username
            session['role'] = user['role']
            return redirect(url_for('index'))
        flash('Invalid username or password', 'error')

    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


def resolve_object_id(object_id_raw):
    """Validates the object_id from the request and returns a correct id (defaults to the first object)."""
    objects = get_objects()
    if not objects:
        init_db()
        objects = get_objects()
    real_ids = [o['id'] for o in objects]
    if object_id_raw is None or (object_id_raw != TOTAL_OBJECT_ID and object_id_raw not in real_ids):
        return objects[0]['id']
    return object_id_raw


def query_filtered_operations(object_id, is_total, filters):
    """Returns (ops, balances, prihid_by_cur, vydatok_by_cur, uah_equivalent) with the
    filters applied — used both on the operations log page and in the Excel export, so the
    data always matches."""
    where_parts = []
    params = []

    if not is_total:
        where_parts.append('operations.object_id = ?')
        params.append(object_id)

    if filters['fcur']:
        placeholders = ','.join('?' * len(filters['fcur']))
        where_parts.append(f'operations.currency IN ({placeholders})')
        params.extend(filters['fcur'])
    else:
        where_parts.append('1=0')  # no currency selected — show nothing

    if filters['ftype']:
        placeholders = ','.join('?' * len(filters['ftype']))
        where_parts.append(f'operations.type IN ({placeholders})')
        params.extend(filters['ftype'])
    else:
        where_parts.append('1=0')  # no type selected — show nothing

    if filters['date_from']:
        where_parts.append('operations.date >= ?')
        params.append(filters['date_from'])

    if filters['date_to']:
        where_parts.append('operations.date <= ?')
        params.append(filters['date_to'])

    if filters['fctr']:
        where_parts.append('operations.contragent_id = ?')
        params.append(filters['fctr'])

    where_sql = ' AND '.join(where_parts) if where_parts else '1=1'

    conn = get_db()
    ops = conn.execute(
        'SELECT operations.*, objects.name AS object_name, contragents.name AS contragent_name, '
        'transfer_obj.name AS transfer_object_name '
        'FROM operations '
        'JOIN objects ON objects.id = operations.object_id '
        'LEFT JOIN contragents ON contragents.id = operations.contragent_id '
        'LEFT JOIN objects AS transfer_obj ON transfer_obj.id = operations.transfer_object_id '
        f'WHERE {where_sql} '
        'ORDER BY date DESC, operations.id DESC',
        params
    ).fetchall()
    conn.close()

    balances = {c: 0.0 for c in CURRENCIES}
    prihid_by_cur = {c: 0.0 for c in CURRENCIES}
    vydatok_by_cur = {c: 0.0 for c in CURRENCIES}
    uah_equivalent = 0.0

    for o in ops:
        signed = o['amount'] if o['type'] == 'prihid' else -o['amount']
        balances[o['currency']] += signed
        uah_equivalent += signed * o['rate']
        if o['type'] == 'prihid':
            prihid_by_cur[o['currency']] += o['amount']
        else:
            vydatok_by_cur[o['currency']] += o['amount']

    return ops, balances, prihid_by_cur, vydatok_by_cur, uah_equivalent


@app.route('/')
@login_required
def index():
    object_id = resolve_object_id(request.args.get('object_id', type=int))
    is_total = (object_id == TOTAL_OBJECT_ID)

    filters = parse_filters(request.args)

    current_object = TOTAL_OBJECT if is_total else get_current_object(object_id)
    ops, balances, prihid_by_cur, vydatok_by_cur, uah_equivalent = query_filtered_operations(
        object_id, is_total, filters
    )

    select_contragent = request.args.get('select_contragent', type=int)
    filter_qs = filters_query_args(filters)

    return render_template(
        'index.html',
        operations=ops,
        objects=get_tabs(),
        current_object=current_object,
        is_total=is_total,
        balances=balances,
        prihid_by_cur=prihid_by_cur,
        vydatok_by_cur=vydatok_by_cur,
        uah_equivalent=uah_equivalent,
        currencies=CURRENCIES,
        currency_labels=CURRENCY_LABELS,
        currency_symbols=CURRENCY_SYMBOLS,
        type_labels=TYPE_LABELS,
        role=session.get('role'),
        username=session.get('username'),
        today=datetime.now().strftime('%Y-%m-%d'),
        contragents=get_contragents(),
        select_contragent=select_contragent,
        filters=filters,
        filter_qs=filter_qs,
    )


def _write_ops_sheet(ws, current_object, is_total, ops_subset, filters, username, kind='full'):
    """Writes one Excel sheet: a header, the operations table and a summary.
    kind='full'   — all operations, summary with income/expenses/balance per currency.
    kind='income' — income only, summary with income totals only.
    kind='expense'— expenses only, summary with expense totals only.
    """
    base_font = 'Arial'
    title_font = Font(name=base_font, size=14, bold=True)
    meta_font = Font(name=base_font, size=10, italic=True, color='666666')
    header_font = Font(name=base_font, size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='2C333B', end_color='2C333B', fill_type='solid')
    normal_font = Font(name=base_font, size=11)
    green_font = Font(name=base_font, size=11, bold=True, color='1F7A4D')
    red_font = Font(name=base_font, size=11, bold=True, color='A83A2C')
    section_font = Font(name=base_font, size=12, bold=True)
    summary_header_font = Font(name=base_font, size=11, bold=True, color='FFFFFF')
    summary_header_fill = PatternFill(start_color='8A723A', end_color='8A723A', fill_type='solid')
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    row = 1
    ws.cell(row=row, column=1, value=f"Cash Register Ledger — {current_object['name']}").font = title_font
    row += 1
    ws.cell(row=row, column=1,
            value=f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M')}   User: {username}"
            ).font = meta_font
    row += 1

    filter_desc_parts = []
    if filters['active']:
        if filters['fcur'] and len(filters['fcur']) < len(CURRENCIES):
            filter_desc_parts.append('Currency: ' + ', '.join(filters['fcur']))
        if filters['ftype'] and len(filters['ftype']) < len(TYPE_LABELS):
            filter_desc_parts.append('Type: ' + ', '.join(TYPE_LABELS[t] for t in filters['ftype']))
        if filters['fctr']:
            c = get_contragent(filters['fctr'])
            if c:
                filter_desc_parts.append(f"Counterparty: {c['name']}")
        if filters['date_from']:
            filter_desc_parts.append(f"From date: {filters['date_from']}")
        if filters['date_to']:
            filter_desc_parts.append(f"To date: {filters['date_to']}")
    filter_desc = 'Filters: ' + ('; '.join(filter_desc_parts) if filter_desc_parts else 'none (all operations)')
    ws.cell(row=row, column=1, value=filter_desc).font = meta_font
    row += 2

    headers = (['Register'] if is_total else []) + [
        'Date', 'Type', 'Currency', 'Rate', 'Amount', 'From / To', 'Purpose'
    ]
    header_row = row
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    row += 1

    prihid_by_cur = {c: 0.0 for c in CURRENCIES}
    vydatok_by_cur = {c: 0.0 for c in CURRENCIES}

    for o in ops_subset:
        col = 1
        if is_total:
            ws.cell(row=row, column=col, value=o['object_name']).border = border
            col += 1
        ws.cell(row=row, column=col, value=o['date']).border = border
        col += 1
        ws.cell(row=row, column=col, value=TYPE_LABELS.get(o['type'], o['type'])).border = border
        col += 1
        ws.cell(row=row, column=col, value=o['currency']).border = border
        col += 1
        rate_cell = ws.cell(row=row, column=col, value=o['rate'])
        rate_cell.number_format = '#,##0.00'
        rate_cell.border = border
        col += 1
        signed_amount = o['amount'] if o['type'] == 'prihid' else -o['amount']
        amount_cell = ws.cell(row=row, column=col, value=signed_amount)
        amount_cell.number_format = '#,##0.00;-#,##0.00'
        amount_cell.font = green_font if o['type'] == 'prihid' else red_font
        amount_cell.border = border
        col += 1
        if o['contragent_name']:
            ctr_value = o['contragent_name']
        elif o['transfer_object_name']:
            ctr_value = f'⇄ Register "{o["transfer_object_name"]}"'
        else:
            ctr_value = '—'
        ws.cell(row=row, column=col, value=ctr_value).border = border
        col += 1
        ws.cell(row=row, column=col, value=o['description']).border = border
        row += 1

        if o['type'] == 'prihid':
            prihid_by_cur[o['currency']] += o['amount']
        else:
            vydatok_by_cur[o['currency']] += o['amount']

    if not ops_subset:
        ws.cell(row=row, column=1, value='No operations match the selected filters').font = meta_font
        row += 1

    row += 2

    ws.cell(row=row, column=1, value='SUMMARY').font = section_font
    row += 1

    if kind == 'full':
        summary_headers = ['Currency', 'Income (+)', 'Expenses (-)', 'Balance']
        summary_header_row = row
        for col, h in enumerate(summary_headers, start=1):
            cell = ws.cell(row=summary_header_row, column=col, value=h)
            cell.font = summary_header_font
            cell.fill = summary_header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        row += 1

        uah_equivalent = 0.0
        for cur in CURRENCIES:
            rate_for_cur = next((o['rate'] for o in ops_subset if o['currency'] == cur), RATES_DEFAULT.get(cur, 1.0))
            balance_cur = prihid_by_cur[cur] - vydatok_by_cur[cur]
            uah_equivalent += balance_cur * rate_for_cur
            ws.cell(row=row, column=1, value=cur).border = border
            c1 = ws.cell(row=row, column=2, value=prihid_by_cur[cur])
            c1.number_format = '#,##0.00'
            c1.font = green_font
            c1.border = border
            c2 = ws.cell(row=row, column=3, value=vydatok_by_cur[cur])
            c2.number_format = '#,##0.00'
            c2.font = red_font
            c2.border = border
            c3 = ws.cell(row=row, column=4, value=balance_cur)
            c3.number_format = '#,##0.00'
            c3.font = normal_font
            c3.border = border
            row += 1

        row += 1
        ws.cell(row=row, column=1, value='Total balance in UAH').font = section_font
        total_cell = ws.cell(row=row, column=2, value=uah_equivalent)
        total_cell.number_format = '#,##0.00'
        total_cell.font = Font(name=base_font, size=12, bold=True, color='8A723A')
    else:
        by_cur = prihid_by_cur if kind == 'income' else vydatok_by_cur
        sum_label = 'Income total' if kind == 'income' else 'Expenses total'
        amount_font = green_font if kind == 'income' else red_font
        total_label = 'Total income in UAH' if kind == 'income' else 'Total expenses in UAH'

        summary_headers = ['Currency', sum_label]
        summary_header_row = row
        for col, h in enumerate(summary_headers, start=1):
            cell = ws.cell(row=summary_header_row, column=col, value=h)
            cell.font = summary_header_font
            cell.fill = summary_header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        row += 1

        uah_total = 0.0
        for cur in CURRENCIES:
            rate_for_cur = next((o['rate'] for o in ops_subset if o['currency'] == cur), RATES_DEFAULT.get(cur, 1.0))
            uah_total += by_cur[cur] * rate_for_cur
            ws.cell(row=row, column=1, value=cur).border = border
            c1 = ws.cell(row=row, column=2, value=by_cur[cur])
            c1.number_format = '#,##0.00'
            c1.font = amount_font
            c1.border = border
            row += 1

        row += 1
        ws.cell(row=row, column=1, value=total_label).font = section_font
        total_cell = ws.cell(row=row, column=2, value=uah_total)
        total_cell.number_format = '#,##0.00'
        total_cell.font = Font(name=base_font, size=12, bold=True, color='8A723A')

    widths = ([14] if is_total else []) + [12, 10, 9, 11, 14, 26, 34]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


RATES_DEFAULT = {'UAH': 1.0, 'USD': 41.5, 'EUR': 45.2}


@app.route('/export_excel')
@login_required
def export_excel():
    object_id = resolve_object_id(request.args.get('object_id', type=int))
    is_total = (object_id == TOTAL_OBJECT_ID)

    filters = parse_filters(request.args)

    current_object = TOTAL_OBJECT if is_total else get_current_object(object_id)
    ops, balances, prihid_by_cur, vydatok_by_cur, uah_equivalent = query_filtered_operations(
        object_id, is_total, filters
    )

    wb = Workbook()
    ws_main = wb.active
    ws_main.title = 'Transactions'
    _write_ops_sheet(ws_main, current_object, is_total, ops, filters, session.get('username'), kind='full')

    # If both operation types (income and expense) are selected — add two more sheets:
    # income separately and expenses separately (respecting all the other active filters).
    if len(filters['ftype']) == 2:
        income_ops = [o for o in ops if o['type'] == 'prihid']
        expense_ops = [o for o in ops if o['type'] == 'vydatok']

        ws_income = wb.create_sheet(title='Income')
        _write_ops_sheet(ws_income, current_object, is_total, income_ops, filters, session.get('username'), kind='income')

        ws_expense = wb.create_sheet(title='Expense')
        _write_ops_sheet(ws_expense, current_object, is_total, expense_ops, filters, session.get('username'), kind='expense')

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    safe_name = ''.join(ch if ch.isalnum() else '_' for ch in current_object['name'])
    filename = f"kasa_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/add', methods=['POST'])
@login_required
@editor_required
def add_operation():
    object_id = request.form.get('object_id', type=int)
    op_type = request.form.get('type')
    currency = request.form.get('currency')
    amount_raw = request.form.get('amount')
    rate_raw = request.form.get('rate')
    description = request.form.get('description', '').strip()
    contragent_id = request.form.get('contragent_id', type=int)
    date = request.form.get('date') or datetime.now().strftime('%Y-%m-%d')

    objects_ids = [o['id'] for o in get_objects()]
    if object_id not in objects_ids:
        flash('Invalid cash register (object)', 'error')
        return redirect(url_for('index'))

    if op_type not in ('prihid', 'vydatok'):
        flash('Invalid operation type', 'error')
        return index_redirect(object_id)

    if currency not in CURRENCIES:
        flash('Invalid currency', 'error')
        return index_redirect(object_id)

    try:
        amount = float(str(amount_raw).replace(',', '.'))
        if amount <= 0:
            raise ValueError
    except (ValueError, TypeError, AttributeError):
        flash('Invalid operation amount', 'error')
        return index_redirect(object_id)

    try:
        rate = float(str(rate_raw).replace(',', '.'))
        if rate <= 0:
            raise ValueError
    except (ValueError, TypeError, AttributeError):
        flash('Rate is required and must be a positive number', 'error')
        return index_redirect(object_id)

    if currency == 'UAH':
        rate = 1.0  # for UAH the rate is always fixed, regardless of what the form sent
    else:
        rate = round(rate, 2)

    if not description:
        flash('Operation purpose is a required field', 'error')
        return index_redirect(object_id)

    if not contragent_id:
        label = 'The "From whom" field' if op_type == 'prihid' else 'The "To whom" field'
        flash(f'{label} is required — select a counterparty', 'error')
        return index_redirect(object_id)

    conn = get_db()
    contragent = conn.execute('SELECT id FROM contragents WHERE id = ?', (contragent_id,)).fetchone()
    if not contragent:
        conn.close()
        flash('The selected counterparty was not found. Choose another or add a new one', 'error')
        return index_redirect(object_id)

    conn.execute(
        'INSERT INTO operations '
        '(object_id, date, type, currency, amount, rate, description, contragent_id, created_by, created_at) '
        'VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
        (object_id, date, op_type, currency, amount, rate, description, contragent_id,
         session['username'], datetime.now().isoformat())
    )
    conn.commit()
    conn.close()

    flash('Operation added successfully', 'success')
    return index_redirect(object_id)


@app.route('/transfer', methods=['POST'])
@login_required
@editor_required
def transfer_between_objects():
    """Transfer funds between cash registers: creates two linked operations at once —
    an 'expense' in the source register ("To whom" = the destination register's name) and
    an 'income' in the destination register ("From whom" = the source register's name)."""
    from_object_id = request.form.get('from_object_id', type=int)
    to_object_id = request.form.get('to_object_id', type=int)
    currency = request.form.get('currency')
    amount_raw = request.form.get('amount')
    rate_raw = request.form.get('rate')
    description = request.form.get('description', '').strip()
    date = request.form.get('date') or datetime.now().strftime('%Y-%m-%d')

    objects_ids = [o['id'] for o in get_objects()]

    if from_object_id not in objects_ids or to_object_id not in objects_ids:
        flash('Invalid cash register (object) for the transfer', 'error')
        return index_redirect(from_object_id)

    if from_object_id == to_object_id:
        flash('The source and destination registers must be different', 'error')
        return index_redirect(from_object_id)

    if currency not in CURRENCIES:
        flash('Invalid currency', 'error')
        return index_redirect(from_object_id)

    try:
        amount = float(str(amount_raw).replace(',', '.'))
        if amount <= 0:
            raise ValueError
    except (ValueError, TypeError, AttributeError):
        flash('Invalid transfer amount', 'error')
        return index_redirect(from_object_id)

    try:
        rate = float(str(rate_raw).replace(',', '.'))
        if rate <= 0:
            raise ValueError
    except (ValueError, TypeError, AttributeError):
        flash('Rate is required and must be a positive number', 'error')
        return index_redirect(from_object_id)

    if currency == 'UAH':
        rate = 1.0
    else:
        rate = round(rate, 2)

    from_obj = get_current_object(from_object_id)
    to_obj = get_current_object(to_object_id)

    out_description = description or f'Funds transfer to register "{to_obj["name"]}"'
    in_description = description or f'Funds transfer from register "{from_obj["name"]}"'

    group_id = uuid.uuid4().hex
    now = datetime.now().isoformat()

    conn = get_db()
    conn.execute(
        'INSERT INTO operations '
        '(object_id, date, type, currency, amount, rate, description, contragent_id, '
        'transfer_object_id, transfer_group_id, created_by, created_at) '
        'VALUES (?, ?, ?, ?, ?, ?, ?, NULL, ?, ?, ?, ?)',
        (from_object_id, date, 'vydatok', currency, amount, rate, out_description,
         to_object_id, group_id, session['username'], now)
    )
    conn.execute(
        'INSERT INTO operations '
        '(object_id, date, type, currency, amount, rate, description, contragent_id, '
        'transfer_object_id, transfer_group_id, created_by, created_at) '
        'VALUES (?, ?, ?, ?, ?, ?, ?, NULL, ?, ?, ?, ?)',
        (to_object_id, date, 'prihid', currency, amount, rate, in_description,
         from_object_id, group_id, session['username'], now)
    )
    conn.commit()
    conn.close()

    flash(f'Transfer completed: "{from_obj["name"]}" -> "{to_obj["name"]}"', 'success')
    return index_redirect(from_object_id)


@app.route('/delete/<int:op_id>', methods=['POST'])
@login_required
@editor_required
def delete_operation(op_id):
    object_id = request.form.get('object_id', type=int)
    conn = get_db()
    op = conn.execute('SELECT id, transfer_group_id FROM operations WHERE id = ?', (op_id,)).fetchone()
    if op and op['transfer_group_id']:
        # This is a transfer operation — delete both paired operations (in both registers)
        # together, otherwise the paired register's balance would be left with an orphan entry.
        conn.execute('DELETE FROM operations WHERE transfer_group_id = ?', (op['transfer_group_id'],))
        flash('Transfer deleted (both paired operations)', 'success')
    else:
        conn.execute('DELETE FROM operations WHERE id = ?', (op_id,))
        flash('Operation deleted', 'success')
    conn.commit()
    conn.close()
    return index_redirect(object_id)


@app.route('/edit_description/<int:op_id>', methods=['POST'])
@login_required
@editor_required
def edit_description(op_id):
    object_id = request.form.get('object_id', type=int)
    description = request.form.get('description', '').strip()

    if not description:
        flash('Operation purpose cannot be empty', 'error')
        return index_redirect(object_id)

    conn = get_db()
    op = conn.execute('SELECT id FROM operations WHERE id = ?', (op_id,)).fetchone()
    if not op:
        conn.close()
        flash('Operation not found', 'error')
        return index_redirect(object_id)

    conn.execute('UPDATE operations SET description = ? WHERE id = ?', (description, op_id))
    conn.commit()
    conn.close()

    flash('Operation purpose updated', 'success')
    return index_redirect(object_id)


@app.route('/edit_contragent/<int:op_id>', methods=['POST'])
@login_required
@editor_required
def edit_operation_contragent(op_id):
    object_id = request.form.get('object_id', type=int)
    contragent_id = request.form.get('contragent_id', type=int)

    if not contragent_id:
        flash('Select a counterparty', 'error')
        return index_redirect(object_id)

    conn = get_db()
    op = conn.execute('SELECT id, transfer_group_id FROM operations WHERE id = ?', (op_id,)).fetchone()
    contragent = conn.execute('SELECT id FROM contragents WHERE id = ?', (contragent_id,)).fetchone()
    if not op:
        conn.close()
        flash('Operation not found', 'error')
        return index_redirect(object_id)
    if op['transfer_group_id']:
        conn.close()
        flash('This is a transfer between registers — "From / To" is set automatically', 'error')
        return index_redirect(object_id)
    if not contragent:
        conn.close()
        flash('Counterparty not found', 'error')
        return index_redirect(object_id)

    conn.execute('UPDATE operations SET contragent_id = ? WHERE id = ?', (contragent_id, op_id))
    conn.commit()
    conn.close()

    flash('Operation counterparty updated', 'success')
    return index_redirect(object_id)


# --- Counterparties: directory (like in accounting software) ---

@app.route('/contragents')
@login_required
def contragents_page():
    q = request.args.get('q', '').strip()
    return render_template(
        'contragents.html',
        contragents=get_contragents(search=q or None),
        q=q,
        role=session.get('role'),
        username=session.get('username'),
    )


@app.route('/contragents/add', methods=['POST'])
@login_required
@editor_required
def add_contragent():
    name = request.form.get('name', '').strip()
    note = request.form.get('note', '').strip()

    if not name:
        flash('Counterparty name / full name is a required field', 'error')
        return redirect(url_for('contragents_page'))

    conn = get_db()
    conn.execute(
        'INSERT INTO contragents (name, note, created_at) VALUES (?, ?, ?)',
        (name, note, datetime.now().isoformat())
    )
    conn.commit()
    conn.close()

    flash('Counterparty added', 'success')
    return redirect(url_for('contragents_page'))


@app.route('/contragents/quick_add', methods=['POST'])
@login_required
@editor_required
def quick_add_contragent():
    """Quickly add a counterparty right from the operation-entry form in the register."""
    name = request.form.get('name', '').strip()
    object_id = request.form.get('object_id', type=int)

    if not name:
        flash('Enter the counterparty name / full name', 'error')
        return index_redirect(object_id)

    conn = get_db()
    cur = conn.execute(
        'INSERT INTO contragents (name, note, created_at) VALUES (?, ?, ?)',
        (name, '', datetime.now().isoformat())
    )
    new_id = cur.lastrowid
    conn.commit()
    conn.close()

    flash(f'Counterparty "{name}" added and selected', 'success')
    return index_redirect(object_id, select_contragent=new_id)


@app.route('/contragents/<int:contragent_id>/edit', methods=['POST'])
@login_required
@editor_required
def edit_contragent(contragent_id):
    name = request.form.get('name', '').strip()
    note = request.form.get('note', '').strip()

    if not name:
        flash('Counterparty name / full name is a required field', 'error')
        return redirect(url_for('contragents_page'))

    conn = get_db()
    conn.execute(
        'UPDATE contragents SET name = ?, note = ? WHERE id = ?',
        (name, note, contragent_id)
    )
    conn.commit()
    conn.close()

    flash('Counterparty details updated', 'success')
    return redirect(url_for('contragents_page'))


@app.route('/contragents/<int:contragent_id>/delete', methods=['POST'])
@login_required
@editor_required
def delete_contragent(contragent_id):
    usage = contragent_usage_count(contragent_id)
    if usage > 0:
        flash(f'Cannot delete: counterparty is used in {usage} operations', 'error')
        return redirect(url_for('contragents_page'))

    conn = get_db()
    conn.execute('DELETE FROM contragents WHERE id = ?', (contragent_id,))
    conn.commit()
    conn.close()

    flash('Counterparty deleted', 'success')
    return redirect(url_for('contragents_page'))


@app.route('/objects/<int:object_id>/rename', methods=['POST'])
@login_required
@editor_required
def rename_object(object_id):
    name = request.form.get('name', '').strip()
    if name:
        conn = get_db()
        conn.execute('UPDATE objects SET name = ? WHERE id = ?', (name, object_id))
        conn.commit()
        conn.close()
        flash('Register name updated', 'success')
    else:
        flash('Name cannot be empty', 'error')
    return index_redirect(object_id)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
