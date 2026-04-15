# coding=utf-8
import os, io, uuid, threading, json, hashlib
from flask import Flask, request, jsonify, send_file, render_template
from flask_socketio import SocketIO, join_room, leave_room, emit
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from converter import krutidev_to_unicode
import db as DB

# ── App setup ─────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config['SECRET_KEY']        = os.environ.get('SECRET_KEY', 'kd-editor-secret-2024')
app.config['UPLOAD_FOLDER']     = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

socketio = SocketIO(app, cors_allowed_origins='*', async_mode='threading')

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
progress_store = {}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
DB.init_db()


def allowed_file(f):
    return '.' in f and f.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def _needs_conversion(text: str) -> bool:
    dev = sum(1 for c in text if '\u0900' <= c <= '\u097F')
    if dev > len(text) * 0.3: return False
    non_alpha = sum(1 for c in text if not c.isalpha())
    return non_alpha != len(text)


# ── Excel helpers ─────────────────────────────────────────────────────────────

def excel_to_sheets(path: str) -> tuple[dict, list]:
    wb = load_workbook(path, data_only=True)
    sheets_data, sheet_names = {}, wb.sheetnames
    for name in sheet_names:
        ws = wb[name]
        rows, col_widths = [], {}
        for row in ws.iter_rows():
            rd = []
            for cell in row:
                val = cell.value
                val = '' if val is None else (str(val) if not isinstance(val, str) else val)
                rd.append(val)
                cl = get_column_letter(cell.column)
                if cl not in col_widths:
                    cd = ws.column_dimensions.get(cl)
                    col_widths[cl] = round(cd.width * 7) if cd and cd.width else 100
            rows.append(rd)
        mc = max((len(r) for r in rows), default=0)
        rows = [r + [''] * (mc - len(r)) for r in rows]
        sheets_data[name] = {'rows': rows, 'col_widths': col_widths,
                             'row_count': len(rows), 'col_count': mc}
    wb.close()
    return sheets_data, sheet_names


def sheets_to_excel(sheets: dict, convert_mode=False) -> io.BytesIO:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, sd in sheets.items():
        ws = wb.create_sheet(title=name[:31])
        for ri, row in enumerate(sd.get('rows', []), 1):
            for ci, val in enumerate(row, 1):
                cell = ws.cell(row=ri, column=ci)
                if convert_mode and isinstance(val, str) and val.strip():
                    cell.value = krutidev_to_unicode(val)
                else:
                    cell.value = val if val != '' else None
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def convert_excel_job(input_path, output_path, job_id):
    try:
        progress_store[job_id] = {'status': 'processing', 'progress': 0}
        wb = load_workbook(input_path)
        sheets = wb.sheetnames
        for si, sn in enumerate(sheets):
            ws = wb[sn]
            rows = list(ws.iter_rows())
            for ri, row in enumerate(rows):
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        v = cell.value.strip()
                        if v and _needs_conversion(v):
                            cell.value = krutidev_to_unicode(cell.value)
                pct = int(((si / len(sheets)) + (ri / len(rows) / len(sheets))) * 100)
                progress_store[job_id]['progress'] = min(pct, 95)
        wb.save(output_path)
        progress_store[job_id] = {'status': 'done', 'progress': 100, 'output': output_path}
    except Exception as e:
        progress_store[job_id] = {'status': 'error', 'message': str(e)}


# ── Pages ─────────────────────────────────────────────────────────────────────

@app.route('/')
def dashboard():
    return render_template('dashboard.html')


@app.route('/convert')
def convert_page():
    return render_template('index.html')


@app.route('/editor')
def editor_new():
    return render_template('editor.html', doc_id=None, access='edit', title='Untitled')


@app.route('/sheet/<doc_id>')
def sheet_view(doc_id):
    doc = DB.get_doc(doc_id)
    if not doc:
        return render_template('404.html'), 404
    return render_template('editor.html',
                           doc_id=doc_id,
                           access=doc['access'],
                           title=doc['title'])


# ── Legacy converter routes ───────────────────────────────────────────────────

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    file = request.files['file']
    if not file.filename or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type'}), 400
    filename = secure_filename(file.filename)
    job_id = str(uuid.uuid4())
    inp = os.path.join(app.config['UPLOAD_FOLDER'], f'{job_id}_input_{filename}')
    out = os.path.join(app.config['UPLOAD_FOLDER'], f'{job_id}_output_{filename}')
    file.save(inp)
    try:
        load_workbook(inp, read_only=True).close()
    except Exception:
        os.remove(inp)
        return jsonify({'error': 'Invalid or corrupted Excel file'}), 400
    t = threading.Thread(target=convert_excel_job, args=(inp, out, job_id), daemon=True)
    t.start()
    return jsonify({'job_id': job_id, 'filename': filename})


@app.route('/progress/<job_id>')
def progress(job_id):
    d = progress_store.get(job_id)
    return jsonify(d) if d else (jsonify({'status': 'not_found'}), 404)


@app.route('/download/<job_id>')
def download(job_id):
    d = progress_store.get(job_id)
    if not d or d.get('status') != 'done':
        return jsonify({'error': 'File not ready'}), 404
    op = d.get('output')
    if not op or not os.path.exists(op):
        return jsonify({'error': 'Output file not found'}), 404
    bn = os.path.basename(op)
    orig = bn.replace(f'{job_id}_output_', '')
    name, ext = os.path.splitext(orig)
    return send_file(op, as_attachment=True, download_name=f'{name}_unicode{ext}',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── Editor API ────────────────────────────────────────────────────────────────

@app.route('/api/load', methods=['POST'])
def api_load():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    file = request.files['file']
    if not file.filename or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type'}), 400
    filename = secure_filename(file.filename)
    fid = str(uuid.uuid4())
    path = os.path.join(app.config['UPLOAD_FOLDER'], f'{fid}_{filename}')
    file.save(path)
    try:
        sheets_data, sheet_names = excel_to_sheets(path)
        return jsonify({'file_id': fid, 'filename': filename,
                        'sheets': sheets_data, 'sheet_names': sheet_names})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/save', methods=['POST'])
def api_save():
    try:
        p = request.get_json(force=True)
        buf = sheets_to_excel(p.get('sheets', {}), p.get('convert', False))
        name = p.get('filename', 'edited').rsplit('.', 1)[0]
        suffix = '_unicode' if p.get('convert') else '_edited'
        return send_file(buf, as_attachment=True, download_name=f'{name}{suffix}.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── Share / Document API ──────────────────────────────────────────────────────

@app.route('/api/share', methods=['POST'])
def api_share():
    """Create or update a shared document. Returns shareable URL."""
    try:
        p = request.get_json(force=True)
        sheets  = p.get('sheets', {})
        names   = p.get('sheet_names', list(sheets.keys()))
        title   = p.get('title', 'Untitled')
        access  = p.get('access', 'edit')   # 'view' | 'edit'
        pwd     = p.get('password')
        doc_id  = p.get('doc_id')           # if updating existing

        pwd_hash = hashlib.sha256(pwd.encode()).hexdigest() if pwd else None

        if doc_id:
            DB.update_doc(doc_id, sheets, names, title)
        else:
            doc_id = DB.create_doc(title, sheets, names, access, pwd_hash)

        url = request.host_url.rstrip('/') + f'/sheet/{doc_id}'
        return jsonify({'doc_id': doc_id, 'url': url, 'access': access})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/doc/<doc_id>', methods=['GET'])
def api_get_doc(doc_id):
    """Load a shared document's data."""
    doc = DB.get_doc(doc_id)
    if not doc:
        return jsonify({'error': 'Document not found'}), 404
    pwd = request.args.get('pwd')
    if doc['password']:
        if not pwd or hashlib.sha256(pwd.encode()).hexdigest() != doc['password']:
            return jsonify({'error': 'Password required', 'protected': True}), 403
    return jsonify({'doc_id': doc_id, 'title': doc['title'],
                    'sheets': doc['data'], 'sheet_names': doc['sheet_names'],
                    'access': doc['access'], 'updated_at': doc['updated_at']})


@app.route('/api/doc/<doc_id>', methods=['DELETE'])
def api_delete_doc(doc_id):
    DB.delete_doc(doc_id)
    return jsonify({'ok': True})


@app.route('/api/docs', methods=['GET'])
def api_list_docs():
    return jsonify({'docs': DB.list_docs()})


@app.route('/api/doc/<doc_id>/download', methods=['GET'])
def api_download_doc(doc_id):
    doc = DB.get_doc(doc_id)
    if not doc:
        return jsonify({'error': 'Not found'}), 404
    convert = request.args.get('convert', 'false').lower() == 'true'
    buf = sheets_to_excel(doc['data'], convert)
    name = doc['title'].replace(' ', '_')
    suffix = '_unicode' if convert else ''
    return send_file(buf, as_attachment=True, download_name=f'{name}{suffix}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── WebSocket collaboration ───────────────────────────────────────────────────

# Track active users per room: { doc_id: { sid: {name, color} } }
_rooms: dict[str, dict] = {}
_COLORS = ['#6366f1','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#f97316','#ec4899']


def _room_users(doc_id):
    return list(_rooms.get(doc_id, {}).values())


@socketio.on('join')
def on_join(data):
    doc_id = data.get('doc_id')
    name   = data.get('name', 'Anonymous')
    if not doc_id:
        return
    join_room(doc_id)
    _rooms.setdefault(doc_id, {})
    color = _COLORS[len(_rooms[doc_id]) % len(_COLORS)]
    _rooms[doc_id][request.sid] = {'name': name, 'color': color, 'sid': request.sid}
    emit('users_update', {'users': _room_users(doc_id)}, to=doc_id)


@socketio.on('disconnect')
def on_disconnect():
    sid = request.sid
    for doc_id, users in list(_rooms.items()):
        if sid in users:
            del users[sid]
            emit('users_update', {'users': _room_users(doc_id)}, to=doc_id)
            if not users:
                del _rooms[doc_id]
            break


@socketio.on('cell_change')
def on_cell_change(data):
    """Broadcast a cell edit to all other users in the room."""
    doc_id = data.get('doc_id')
    if not doc_id:
        return
    # Persist the change to DB (last-write-wins)
    doc = DB.get_doc(doc_id)
    if doc and doc['access'] == 'edit':
        sheet  = data.get('sheet')
        row    = data.get('row')
        col    = data.get('col')
        value  = data.get('value', '')
        sheets = doc['data']
        if sheet in sheets and row < len(sheets[sheet]['rows']):
            while col >= len(sheets[sheet]['rows'][row]):
                sheets[sheet]['rows'][row].append('')
            sheets[sheet]['rows'][row][col] = value
            DB.update_doc(doc_id, sheets, doc['sheet_names'])
    # Broadcast to everyone else in the room
    emit('cell_change', data, to=doc_id, include_self=False)


@socketio.on('cursor_move')
def on_cursor(data):
    doc_id = data.get('doc_id')
    if doc_id:
        user = _rooms.get(doc_id, {}).get(request.sid, {})
        emit('cursor_move', {**data, 'user': user}, to=doc_id, include_self=False)


@socketio.on('title_change')
def on_title(data):
    doc_id = data.get('doc_id')
    title  = data.get('title', 'Untitled')
    if doc_id:
        doc = DB.get_doc(doc_id)
        if doc:
            DB.update_doc(doc_id, doc['data'], doc['sheet_names'], title)
        emit('title_change', {'title': title}, to=doc_id, include_self=False)


@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'File too large. Max 50MB'}), 413


if __name__ == '__main__':
    socketio.run(app, debug=True, port=5000, allow_unsafe_werkzeug=True)
