from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from pathlib import Path
import pandas as pd
import io, os, hmac, secrets, re
from datetime import datetime
from urllib.parse import urlparse
from functools import wraps

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / 'data'
UPLOAD_PATH = DATA_DIR / 'progres_lintor.xlsx'
DATA_DIR.mkdir(exist_ok=True)
app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY') or secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=os.environ.get('COOKIE_SECURE', '0') == '1',
    PERMANENT_SESSION_LIFETIME=60 * 60 * 8,
    MAX_CONTENT_LENGTH=10 * 1024 * 1024,  # Maksimal unggahan 10 MB
)
# Kredensial demo untuk penggunaan lokal. Ubah melalui environment variable sebelum dipakai bersama.
LOGIN_USERNAME = os.environ.get('LINTOR_USERNAME', 'admin')
LOGIN_PASSWORD = os.environ.get('LINTOR_PASSWORD', 'Lintor@2026')

@app.before_request
def require_login():
    endpoint = request.endpoint or ''
    if endpoint in ('login', 'static', 'healthz') or endpoint.startswith('static'):
        return None
    if not session.get('logged_in'):
        return redirect(url_for('login', next=request.full_path if request.query_string else request.path))
    return None


STAGES = ['Pemberkasan', 'SU', 'Pengumuman', 'Pengesahan Pengumuman', 'K1']
NUMERIC_COLS = ['Target'] + STAGES
RENAME_MAP = {
    'Program': 'Program', 'Desa': 'Desa', 'Kecamatan': 'Kecamatan', 'Target': 'Target',
    'Pemberkasan': 'Pemberkasan', 'SU': 'SU', 'Pengumuman': 'Pengumuman',
    'Pengesahan Pengumuman': 'Pengesahan Pengumuman', 'Pengesahan\nPengumuman': 'Pengesahan Pengumuman', 'K1': 'K1'
}


BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
BULAN_LOOKUP = {m.lower(): i + 1 for i, m in enumerate(BULAN_ID)}
BULAN_LOOKUP.update({'agu': 8, 'agt': 8, 'agust': 8, 'okt': 10, 'des': 12, 'jan': 1, 'feb': 2, 'mar': 3,
                     'apr': 4, 'jun': 6, 'jul': 7, 'sep': 9, 'sept': 9, 'nov': 11, 'nop': 11, 'january': 1,
                     'february': 2, 'march': 3, 'may': 5, 'june': 6, 'july': 7, 'august': 8, 'october': 10,
                     'december': 12})


def format_date_id(d, short=False):
    name = BULAN_ID[d.month - 1]
    return f"{d.day} {name[:3] if short else name} {d.year}"


def extract_data_date(file_path=UPLOAD_PATH):
    """Baca tanggal progres dari judul file Excel (mis. '(21 OKTOBER 2026)').
    Bila tidak ditemukan, pakai waktu terakhir file diubah/diimpor."""
    try:
        raw = pd.read_excel(file_path, header=None, nrows=10)
        for i in range(len(raw)):
            for v in raw.iloc[i].tolist():
                if pd.isna(v):
                    continue
                if isinstance(v, (datetime, pd.Timestamp)):
                    return pd.Timestamp(v).to_pydatetime(), 'excel'
                text = str(v).strip()
                m = re.search(r'(\d{1,2})\s+([A-Za-z]{3,10})\.?\s+(\d{4})', text)
                if m and m.group(2).lower() in BULAN_LOOKUP:
                    return datetime(int(m.group(3)), BULAN_LOOKUP[m.group(2).lower()], int(m.group(1))), 'excel'
                m = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', text)
                if m:
                    return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))), 'excel'
                m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', text)
                if m:
                    return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))), 'excel'
    except Exception:
        pass
    try:
        return datetime.fromtimestamp(Path(file_path).stat().st_mtime), 'file'
    except Exception:
        return datetime.now(), 'file'


_DATE_CACHE = {}


def get_data_date(file_path=UPLOAD_PATH):
    try:
        key = (str(file_path), Path(file_path).stat().st_mtime)
    except Exception:
        key = (str(file_path), None)
    if key not in _DATE_CACHE:
        _DATE_CACHE.clear()
        _DATE_CACHE[key] = extract_data_date(file_path)
    return _DATE_CACHE[key]


@app.context_processor
def inject_data_date():
    d, source = get_data_date()
    return {'data_date': format_date_id(d), 'data_date_short': format_date_id(d, short=True),
            'data_date_source': source}


def load_data(file_path=UPLOAD_PATH):
    """Load the workbook and locate the real header row; tolerate title rows above the table."""
    try:
        raw = pd.read_excel(file_path, header=None)
    except Exception as exc:
        raise ValueError(f'File tidak dapat dibaca sebagai Excel: {exc}')
    header_idx = None
    for i in range(min(12, len(raw))):
        vals = [str(v).strip().replace('\n', ' ') for v in raw.iloc[i].tolist()]
        if {'Program', 'Desa', 'Kecamatan', 'Target'}.issubset(set(vals)):
            header_idx = i
            break
    if header_idx is None:
        # Support a clean spreadsheet whose column headers are already on row 1.
        try:
            df = pd.read_excel(file_path)
        except Exception:
            raise ValueError('Header tidak ditemukan. Pastikan kolom Program, Desa, Kecamatan, dan Target tersedia.')
    else:
        headers = [str(v).strip().replace('\n', ' ') if pd.notna(v) else '' for v in raw.iloc[header_idx].tolist()]
        df = raw.iloc[header_idx + 1:].copy()
        df.columns = headers
    df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
    df = df.rename(columns=RENAME_MAP)
    needed = ['Program', 'Desa', 'Kecamatan', 'Target'] + STAGES
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError('Kolom wajib belum tersedia: ' + ', '.join(missing))
    df = df[needed].copy()
    # Drop aggregate rows (e.g. Jumlah) and rows without a village name.
    df = df[df['Desa'].notna()]
    df = df[~df['Desa'].astype(str).str.strip().str.lower().isin(['jumlah', 'total'])]
    df = df[df['Program'].notna()]
    for c in ['Program', 'Desa', 'Kecamatan']:
        df[c] = df[c].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df.loc[df[c].str.casefold().isin(['', 'nan', 'none', 'nat']), c] = pd.NA
    # Do not allow aggregate/placeholder rows to become selectable locations.
    df = df[df['Desa'].notna() & df['Kecamatan'].notna() & df['Program'].notna()]
    df = df[~df['Desa'].str.casefold().isin(['jumlah', 'total'])]
    for c in NUMERIC_COLS:
        df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', '', regex=False), errors='coerce')
    # Blank cells remain missing; do not silently interpret them as zero.
    df = df.dropna(subset=['Program', 'Desa', 'Kecamatan', 'Target'])
    if df.empty:
        raise ValueError('Tidak ada baris data yang valid setelah validasi.')
    for c in NUMERIC_COLS:
        if (df[c].dropna() < 0).any():
            raise ValueError(f'Kolom {c} memiliki angka negatif. Periksa file sumber.')
    df['Tanggal Data'] = format_date_id(get_data_date(file_path)[0])
    return df.reset_index(drop=True)


def calculate(df, weights=(0.40, 0.35, 0.25)):
    """Transparent, rule-based SAW demo. Higher score = higher monitoring priority."""
    out = df.copy()
    # Calculate rates only when target is positive. Missing K1 is kept missing.
    out['Persentase K1'] = (out['K1'] / out['Target'].where(out['Target'] > 0)) * 100
    out['Sisa Target'] = (out['Target'] - out['K1']).clip(lower=0)
    # Difference between Pengumuman and K1 is a diagnostic proxy, not an official SLA.
    out['Selisih Pengumuman-K1'] = (out['Pengumuman'] - out['K1']).clip(lower=0)
    out['Status Data'] = out[STAGES].isna().any(axis=1).map({True: 'Perlu verifikasi', False: 'Lengkap'})

    # Criteria: C1 low K1 rate (cost); C2 remaining target (benefit); C3 late-stage gap (benefit).
    # Normalize cost with (max-x)/(max-min); robust for equal values and zero values.
    rate = out['Persentase K1'].fillna(out['Persentase K1'].max() if out['Persentase K1'].notna().any() else 0)
    gap = out['Sisa Target'].fillna(0)
    late_gap = out['Selisih Pengumuman-K1'].fillna(0)
    def norm_cost(s):
        lo, hi = float(s.min()), float(s.max())
        return pd.Series(1.0, index=s.index) if hi == lo else (hi - s) / (hi - lo)
    def norm_benefit(s):
        lo, hi = float(s.min()), float(s.max())
        return pd.Series(0.0, index=s.index) if hi == lo else (s - lo) / (hi - lo)
    c1, c2, c3 = norm_cost(rate), norm_benefit(gap), norm_benefit(late_gap)
    w1, w2, w3 = weights
    out['Skor Prioritas'] = (c1*w1 + c2*w2 + c3*w3) * 100
    out['Prioritas'] = pd.cut(out['Skor Prioritas'], bins=[-0.001, 33.333, 66.666, 100.001], labels=['Rendah', 'Sedang', 'Tinggi'], include_lowest=True).astype(str)
    out['Alasan Prioritas'] = out.apply(explain_priority, axis=1)
    return out.sort_values(['Skor Prioritas', 'Sisa Target'], ascending=[False, False]).reset_index(drop=True)


def explain_priority(row):
    reasons = []
    if pd.notna(row['Persentase K1']) and row['Persentase K1'] < 25:
        reasons.append('rasio K1 masih rendah')
    if row['Sisa Target'] > 0:
        reasons.append(f'sisa target {int(row["Sisa Target"]):,}')
    if row['Selisih Pengumuman-K1'] > 0:
        reasons.append(f'selisih Pengumuman–K1 {int(row["Selisih Pengumuman-K1"]):,}')
    if row['Status Data'] != 'Lengkap':
        reasons.append('data tahapan perlu diverifikasi')
    return '; '.join(reasons) if reasons else 'pantau progres rutin'


def fmt_int(value):
    if pd.isna(value): return '—'
    return f'{int(round(value)):,}'.replace(',', '.')

app.jinja_env.filters['idnum'] = fmt_int
app.jinja_env.filters['pct'] = lambda x: '—' if pd.isna(x) else f'{x:.1f}%'.replace('.', ',')


def _clean_filter_value(value):
    """Normalize a query-string value without accepting invalid/stale options."""
    value = str(value or 'Semua').strip()
    return value if value else 'Semua'


def get_view_data(raw=None):
    """Apply only valid filters that exist in the currently uploaded Excel workbook.

    Invalid/stale child selections are reset instead of silently returning an empty or
    misleading result after the parent filter changes.
    """
    raw = load_data() if raw is None else raw.copy()
    program = _clean_filter_value(request.args.get('program', 'Semua'))
    kecamatan = _clean_filter_value(request.args.get('kecamatan', 'Semua'))
    desa = _clean_filter_value(request.args.get('desa', 'Semua'))

    valid_programs = set(raw['Program'].dropna().astype(str).str.strip().unique())
    if program != 'Semua' and program not in valid_programs:
        program = 'Semua'

    by_program = raw if program == 'Semua' else raw[raw['Program'] == program]
    valid_kecamatan = set(by_program['Kecamatan'].dropna().astype(str).str.strip().unique())
    if kecamatan != 'Semua' and kecamatan not in valid_kecamatan:
        kecamatan = 'Semua'

    by_location = by_program if kecamatan == 'Semua' else by_program[by_program['Kecamatan'] == kecamatan]
    valid_desa = set(by_location['Desa'].dropna().astype(str).str.strip().unique())
    if desa != 'Semua' and desa not in valid_desa:
        desa = 'Semua'

    df = by_location if desa == 'Semua' else by_location[by_location['Desa'] == desa]
    return df.copy(), program, kecamatan, desa


def filter_options(raw, selected_program='Semua', selected_kecamatan='Semua'):
    """Build dropdown options from valid rows in the uploaded workbook."""
    programs = sorted(raw['Program'].dropna().astype(str).str.strip().unique().tolist())
    df = raw.copy()
    if selected_program != 'Semua':
        df = df[df['Program'] == selected_program]
    kecamatans = sorted(df['Kecamatan'].dropna().astype(str).str.strip().unique().tolist())
    if selected_kecamatan != 'Semua':
        df = df[df['Kecamatan'] == selected_kecamatan]
    desas = sorted(df['Desa'].dropna().astype(str).str.strip().unique().tolist())
    return programs, kecamatans, desas

@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('logged_in'):
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        username_ok = hmac.compare_digest(username, LOGIN_USERNAME)
        password_ok = hmac.compare_digest(password, LOGIN_PASSWORD)
        if username_ok and password_ok:
            session.clear()
            session['logged_in'] = True
            session['username'] = username
            session.permanent = True
            flash('Berhasil masuk. Selamat datang di SPK LINTOR.', 'success')
            next_url = request.args.get('next', '')
            # Only allow relative paths to prevent open redirects.
            parsed = urlparse(next_url)
            if next_url.startswith('/') and not next_url.startswith('//') and not parsed.netloc:
                return redirect(next_url)
            return redirect(url_for('dashboard'))
        flash('Nama pengguna atau kata sandi tidak sesuai.', 'error')
    return render_template('login.html')

@app.route('/logout', methods=['POST'])
def logout():
    session.clear()
    flash('Kamu telah keluar dari sistem.', 'success')
    return redirect(url_for('login'))

@app.route('/')
def dashboard():
    try:
        raw = load_data()
        df, selected_program, selected_kecamatan, selected_desa = get_view_data(raw)
        scored = calculate(df)
        total_target = df['Target'].sum(min_count=1)
        k1_total = df['K1'].sum(min_count=1)
        k1_rate = (k1_total / total_target * 100) if pd.notna(total_target) and total_target else 0
        stage_totals = {c: df[c].sum(min_count=1) for c in STAGES}
        priority_count = int((scored['Prioritas'] == 'Tinggi').sum())
        pie_values = [0 if pd.isna(stage_totals[c]) else int(stage_totals[c]) for c in STAGES]
        # Limit chart payload to the first 30 rows to keep rendering lightweight; table remains available in full.
        chart_df = df.head(30)
        chart_data = {
            'stages': STAGES,
            'pie': pie_values,
            'bar': {
                'labels': [str(r['Desa']) for r in chart_df.to_dict('records')],
                'kecamatan': chart_df['Kecamatan'].astype(str).tolist(),
                'series': {c: [None if pd.isna(v) else int(v) for v in chart_df[c]] for c in ['Target'] + STAGES},
            },
        }
        programs, kecamatans, desas = filter_options(raw, selected_program, selected_kecamatan)
        return render_template('dashboard.html', data=df, scored=scored, total_target=total_target,
            k1_total=k1_total, k1_rate=k1_rate, stage_totals=stage_totals, programs=programs,
            kecamatans=kecamatans, desas=desas, selected_program=selected_program,
            selected_kecamatan=selected_kecamatan, selected_desa=selected_desa,
            priority_count=priority_count, chart_data=chart_data, stages=STAGES)
    except Exception as exc:
        return render_template('error.html', error=str(exc)), 400

@app.route('/prioritas')
def priorities():
    try:
        raw = load_data(); df, selected_program, selected_kecamatan, selected_desa = get_view_data(raw)
        scored = calculate(df)
        programs, kecamatans, desas = filter_options(raw, selected_program, selected_kecamatan)
        return render_template('prioritas.html', scored=scored, programs=programs,
            kecamatans=kecamatans, desas=desas, selected_program=selected_program,
            selected_kecamatan=selected_kecamatan, selected_desa=selected_desa, stages=STAGES)
    except Exception as exc:
        return render_template('error.html', error=str(exc)), 400

@app.route('/data')
def data_page():
    try:
        raw = load_data(); df, selected_program, selected_kecamatan, selected_desa = get_view_data(raw)
        programs, kecamatans, desas = filter_options(raw, selected_program, selected_kecamatan)
        return render_template('data.html', data=df, programs=programs,
            kecamatans=kecamatans, desas=desas, selected_program=selected_program,
            selected_kecamatan=selected_kecamatan, selected_desa=selected_desa, stages=STAGES)
    except Exception as exc:
        return render_template('error.html', error=str(exc)), 400

@app.route('/healthz')
def healthz():
    # Endpoint sederhana untuk pemeriksaan status layanan hosting.
    return {'status': 'ok'}, 200


@app.route('/unggah', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename:
            flash('Pilih file Excel terlebih dahulu.', 'error'); return redirect(url_for('upload'))
        # Versi hosting hanya menerima .xlsx karena dependensi pembaca Excel
        # yang disediakan adalah openpyxl. File .xls lama perlu dikonversi dahulu.
        if not file.filename.lower().endswith('.xlsx'):
            flash('Format belum didukung. Simpan/konversi file menjadi .xlsx terlebih dahulu.', 'error'); return redirect(url_for('upload'))
        temp_path = DATA_DIR / 'upload_temp.xlsx'
        try:
            file.save(temp_path)
            test_df = load_data(temp_path)
            if len(test_df) == 0: raise ValueError('File tidak memiliki data desa.')
            temp_path.replace(UPLOAD_PATH)
            flash(f'Berhasil memuat {len(test_df)} baris data. Periksa kembali kesesuaian kolom dan angka dengan sumber resmi.', 'success')
            return redirect(url_for('dashboard'))
        except Exception as exc:
            if temp_path.exists(): temp_path.unlink()
            flash(f'File belum dapat digunakan: {exc}', 'error')
            return redirect(url_for('upload'))
    return render_template('upload.html')

@app.route('/ekspor')
def export_csv():
    try:
        df, _, _, _ = get_view_data(); scored = calculate(df)
        cols = ['Program', 'Desa', 'Kecamatan', 'Target'] + STAGES + ['Persentase K1', 'Sisa Target', 'Selisih Pengumuman-K1', 'Skor Prioritas', 'Prioritas', 'Alasan Prioritas']
        output = io.StringIO(); scored[cols].to_csv(output, index=False, sep=';')
        mem = io.BytesIO(output.getvalue().encode('utf-8-sig')); mem.seek(0)
        return send_file(mem, mimetype='text/csv', as_attachment=True, download_name='hasil_analisis_prioritas_lintor.csv')
    except Exception as exc:
        flash(f'Ekspor gagal: {exc}', 'error'); return redirect(url_for('dashboard'))

@app.route('/tentang')
def about(): return render_template('about.html')

if __name__ == '__main__':
    # Untuk pengembangan lokal. Hosting produksi menjalankan app melalui Gunicorn.
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
