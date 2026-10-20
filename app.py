from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, abort
from werkzeug.middleware.proxy_fix import ProxyFix
from database import get_db_connection, init_db
from functools import wraps
import urllib.request
import urllib.parse
import secrets
import json
import os

app = Flask(__name__)

# Hỗ trợ nhận diện đúng giao thức HTTPS phía sau Reverse Proxy của PythonAnywhere
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

# ==========================================
# CẤU HÌNH BẢO MẬT & GOOGLE OAUTH 2.0
# ==========================================
# Điền trực tiếp GOOGLE_CLIENT_ID và GOOGLE_CLIENT_SECRET tại đây hoặc cấu hình trong file WSGI trên PythonAnywhere
GOOGLE_CLIENT_ID = os.environ.get('GOOGLE_CLIENT_ID', '').strip()
GOOGLE_CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET', '').strip()

# Khai báo danh sách Quản trị viên
ADMIN_EMAILS = [
    "mlhang@hcmus.edu.vn",
    "tthan@hcmus.edu.vn"
]

# Tùy chọn: Giới hạn tên miền email được phép đăng nhập (Để danh sách rỗng [] nếu cho phép mọi tài khoản Google)
ALLOWED_EMAIL_DOMAINS = []  # Ví dụ: ["hcmus.edu.vn", "student.hcmus.edu.vn"]


def get_current_user_email():
    """Lấy email người dùng đã xác thực từ Flask Session."""
    return session.get("user_email", "")


def get_current_user_name():
    """Lấy tên hiển thị của người dùng đã xác thực từ Flask Session."""
    email = get_current_user_email()
    name = session.get("user_name", "")
    if not name and email:
        return email.split("@")[0]
    return name


@app.context_processor
def inject_user_info():
    return dict(
        ADMIN_EMAILS=ADMIN_EMAILS,
        current_user_email=get_current_user_email(),
        current_user_name=get_current_user_name()
    )


def login_required(f):
    """Yêu cầu đăng nhập Google trước khi thực hiện thao tác."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not get_current_user_email():
            if request.is_json or request.path.startswith('/submit_bulk_booking'):
                return jsonify({'success': False, 'message': 'Vui lòng đăng nhập bằng tài khoản Google trước khi gửi yêu cầu.'}), 401
            flash("Vui lòng đăng nhập bằng tài khoản Google để tiếp tục.", "warning")
            return redirect(url_for('login_page', next=request.url))
        return f(*args, **kwargs)
    return decorated_function


def admin_required(f):
    """Yêu cầu tài khoản Google thuộc danh sách ADMIN_EMAILS."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        user_email = get_current_user_email()
        if not user_email:
            flash("Vui lòng đăng nhập bằng tài khoản Google Quản trị viên.", "warning")
            return redirect(url_for('login_page', next=request.url))
        if user_email not in ADMIN_EMAILS:
            flash(f"Bạn không có quyền quản trị! (Tài khoản hiện tại: {user_email})", "danger")
            return redirect(url_for('home_page'))
        return f(*args, **kwargs)
    return decorated_function


# Thiết lập secret_key để mã hóa session
app.secret_key = os.environ.get('SECRET_KEY', 'KHTN_TKB_Secret_Key_@2026')
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # Giới hạn upload tối đa 10MB
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.jinja_env.auto_reload = True

# Khởi tạo DB khi app bắt đầu chạy
with app.app_context():
    init_db()


# ==========================================
# PHÂN HỆ XÁC THỰC GOOGLE OAUTH 2.0
# ==========================================

@app.route('/login')
def login_page():
    next_url = request.args.get('next') or url_for('home_page')
    if get_current_user_email():
        return redirect(next_url)
    oauth_configured = bool(GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET)
    return render_template('login.html', next_url=next_url, oauth_configured=oauth_configured)


@app.route('/login/google')
def google_login():
    next_url = request.args.get('next') or url_for('home_page')
    session['next_url'] = next_url

    if not GOOGLE_CLIENT_ID or not GOOGLE_CLIENT_SECRET:
        flash("Hệ thống chưa được cấu hình GOOGLE_CLIENT_ID và GOOGLE_CLIENT_SECRET.", "danger")
        return redirect(url_for('login_page', next=next_url))

    state = secrets.token_urlsafe(24)
    session['oauth_state'] = state
    redirect_uri = url_for('google_callback', _external=True)

    params = {
        'client_id': GOOGLE_CLIENT_ID,
        'redirect_uri': redirect_uri,
        'response_type': 'code',
        'scope': 'openid email profile',
        'state': state,
        'prompt': 'select_account'
    }
    auth_url = 'https://accounts.google.com/o/oauth2/v2/auth?' + urllib.parse.urlencode(params)
    return redirect(auth_url)


@app.route('/auth/google/callback')
def google_callback():
    error = request.args.get('error')
    if error:
        flash(f"Đăng nhập Google bị hủy hoặc gặp lỗi: {error}", "danger")
        return redirect(url_for('login_page'))

    state = request.args.get('state')
    if not state or state != session.pop('oauth_state', None):
        flash("Phiên xác thực không hợp lệ (CSRF State Mismatch). Vui lòng thử lại.", "danger")
        return redirect(url_for('login_page'))

    code = request.args.get('code')
    if not code:
        flash("Không nhận được mã xác thực từ Google.", "danger")
        return redirect(url_for('login_page'))

    redirect_uri = url_for('google_callback', _external=True)
    token_data = urllib.parse.urlencode({
        'code': code,
        'client_id': GOOGLE_CLIENT_ID,
        'client_secret': GOOGLE_CLIENT_SECRET,
        'redirect_uri': redirect_uri,
        'grant_type': 'authorization_code'
    }).encode('utf-8')

    try:
        token_req = urllib.request.Request(
            'https://oauth2.googleapis.com/token',
            data=token_data,
            headers={'Content-Type': 'application/x-www-form-urlencoded'}
        )
        with urllib.request.urlopen(token_req, timeout=10) as resp:
            token_json = json.loads(resp.read().decode('utf-8'))

        access_token = token_json.get('access_token')
        if not access_token:
            flash("Không lấy được Access Token từ Google.", "danger")
            return redirect(url_for('login_page'))

        userinfo_req = urllib.request.Request(
            'https://www.googleapis.com/oauth2/v3/userinfo',
            headers={'Authorization': f'Bearer {access_token}'}
        )
        with urllib.request.urlopen(userinfo_req, timeout=10) as resp:
            userinfo = json.loads(resp.read().decode('utf-8'))

        email = (userinfo.get('email') or '').strip().lower()
        name = (userinfo.get('name') or '').strip()

        if not email:
            flash("Tài khoản Google không cung cấp địa chỉ email hợp lệ.", "danger")
            return redirect(url_for('login_page'))

        if ALLOWED_EMAIL_DOMAINS:
            domain = email.split('@')[-1]
            if domain not in ALLOWED_EMAIL_DOMAINS and email not in ADMIN_EMAILS:
                flash(f"Vui lòng sử dụng email thuộc tên miền được phép ({', '.join(ALLOWED_EMAIL_DOMAINS)}).", "danger")
                return redirect(url_for('login_page'))

        session['user_email'] = email
        session['user_name'] = name or email.split('@')[0]

        next_url = session.pop('next_url', None) or url_for('home_page')
        flash(f"Xin chào {session['user_name']} ({email})!", "success")
        return redirect(next_url)

    except Exception as e:
        flash(f"Lỗi kết nối xác thực Google OAuth: {e}", "danger")
        return redirect(url_for('login_page'))


@app.route('/logout')
def logout():
    session.clear()
    flash("Đã đăng xuất khỏi hệ thống.", "success")
    return redirect(url_for('home_page'))


# ==========================================
# PHÂN HỆ USER (PUBLIC & AUTHENTICATED)
# ==========================================

import datetime
import re

ROOMS_LIST = ['5.1', '5.2', '5.3', '6.1', '6.2', '6.3', '6.5', '6.6']

def parse_study_period_date(raw_val):
    """Phân tích chuỗi Giai_Doan (VD: '28/09/2026', '2026-09-28 00:00:00') thành datetime.date."""
    if not raw_val:
        return None
    if isinstance(raw_val, datetime.datetime):
        return raw_val.date()
    if isinstance(raw_val, datetime.date):
        return raw_val
    s = str(raw_val).strip()
    if not s:
        return None
    # Thử định dạng YYYY-MM-DD (hoặc YYYY-MM-DD HH:MM:SS từ Excel)
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d'):
        try:
            part = s[:10] if fmt in ('%Y-%m-%d', '%Y/%m/%d') and len(s) >= 10 else s
            return datetime.datetime.strptime(part, fmt).date()
        except ValueError:
            pass
    # Tìm mẫu DD/MM/YYYY hoặc DD-MM-YYYY trong chuỗi bất kỳ
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', s)
    if m:
        try:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    return None

def normalize_study_period_str(raw_val):
    """Chuẩn hóa giá trị Giai_Doan về chuỗi DD/MM/YYYY nếu là ngày hợp lệ."""
    dt = parse_study_period_date(raw_val)
    if dt:
        return dt.strftime('%d/%m/%Y')
    return str(raw_val).strip() if raw_val else ''

def enrich_schedule_row(row, ref_date=None):
    """Bổ sung thông tin chu kỳ lặp (Tuần X/N, Ngày kết thúc, Trạng thái Đóng/Mở cửa) cho 1 dòng lịch học."""
    if ref_date is None:
        ref_date = datetime.date.today()
    d = dict(row)
    num_weeks = int(d.get('num_weeks') or 15)
    is_ended_early = int(d.get('is_ended_early') or 0)
    ended_early_date = d.get('ended_early_date') or ''
    raw_period = d.get('study_period') or ''

    start_dt = parse_study_period_date(raw_period)
    ref_monday = ref_date - datetime.timedelta(days=ref_date.weekday())

    d['num_weeks'] = num_weeks
    d['is_ended_early'] = is_ended_early
    d['ended_early_date'] = ended_early_date

    if start_dt:
        start_monday = start_dt - datetime.timedelta(days=start_dt.weekday())
        end_dt = start_monday + datetime.timedelta(days=num_weeks * 7 - 1)
        week_num = ((ref_monday - start_monday).days // 7) + 1

        d['start_date_str'] = start_dt.strftime('%d/%m/%Y')
        d['end_date_str'] = end_dt.strftime('%d/%m/%Y')
        d['week_num'] = week_num
        d['has_valid_date'] = True
        d['is_in_selected_week'] = (1 <= week_num <= num_weeks)

        # Tính tuần dừng nếu kết thúc sớm
        stopped_week = week_num
        if ended_early_date:
            early_dt = parse_study_period_date(ended_early_date)
            if early_dt:
                early_monday = early_dt - datetime.timedelta(days=early_dt.weekday())
                stopped_week = max(1, min(num_weeks, ((early_monday - start_monday).days // 7) + 1))
        d['stopped_week'] = max(1, min(num_weeks, stopped_week))

        if is_ended_early:
            d['door_status'] = 'ended_early'
            d['door_open'] = False
        elif week_num < 1:
            d['door_status'] = 'upcoming'
            d['door_open'] = False
        elif week_num > num_weeks:
            d['door_status'] = 'completed'
            d['door_open'] = False
        else:
            d['door_status'] = 'active'
            d['door_open'] = True
    else:
        d['start_date_str'] = raw_period if raw_period else 'Thường xuyên'
        d['end_date_str'] = ''
        d['week_num'] = 1
        d['stopped_week'] = 1
        d['has_valid_date'] = False
        d['is_in_selected_week'] = True
        if is_ended_early:
            d['door_status'] = 'ended_early'
            d['door_open'] = False
        else:
            d['door_status'] = 'active'
            d['door_open'] = True

    return d

def get_detailed_free_slots_data(day_of_week, ref_date=None):
    """Tính toán chi tiết các khoảng phòng trống dạng structured data cho ngày chỉ định."""
    conn = get_db_connection()
    schedules = conn.execute(
        'SELECT * FROM schedules WHERE day_of_week = ? AND COALESCE(is_ended_early, 0) = 0 ORDER BY start_period ASC',
        (day_of_week,)
    ).fetchall()
    conn.close()
    
    room_schedules = {room: [] for room in ROOMS_LIST}
    for row in schedules:
        enriched = enrich_schedule_row(row, ref_date=ref_date)
        # Nếu môn chưa đến tuần bắt đầu (upcoming), đã hết chu kỳ tuần lặp (completed), hoặc đã kết thúc sớm (ended_early)
        # thì giải phóng phòng trống cho ngày/tuần đang tra cứu
        if not enriched['door_open']:
            continue
        room = row['room']
        if room in room_schedules:
            room_schedules[room].append(row)
            
    def parse_slots(slots):
        res = []
        for s, e in slots:
            start_tiet = s + 1.0
            if e.is_integer():
                end_tiet = e
            else:
                end_tiet = e + 1.0
            st_str = f"{int(start_tiet)}" if start_tiet.is_integer() else f"{start_tiet}"
            ed_str = f"{int(end_tiet)}" if end_tiet.is_integer() else f"{end_tiet}"
            st_val = int(start_tiet) if start_tiet.is_integer() else start_tiet
            ed_val = int(end_tiet) if end_tiet.is_integer() else end_tiet
            res.append({
                'start_period': st_val,
                'end_period': ed_val,
                'label': f"Tiết {st_str} - {ed_str}"
            })
        return res

    morning = []
    afternoon = []
    
    for room, classes in room_schedules.items():
        # Sáng (0.0 đến 5.0)
        m_slots = []
        t_trong = 0.0
        for c in classes:
            start = float(c['start_period'])
            end = float(c['end_period'])
            if start > 5.0: continue
            c_start_time = start - 1.0
            c_end_time = end if end.is_integer() else float(int(end)) - 0.5
            if c_start_time > t_trong:
                if (c_start_time - t_trong) >= 2.5:
                    m_slots.append((t_trong, c_start_time))
            t_trong = max(t_trong, c_end_time)
        if 5.0 - t_trong >= 2.5:
            m_slots.append((t_trong, 5.0))
        for item in parse_slots(m_slots):
            morning.append({'room': room, **item})

        # Chiều (5.0 đến 10.0)
        a_slots = []
        t_trong = 5.0
        for c in classes:
            start = float(c['start_period'])
            end = float(c['end_period'])
            if end <= 5.0: continue
            c_start_time = max(start - 1.0, 5.0)
            c_end_time = end if end.is_integer() else float(int(end)) - 0.5
            if c_start_time > t_trong:
                if (c_start_time - t_trong) >= 2.5:
                    a_slots.append((t_trong, c_start_time))
            t_trong = max(t_trong, c_end_time)
        if 10.0 - t_trong >= 2.5:
            a_slots.append((t_trong, 10.0))
        for item in parse_slots(a_slots):
            afternoon.append({'room': room, **item})
            
    return morning, afternoon

def get_free_slots_for_day(day_of_week, ref_date=None):
    """Giữ nguyên tương thích ngược cho các nơi sử dụng định dạng dict, có hỗ trợ ref_date."""
    morning, afternoon = get_detailed_free_slots_data(day_of_week, ref_date=ref_date)
    morning_free = {}
    for item in morning:
        morning_free.setdefault(item['room'], []).append(item['label'])
    afternoon_free = {}
    for item in afternoon:
        afternoon_free.setdefault(item['room'], []).append(item['label'])
    return morning_free, afternoon_free

@app.route('/')
def index():
    """Chuyển hướng đến trang chủ"""
    return redirect(url_for('home_page'))

@app.route('/bao-cao-su-co')
@login_required
def report_page():
    return render_template('report.html')

@app.route('/submit_issue', methods=['POST'])
@login_required
def submit_issue():
    reporter_name = request.form.get('reporter_name') or get_current_user_name()
    department = request.form.get('department')
    room = request.form.get('room')
    machine_number = request.form.get('machine_number')
    issue_category = request.form.get('issue_category')
    description = request.form.get('description')
    
    conn = get_db_connection()
    try:
        user_email = get_current_user_email()
        conn.execute('''
        INSERT INTO issue_reports (reporter_name, department, room, machine_number, issue_category, description, email)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (reporter_name, department, room, machine_number, issue_category, description, user_email))
        conn.commit()
        flash("Gửi báo cáo sự cố thành công! Bộ phận QTTB sẽ nhanh chóng xử lý.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Có lỗi xảy ra: {e}", "error")
    finally:
        conn.close()
        
    return redirect(url_for('report_page'))

@app.route('/dang-ky-muon-phong')
@login_required
def booking_page():
    """Hiển thị trang chứa Form đăng ký mượn phòng."""
    selected_day = request.args.get('filter_day', 2, type=int)
    raw_date = request.args.get('filter_date', '').strip()
    ref_date = parse_study_period_date(raw_date) if raw_date else datetime.date.today()
    
    morning_free, afternoon_free = get_free_slots_for_day(selected_day, ref_date=ref_date)
    return render_template('booking.html', 
                           selected_day=selected_day, 
                           morning_free=morning_free, 
                           afternoon_free=afternoon_free)

@app.route('/api/free-slots')
def api_free_slots():
    """API trả về danh sách chi tiết các ca phòng trống theo ngày và mốc tuần (cho sidebar & popup chọn nhanh)."""
    day = request.args.get('day', 2, type=int)
    if day < 2 or day > 7:
        day = 2
    raw_date = request.args.get('date', '').strip()
    ref_date = parse_study_period_date(raw_date) if raw_date else None
    if ref_date is None:
        ref_date = datetime.date.today()

    # Tính đúng ngày của Thứ (day) trong tuần chứa ref_date
    monday_of_ref = ref_date - datetime.timedelta(days=ref_date.weekday())
    target_date = monday_of_ref + datetime.timedelta(days=day - 2)
    saturday_of_ref = monday_of_ref + datetime.timedelta(days=5)

    morning, afternoon = get_detailed_free_slots_data(day, ref_date=target_date)
    return jsonify({
        'day': day,
        'target_date': target_date.strftime('%d/%m/%Y'),
        'target_date_iso': target_date.strftime('%Y-%m-%d'),
        'week_range': f"{monday_of_ref.strftime('%d/%m/%Y')} – {saturday_of_ref.strftime('%d/%m/%Y')}",
        'morning': morning,
        'afternoon': afternoon
    })

@app.route('/trang-chu')
def home_page():
    """Hiển thị trang chủ phòng máy."""
    return render_template('home.html')

@app.route('/booking', methods=['POST'])
@login_required
def submit_booking():
    """Xử lý data khi User gửi form mượn phòng."""
    requester_name = request.form.get('requester_name') or get_current_user_name()
    room = request.form.get('room')
    day_of_week = int(request.form.get('day_of_week')) if request.form.get('day_of_week') else None
    start_period = float(request.form.get('start_period')) if request.form.get('start_period') else None
    end_period = float(request.form.get('end_period')) if request.form.get('end_period') else None
    purpose = request.form.get('purpose')
    subject_code = request.form.get('subject_code')
    class_code = request.form.get('class_code')
    subject_name = request.form.get('subject_name')
    department = request.form.get('department')
    study_period = request.form.get('study_period') or request.form.get('study_date') or ''
    software_required = request.form.get('software_required')
    
    # Bug #9: Validate tiết bắt đầu và kết thúc
    if start_period is not None and end_period is not None:
        if start_period >= end_period:
            flash("Tiết bắt đầu phải nhỏ hơn tiết kết thúc!", "danger")
            return redirect(url_for('booking_page'))
        if start_period < 1 or end_period > 10:
            flash("Tiết phải nằm trong khoảng từ 1 đến 10!", "danger")
            return redirect(url_for('booking_page'))

    conn = get_db_connection()
    user_email = get_current_user_email()
    conn.execute('''
        INSERT INTO booking_requests (requester_name, room, day_of_week, start_period, end_period, purpose, subject_code, class_code, subject_name, department, study_period, software_required, status, email, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pending', ?, datetime('now', 'localtime'))
    ''', (requester_name, room, day_of_week, start_period, end_period, purpose, subject_code, class_code, subject_name, department, study_period, software_required, user_email))
    conn.commit()
    conn.close()
    
    flash("Gửi yêu cầu mượn phòng thành công! Vui lòng chờ Bộ phận Quản trị thiết bị duyệt.", "success")
    return redirect(url_for('booking_page'))

def is_pending_request_active_on_date(pending_row, ref_date):
    """Kiểm tra yêu cầu đang chờ duyệt (Pending) có trùng giai đoạn tuần với ngày đăng ký mới hay không."""
    if not ref_date:
        return True
    p_dt = parse_study_period_date(pending_row['study_period'] if 'study_period' in pending_row.keys() else '')
    if not p_dt:
        return True
    p_monday = p_dt - datetime.timedelta(days=p_dt.weekday())
    ref_monday = ref_date - datetime.timedelta(days=ref_date.weekday())
    week_diff = ((ref_monday - p_monday).days // 7) + 1
    return 1 <= week_diff <= 15

@app.route('/submit_bulk_booking', methods=['POST'])
@login_required
def submit_bulk_booking():
    """Xử lý data khi User gửi giỏ hàng đăng ký (nhiều môn)."""
    data = request.get_json()
    if not data or not isinstance(data, list):
        return jsonify({'success': False, 'message': 'Dữ liệu không hợp lệ'}), 400
    
    conn = get_db_connection()
    
    # 1. KIỂM TRA TRÙNG LỊCH VÀ TỰ ĐỘNG SẮP XẾP PHÒNG THEO ĐÚNG NGÀY/GIAI ĐOẠN CỦA TỪNG MÔN
    schedules_db = conn.execute(
        'SELECT room, day_of_week, start_period, end_period, study_period, num_weeks, is_ended_early, ended_early_date FROM schedules WHERE COALESCE(is_ended_early, 0) = 0'
    ).fetchall()
    pending_db = conn.execute(
        "SELECT room, day_of_week, start_period, end_period, study_period FROM booking_requests WHERE status = 'Pending'"
    ).fetchall()
    cart_assigned = []
    
    for item in data:
        room = item.get('room')
        try:
            dow = int(item.get('day_of_week'))
            req_start = float(item.get('start_period'))
            req_end = float(item.get('end_period'))
        except (ValueError, TypeError):
            conn.close()
            return jsonify({'success': False, 'message': 'Dữ liệu thời gian không hợp lệ'}), 400

        raw_item_date = item.get('study_period') or item.get('study_date') or ''
        item_ref_date = parse_study_period_date(raw_item_date) or datetime.date.today()

        active_for_item = [
            dict(s) for s in schedules_db
            if enrich_schedule_row(s, ref_date=item_ref_date)['door_open']
        ]
        pending_for_item = [
            dict(p) for p in pending_db
            if is_pending_request_active_on_date(p, item_ref_date)
        ]
        cart_for_item = [
            c for c in cart_assigned
            if is_pending_request_active_on_date(c, item_ref_date)
        ]
        schedules = active_for_item + pending_for_item + cart_for_item
            
        subject_name = item.get('subject_name', 'Môn học')
        rooms_to_check = [room] if room != 'Any' else ROOMS_LIST
        is_free_in_some_room = False
        assigned_room = None
        
        for r in rooms_to_check:
            overlap = False
            for s in schedules:
                if s['room'] == r and s['day_of_week'] == dow:
                    s_start = float(s['start_period'])
                    s_end = float(s['end_period'])
                    if req_start < s_end and req_end > s_start:
                        overlap = True
                        break
            if not overlap:
                is_free_in_some_room = True
                assigned_room = r
                break
                
        if not is_free_in_some_room:
            conn.close()
            error_msg = f"Môn '{subject_name}': Ca này đã được sử dụng (kể cả khi tìm phòng tự động), vui lòng chọn lịch khác."
            return jsonify({'success': False, 'message': error_msg}), 200
            
        # Gán phòng tự động nếu chọn 'Any'
        if room == 'Any':
            item['room'] = assigned_room
            
        # Thêm ca này vào danh sách cart_assigned để môn tiếp theo trong giỏ hàng không được xếp đè vào đây
        cart_assigned.append({
            'room': item['room'],
            'day_of_week': dow,
            'start_period': req_start,
            'end_period': req_end,
            'study_period': raw_item_date
        })

    # 2. TIẾN HÀNH LƯU VÀO DATABASE
    user_email = get_current_user_email()
    default_name = get_current_user_name()
    try:
        for item in data:
            conn.execute('''
                INSERT INTO booking_requests (
                    requester_name, room, day_of_week, start_period, end_period, purpose, 
                    subject_code, class_code, subject_name, department, study_period, software_required, 
                    status, email, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pending', ?, datetime('now', 'localtime'))
            ''', (
                item.get('requester_name') or default_name, 
                item.get('room'), 
                int(item.get('day_of_week')),
                float(item.get('start_period')), 
                float(item.get('end_period')), 
                item.get('purpose'), 
                item.get('subject_code'), 
                item.get('class_code'), 
                item.get('subject_name'), 
                item.get('department'), 
                item.get('study_period'), 
                item.get('software_required'),
                user_email
            ))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()
    
    flash("Gửi danh sách yêu cầu mượn phòng thành công! Các phòng Tùy ý (nếu có) đã được hệ thống tự động xếp vào phòng trống.", "success")
    return jsonify({'success': True, 'redirect_url': url_for('booking_page')})

@app.route('/check_conflict', methods=['POST'])
def check_conflict():
    data = request.get_json()
    room = data.get('room')
    try:
        dow = int(data.get('day_of_week'))
        req_start = float(data.get('start_period'))
        req_end = float(data.get('end_period'))
    except (ValueError, TypeError):
        return jsonify({'is_free': False})

    raw_req_date = data.get('study_period') or data.get('study_date') or ''
    req_ref_date = parse_study_period_date(raw_req_date) or datetime.date.today()
        
    conn = get_db_connection()
    # Kiểm tra cả lịch đã duyệt (đang còn hiệu lực trong tuần của req_ref_date) và yêu cầu đang chờ duyệt
    schedules_confirmed = conn.execute(
        'SELECT room, day_of_week, start_period, end_period, study_period, num_weeks, is_ended_early, ended_early_date FROM schedules WHERE day_of_week = ? AND COALESCE(is_ended_early, 0) = 0', (dow,)
    ).fetchall()
    active_confirmed = [
        s for s in schedules_confirmed
        if enrich_schedule_row(s, ref_date=req_ref_date)['door_open']
    ]
    schedules_pending = conn.execute(
        'SELECT room, day_of_week, start_period, end_period, study_period FROM booking_requests WHERE day_of_week = ? AND status = ?',
        (dow, 'Pending')
    ).fetchall()
    active_pending = [
        p for p in schedules_pending
        if is_pending_request_active_on_date(p, req_ref_date)
    ]
    schedules = list(active_confirmed) + list(active_pending)
    conn.close()
    
    rooms_to_check = [room] if room != 'Any' else ROOMS_LIST
    
    is_free_in_some_room = False
    assigned_room = None
    for r in rooms_to_check:
        overlap = False
        for s in schedules:
            if s['room'] == r:
                s_start = float(s['start_period'])
                s_end = float(s['end_period'])
                if req_start < s_end and req_end > s_start:
                    overlap = True
                    break
        if not overlap:
            is_free_in_some_room = True
            assigned_room = r
            break
            
    return jsonify({'is_free': is_free_in_some_room, 'assigned_room': assigned_room})

@app.route('/thong-bao-nghi')
@login_required
def cancellation_page():
    """Hiển thị trang thông báo nghỉ gấp."""
    return render_template('cancellation.html')

@app.route('/submit-cancellation', methods=['POST'])
@login_required
def submit_cancellation():
    """Xử lý data khi giáo viên báo nghỉ gấp."""
    teacher_name = request.form.get('teacher_name') or get_current_user_name()
    room = request.form.get('room')
    cancel_date = request.form.get('cancel_date')
    period = request.form.get('period')
    reason = request.form.get('reason')
    
    conn = get_db_connection()
    user_email = get_current_user_email()
    conn.execute('''
        INSERT INTO cancellations (teacher_name, room, cancel_date, period, reason, email)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (teacher_name, room, cancel_date, period, reason, user_email))
    conn.commit()
    conn.close()
    
    flash("Gửi thông báo nghỉ thành công! Phòng QTTB sẽ tiến hành đóng cửa phòng để bảo vệ tài sản.", "success")
    return redirect(url_for('cancellation_page'))

# ==========================================
# PHÂN HỆ ADMIN (PROTECTED)
# ==========================================

@app.route('/admin/login')
@app.route('/admin')
def redirect_to_admin():
    if not get_current_user_email():
        return redirect(url_for('login_page', next=url_for('admin_panel')))
    return redirect(url_for('admin_panel'))

@app.route('/admin/panel')
@admin_required
def admin_panel():
    """Trang quản trị (Dashboard): Chỉ Admin mới vào được."""
    conn = get_db_connection()
    today = datetime.date.today()

    # 1. Lấy danh sách yêu cầu mượn phòng đang chờ duyệt (Pending)
    pending_requests = conn.execute("SELECT * FROM booking_requests WHERE status = 'Pending'").fetchall()

    # 2. Lấy danh sách lịch học hiện tại (để hiển thị trong bảng quản lý) và bổ sung thông tin chu kỳ tuần
    raw_schedules = conn.execute("SELECT * FROM schedules ORDER BY room ASC, day_of_week ASC, start_period ASC").fetchall()
    all_enriched_schedules = [enrich_schedule_row(r, ref_date=today) for r in raw_schedules]

    schedule_filter = request.args.get('schedule_filter', 'all')
    active_list = [s for s in all_enriched_schedules if s['door_open']]
    locked_list = [s for s in all_enriched_schedules if not s['door_open']]
    schedule_counts = {
        'all': len(all_enriched_schedules),
        'active': len(active_list),
        'locked': len(locked_list)
    }
    if schedule_filter == 'active':
        current_schedules = active_list
    elif schedule_filter == 'locked':
        current_schedules = locked_list
    else:
        current_schedules = all_enriched_schedules

    # Logic tìm phòng trống theo Thứ & Ngày/Tuần (hỗ trợ tra cứu Học kỳ 2)
    filter_day = request.args.get('filter_day', 2, type=int)
    raw_free_date = request.args.get('filter_date', '').strip()
    free_ref_date = parse_study_period_date(raw_free_date) if raw_free_date else today
    if not free_ref_date:
        free_ref_date = today
    free_monday = free_ref_date - datetime.timedelta(days=free_ref_date.weekday())
    free_target_date = free_monday + datetime.timedelta(days=filter_day - 2)
    filter_date_str = free_target_date.strftime('%Y-%m-%d')
    free_target_date_display = free_target_date.strftime('%d/%m/%Y')
    admin_morning_free, admin_afternoon_free = get_free_slots_for_day(filter_day, ref_date=free_target_date)

    # 3. Lấy dữ liệu hiển thị dạng Lưới theo tuần (Chế độ Google Calendar)
    rooms = ['5.1', '5.2', '5.3', '6.1', '6.2', '6.3', '6.5', '6.6']
    selected_room = request.args.get('room', '5.1')

    raw_tt_date = request.args.get('timetable_date', '').strip()
    selected_date = parse_study_period_date(raw_tt_date) if raw_tt_date else today
    if not selected_date:
        selected_date = today

    monday_date = selected_date - datetime.timedelta(days=selected_date.weekday())
    saturday_date = monday_date + datetime.timedelta(days=5)
    prev_week_date_str = (monday_date - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
    next_week_date_str = (monday_date + datetime.timedelta(days=7)).strftime('%Y-%m-%d')
    timetable_date_str = selected_date.strftime('%Y-%m-%d')
    timetable_date_display = selected_date.strftime('%d/%m/%Y')
    today_date_str = today.strftime('%Y-%m-%d')
    week_range_label = f"Thứ Hai {monday_date.strftime('%d/%m')} – Thứ Bảy {saturday_date.strftime('%d/%m/%Y')}"

    day_headers = {}
    for d in range(2, 8):
        d_date = monday_date + datetime.timedelta(days=d - 2)
        day_headers[d] = {
            'date_str': d_date.strftime('%d/%m/%Y'),
            'short_date': d_date.strftime('%d/%m'),
            'is_today': (d_date == today)
        }

    schedules_grid = conn.execute('SELECT * FROM schedules WHERE room = ? ORDER BY start_period ASC', (selected_room,)).fetchall()
    week_data = {day: [] for day in range(2, 8)}

    for row in schedules_grid:
        day = row['day_of_week']
        if day in week_data:
            enriched = enrich_schedule_row(row, ref_date=monday_date)
            # Chỉ hiển thị trên Thời khóa biểu tuần này nếu học phần nằm trong chu kỳ lặp 1..num_weeks của tuần đang chọn
            if not enriched['is_in_selected_week']:
                continue

            start = float(enriched['start_period'])
            end = float(enriched['end_period'])

            start_time = start - 1.0
            if end.is_integer():
                end_time = end
            else:
                end_time = float(int(end)) - 0.5

            top_pct = start_time / 10 * 100
            height_pct = (end_time - start_time) / 10 * 100

            week_data[day].append({
                'data': enriched,
                'top': top_pct,
                'height': height_pct
            })

    cancellations = conn.execute("SELECT * FROM cancellations ORDER BY created_at DESC").fetchall()
    issues = conn.execute("SELECT * FROM issue_reports ORDER BY created_at DESC").fetchall()
    conn.close()
    return render_template(
        'admin_panel.html',
        pending_requests=pending_requests,
        schedules=current_schedules,
        schedule_counts=schedule_counts,
        schedule_filter=schedule_filter,
        rooms=rooms,
        selected_room=selected_room,
        week_data=week_data,
        day_headers=day_headers,
        timetable_date_str=timetable_date_str,
        timetable_date_display=timetable_date_display,
        today_date_str=today_date_str,
        prev_week_date_str=prev_week_date_str,
        next_week_date_str=next_week_date_str,
        week_range_label=week_range_label,
        cancellations=cancellations,
        filter_day=filter_day,
        filter_date_str=filter_date_str,
        free_target_date_display=free_target_date_display,
        morning_free=admin_morning_free,
        afternoon_free=admin_afternoon_free,
        issues=issues
    )

@app.route('/admin/approve-cancellation/<int:cancellation_id>', methods=['POST'])
@admin_required
def approve_cancellation(cancellation_id):
    """Xử lý Duyệt thông báo nghỉ gấp."""
    conn = get_db_connection()
    conn.execute("UPDATE cancellations SET status = 'Approved' WHERE id = ?", (cancellation_id,))
    conn.commit()
    conn.close()
    
    flash("Đã xác nhận thông báo nghỉ. Đã lưu lịch sử để hỗ trợ việc sắp xếp lịch học bù.", "success")
    return redirect(url_for('admin_panel', tab='cancellations'))

@app.route('/admin/approve/<int:request_id>', methods=['POST'])
@admin_required
def approve_request(request_id):
    """Xử lý Duyệt yêu cầu mượn phòng."""
    conn = get_db_connection()
    req = conn.execute("SELECT * FROM booking_requests WHERE id = ?", (request_id,)).fetchone()
    
    if req:
        class_code = req['class_code'] if req['class_code'] else 'MƯỢN PHÒNG'
        subject_name = req['subject_name'] if req['subject_name'] else req['purpose']
        department = req['department'] if req['department'] else req['requester_name']
        study_period = normalize_study_period_str(req['study_period']) if req['study_period'] else datetime.date.today().strftime('%d/%m/%Y')

        conn.execute('''
            INSERT INTO schedules (room, day_of_week, start_period, end_period, class_code, subject_name, department, study_period, num_weeks, is_ended_early)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, 0)
        ''', (req['room'], req['day_of_week'], req['start_period'], req['end_period'], class_code, subject_name, department, study_period))
        
        conn.execute("UPDATE booking_requests SET status = 'Approved' WHERE id = ?", (request_id,))
        conn.commit()
        flash(f"Đã duyệt yêu cầu của {req['requester_name']} vào phòng {req['room']}.", "success")
        
    conn.close()
    return redirect(url_for('admin_panel', tab='requests'))

@app.route('/admin/reject/<int:request_id>', methods=['POST'])
@admin_required
def reject_request(request_id):
    """Xử lý Từ chối yêu cầu mượn phòng."""
    admin_note = request.form.get('admin_note', '')
    conn = get_db_connection()
    conn.execute("UPDATE booking_requests SET status = 'Rejected', admin_note = ? WHERE id = ?", (admin_note, request_id))
    conn.commit()
    conn.close()
    
    flash("Đã từ chối yêu cầu mượn phòng.", "warning")
    return redirect(url_for('admin_panel', tab='requests'))

@app.route('/admin/schedule/<int:schedule_id>/toggle_early_end', methods=['POST'])
@admin_required
def toggle_early_end(schedule_id):
    """Đánh dấu Kết thúc môn sớm (Khóa cửa phòng & Giải phóng phòng trống) hoặc Khôi phục lại lịch."""
    return_tab = request.form.get('return_tab', 'schedules')
    room = request.form.get('room', '5.1')
    timetable_date = request.form.get('timetable_date', '')
    timetable_day = request.form.get('timetable_day', 'all')
    schedule_filter = request.form.get('schedule_filter', 'all')

    conn = get_db_connection()
    row = conn.execute("SELECT * FROM schedules WHERE id = ?", (schedule_id,)).fetchone()
    if row:
        current_ended = int(row['is_ended_early'] or 0)
        if current_ended == 0:
            today_str = datetime.date.today().strftime('%d/%m/%Y')
            conn.execute(
                "UPDATE schedules SET is_ended_early = 1, ended_early_date = ? WHERE id = ?",
                (today_str, schedule_id)
            )
            conn.commit()
            flash(f"🔒 Đã đánh dấu kết thúc sớm môn '{row['subject_name']}' (P.{row['room']}). Ca máy này đã chuyển sang Khóa cửa và giải phóng phòng trống!", "warning")
        else:
            conn.execute(
                "UPDATE schedules SET is_ended_early = 0, ended_early_date = NULL WHERE id = ?",
                (schedule_id,)
            )
            conn.commit()
            flash(f"🔓 Đã khôi phục lại lịch học môn '{row['subject_name']}' (P.{row['room']}) về trạng thái Mở cửa bình thường.", "success")
    conn.close()

    if return_tab == 'timetable':
        return redirect(url_for('admin_panel', tab='timetable', room=room, timetable_date=timetable_date, timetable_day=timetable_day))
    return redirect(url_for('admin_panel', tab='schedules', schedule_filter=schedule_filter))

@app.route('/admin/schedule/<int:schedule_id>/update_weeks', methods=['POST'])
@admin_required
def update_schedule_weeks(schedule_id):
    """Cập nhật số tuần lặp lại (num_weeks) cho một môn học cụ thể."""
    return_tab = request.form.get('return_tab', 'schedules')
    room = request.form.get('room', '5.1')
    timetable_date = request.form.get('timetable_date', '')
    timetable_day = request.form.get('timetable_day', 'all')
    schedule_filter = request.form.get('schedule_filter', 'all')
    new_weeks = request.form.get('num_weeks', 15, type=int)
    new_weeks = max(1, min(52, new_weeks))

    conn = get_db_connection()
    row = conn.execute("SELECT subject_name, room FROM schedules WHERE id = ?", (schedule_id,)).fetchone()
    if row:
        conn.execute("UPDATE schedules SET num_weeks = ? WHERE id = ?", (new_weeks, schedule_id))
        conn.commit()
        flash(f"Đã cập nhật chu kỳ lặp của môn '{row['subject_name']}' (P.{row['room']}) thành {new_weeks} tuần.", "success")
    conn.close()

    if return_tab == 'timetable':
        return redirect(url_for('admin_panel', tab='timetable', room=room, timetable_date=timetable_date, timetable_day=timetable_day))
    return redirect(url_for('admin_panel', tab='schedules', schedule_filter=schedule_filter))

@app.route('/admin/delete_schedule/<int:schedule_id>', methods=['POST'])
@admin_required
def delete_schedule(schedule_id):
    """Xóa một lịch học khỏi cơ sở dữ liệu."""
    conn = get_db_connection()
    conn.execute("DELETE FROM schedules WHERE id = ?", (schedule_id,))
    conn.commit()
    conn.close()
    
    flash("Đã xóa lịch học thành công.", "success")
    return redirect(url_for('admin_panel', tab='schedules'))

@app.route('/admin/delete_all_schedules', methods=['POST'])
@admin_required
def delete_all_schedules():
    """Xóa tất cả lịch học khỏi cơ sở dữ liệu."""
    conn = get_db_connection()
    conn.execute("DELETE FROM schedules")
    conn.commit()
    conn.close()
    
    flash("Đã xóa toàn bộ lịch học thành công.", "success")
    return redirect(url_for('admin_panel', tab='schedules'))

@app.route('/admin/add_schedule', methods=['POST'])
@admin_required
def add_schedule():
    """Xử lý form Thêm lịch mới thủ công của Admin."""
    room = request.form.get('room')
    day_of_week = int(request.form.get('day_of_week'))
    start_period = float(request.form.get('start_period'))
    end_period = float(request.form.get('end_period'))
    subject_name = request.form.get('subject_name')
    class_code = request.form.get('class_code')
    department = request.form.get('department')
    raw_study_period = request.form.get('study_period', '')
    study_period = normalize_study_period_str(raw_study_period)
    num_weeks = request.form.get('num_weeks', 15, type=int)
    num_weeks = max(1, min(52, num_weeks))
    
    if start_period >= end_period:
        flash("Tiết bắt đầu phải nhỏ hơn tiết kết thúc!", "danger")
        return redirect(url_for('admin_panel', tab='schedules'))
    if start_period < 1 or end_period > 10:
        flash("Tiết phải nằm trong khoảng từ 1 đến 10!", "danger")
        return redirect(url_for('admin_panel', tab='schedules'))

    conn = get_db_connection()
    conn.execute('''
        INSERT INTO schedules (room, day_of_week, start_period, end_period, class_code, subject_name, department, study_period, num_weeks, is_ended_early)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
    ''', (room, day_of_week, start_period, end_period, class_code, subject_name, department, study_period, num_weeks))
    
    conn.commit()
    conn.close()
    
    flash(f"Đã thêm lịch môn {subject_name} vào phòng {room} (lặp {num_weeks} tuần) thành công.", "success")
    return redirect(url_for('admin_panel', tab='schedules'))

import pandas as pd
from werkzeug.utils import secure_filename

@app.route('/admin/import_schedule', methods=['POST'])
@admin_required
def import_schedule():
    """Xử lý upload file Excel/CSV và import hàng loạt vào Database."""
    if 'file' not in request.files:
        flash("Không tìm thấy file tải lên.", "danger")
        return redirect(url_for('admin_panel', tab='schedules'))
        
    file = request.files['file']
    default_num_weeks = request.form.get('default_num_weeks', 15, type=int)
    default_num_weeks = max(1, min(52, default_num_weeks))
    
    if file.filename == '':
        flash("Chưa chọn file nào.", "warning")
        return redirect(url_for('admin_panel', tab='schedules'))

    if file:
        try:
            filename = secure_filename(file.filename)
            if filename.endswith('.csv'):
                df = pd.read_csv(file)
            elif filename.endswith(('.xls', '.xlsx')):
                df = pd.read_excel(file)
            else:
                flash("Định dạng file không được hỗ trợ. Chỉ nhận CSV hoặc Excel.", "danger")
                return redirect(url_for('admin_panel', tab='schedules'))

            df = df.fillna('')
            
            conn = get_db_connection()
            count = 0
            
            for index, row in df.iterrows():
                room = str(row.get('Phong', '')).strip()
                day_of_week = row.get('Thu', '')
                start_period = row.get('Tiet_BD', '')
                end_period = row.get('Tiet_KT', '')
                class_code = str(row.get('Ma_Lop', '')).strip()
                subject_name = str(row.get('Ten_Mon', '')).strip()
                department = str(row.get('Khoa', '')).strip()
                raw_period = row.get('Giai_Doan', '')
                study_period = normalize_study_period_str(raw_period)
                
                if not room or not day_of_week or start_period == '' or end_period == '':
                    continue

                try:
                    day_val = int(day_of_week)
                    start_val = float(start_period)
                    end_val = float(end_period)
                except ValueError:
                    continue

                # Kiểm tra trùng lịch (cùng phòng, thứ, tiết và cùng giai đoạn bắt đầu)
                existing = conn.execute('''
                    SELECT id FROM schedules
                    WHERE room = ? AND day_of_week = ? AND start_period = ? AND end_period = ? AND COALESCE(study_period, '') = ?
                ''', (room, day_val, start_val, end_val, study_period)).fetchone()
                if existing:
                    continue

                conn.execute('''
                    INSERT INTO schedules (room, day_of_week, start_period, end_period, class_code, subject_name, department, study_period, num_weeks, is_ended_early)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
                ''', (room, day_val, start_val, end_val, class_code, subject_name, department, study_period, default_num_weeks))
                count += 1
                
            conn.commit()
            conn.close()
            
            flash(f"Đã import thành công {count} dòng lịch học (chu kỳ lặp {default_num_weeks} tuần) vào hệ thống!", "success")
            
        except Exception as e:
            flash(f"Có lỗi xảy ra khi đọc file: {str(e)}", "danger")

    return redirect(url_for('admin_panel', tab='schedules'))

@app.route('/update_issue_status/<int:issue_id>', methods=['POST'])
@admin_required
def update_issue_status(issue_id):
    new_status = request.form.get('status')
    if new_status in ['Pending', 'Processing', 'Resolved', 'Đã tiếp nhận', 'Đã khắc phục']:
        if new_status == 'Đã tiếp nhận':
            new_status = 'Processing'
        elif new_status == 'Đã khắc phục':
            new_status = 'Resolved'
        conn = get_db_connection()
        conn.execute('UPDATE issue_reports SET status = ? WHERE id = ?', (new_status, issue_id))
        conn.commit()
        conn.close()
        status_label = "Đã tiếp nhận" if new_status == 'Processing' else ("Đã khắc phục" if new_status == 'Resolved' else "Chờ xử lý")
        flash(f"Đã cập nhật trạng thái sự cố: {status_label}", "success")
    
    return redirect(url_for('admin_panel', tab='issues'))


@app.route('/my-dashboard')
@login_required
def my_dashboard():
    """Trang cá nhân của User để xem lịch sử yêu cầu."""
    user_email = get_current_user_email()
        
    conn = get_db_connection()
    # Lấy booking requests
    my_bookings = conn.execute("SELECT * FROM booking_requests WHERE email = ? ORDER BY created_at DESC", (user_email,)).fetchall()
    # Lấy issue reports
    my_issues = conn.execute("SELECT * FROM issue_reports WHERE email = ? ORDER BY created_at DESC", (user_email,)).fetchall()
    # Lấy báo nghỉ
    my_cancellations = conn.execute("SELECT * FROM cancellations WHERE email = ? ORDER BY created_at DESC", (user_email,)).fetchall()
    conn.close()
    
    return render_template('my_dashboard.html', bookings=my_bookings, issues=my_issues, cancellations=my_cancellations)

if __name__ == '__main__':
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    app.run(debug=debug_mode, host='127.0.0.1', port=5000)
