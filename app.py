from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_file, Response
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from database import get_db_connection, init_db
from report_exporter import (
    ROOMS_LIST,
    parse_report_date as parse_study_period_date,
    clean_and_group_departments,
    get_schedule_ca_slots,
    build_semester_report_excel,
    summarize_available_semesters,
    filter_schedules_for_report
)
from functools import wraps
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
import html as html_lib
import pandas as pd
import urllib.request
import urllib.parse
import subprocess
import threading
import platform
import datetime
import smtplib
import secrets
import time
import json
import os
import io
import re

app = Flask(__name__)

# Hỗ trợ nhận diện đúng giao thức HTTPS phía sau Reverse Proxy của PythonAnywhere
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

# ==========================================
# CẤU HÌNH BẢO MẬT & GOOGLE OAUTH 2.0
# ==========================================
# Điền trực tiếp GOOGLE_CLIENT_ID và GOOGLE_CLIENT_SECRET tại đây hoặc cấu hình trong file WSGI trên PythonAnywhere
GOOGLE_CLIENT_ID = os.environ.get('GOOGLE_CLIENT_ID', '').strip()
GOOGLE_CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET', '').strip()

# Tự động đọc từ file "gg client id code.txt" hoặc "gg clent id code.txt" khi chạy thử ở máy cá nhân
_possible_secret_files = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gg client id code.txt'),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gg clent id code.txt'),
]
for _sec_file in _possible_secret_files:
    if (not GOOGLE_CLIENT_ID or not GOOGLE_CLIENT_SECRET) and os.path.exists(_sec_file):
        try:
            with open(_sec_file, 'r', encoding='utf-8') as _f:
                for _line in _f:
                    if _line.lower().startswith('client id:'):
                        GOOGLE_CLIENT_ID = _line.split(':', 1)[1].strip()
                    elif _line.lower().startswith('client secret:'):
                        GOOGLE_CLIENT_SECRET = _line.split(':', 1)[1].strip()
        except Exception:
            pass
    if GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET:
        break

# Khai báo danh sách Quản trị viên
ADMIN_EMAILS = [
    "mlhang@hcmus.edu.vn",
    "tthan@hcmus.edu.vn",
    "nvtinh@hcmus.edu.vn"
]

# Tùy chọn: Giới hạn tên miền email được phép đăng nhập (Để danh sách rỗng [] nếu cho phép mọi tài khoản Google)
ALLOWED_EMAIL_DOMAINS = ["hcmus.edu.vn", "gmail.com"]


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


DEFAULT_TEACHER_PASSCODE = "QttB@2026"
RATE_LIMIT_MAX_REQUESTS = 2
RATE_LIMIT_WINDOW_SECONDS = 60


def is_user_blocked(email=None, conn=None):
    """Kiểm tra tài khoản email có nằm trong danh sách bị khóa vĩnh viễn (Block User) hay không."""
    target_email = (email if email is not None else get_current_user_email() or '').strip().lower()
    if not target_email:
        return False
    if target_email in [e.lower() for e in ADMIN_EMAILS]:
        return False

    close_after = False
    if conn is None:
        conn = get_db_connection()
        close_after = True
    row = conn.execute("SELECT id FROM blocked_users WHERE LOWER(email) = LOWER(?)", (target_email,)).fetchone()
    if close_after:
        conn.close()
    return bool(row)


def check_submission_rate_limit(email=None, conn=None):
    """Giới hạn tốc độ gửi (Rate Limiting): Mỗi tài khoản chỉ được gửi tối đa 2 đơn trong vòng 1 phút (60 giây).
    Trả về (allowed: bool, retry_after_seconds: int).
    """
    target_email = (email if email is not None else get_current_user_email() or '').strip().lower()
    if not target_email:
        return False, RATE_LIMIT_WINDOW_SECONDS

    now_ts = time.time()
    cutoff_ts = now_ts - RATE_LIMIT_WINDOW_SECONDS

    close_after = False
    if conn is None:
        conn = get_db_connection()
        close_after = True

    # Dọn dẹp các log cũ quá 60 giây
    conn.execute("DELETE FROM submission_rate_logs WHERE created_ts < ?", (cutoff_ts,))
    conn.commit()

    rows = conn.execute(
        "SELECT created_ts FROM submission_rate_logs WHERE LOWER(email) = LOWER(?) AND created_ts >= ? ORDER BY created_ts ASC",
        (target_email, cutoff_ts)
    ).fetchall()

    if close_after:
        conn.close()

    if len(rows) >= RATE_LIMIT_MAX_REQUESTS:
        oldest_ts = float(rows[0]['created_ts'])
        retry_after = max(1, int(round((oldest_ts + RATE_LIMIT_WINDOW_SECONDS) - now_ts)))
        return False, retry_after

    return True, 0


def record_submission_rate_log(email=None, action_type='submit', conn=None):
    """Ghi nhận 1 lượt gửi đơn vào bảng kiểm soát tần suất (Rate Limiting)."""
    target_email = (email if email is not None else get_current_user_email() or '').strip().lower()
    if not target_email:
        return
    close_after = False
    if conn is None:
        conn = get_db_connection()
        close_after = True
    conn.execute(
        "INSERT INTO submission_rate_logs (email, action_type, created_ts) VALUES (?, ?, ?)",
        (target_email, action_type, time.time())
    )
    if close_after:
        conn.commit()
        conn.close()


def get_teacher_passcode(conn=None):
    """Lấy mã passcode nội bộ dành cho Thầy/Cô dùng Gmail cá nhân."""
    close_after = False
    if conn is None:
        conn = get_db_connection()
        close_after = True
    row = conn.execute("SELECT value FROM system_settings WHERE key = 'teacher_passcode'").fetchone()
    if close_after:
        conn.close()
    return row['value'] if (row and row['value']) else DEFAULT_TEACHER_PASSCODE


def is_current_user_verified(conn=None):
    """Kiểm tra tài khoản hiện tại đã được xác thực tư cách Giảng viên hay chưa.
    - Tài khoản @hcmus.edu.vn hoặc ADMIN_EMAILS: Tự động hợp lệ (Trusted).
    - Tài khoản @gmail.com hoặc domain khác: Phải có trong bảng verified_users (đã nhập Passcode).
    """
    email = get_current_user_email()
    if not email:
        return False
    if email in ADMIN_EMAILS:
        return True
    domain = email.split('@')[-1].lower()
    if domain == 'hcmus.edu.vn':
        return True

    close_after = False
    if conn is None:
        conn = get_db_connection()
        close_after = True
    row = conn.execute("SELECT id FROM verified_users WHERE LOWER(email) = LOWER(?)", (email,)).fetchone()
    if close_after:
        conn.close()

    is_ver = bool(row)
    session['is_verified'] = is_ver
    return is_ver


@app.before_request
def enforce_blocked_users():
    """Tự động kiểm tra và ngắt phiên truy cập ngay lập tức nếu tài khoản đang đăng nhập bị Admin khóa."""
    if request.endpoint in ('static', 'logout', 'login_page'):
        return None
    email = get_current_user_email()
    if email and is_user_blocked(email):
        session.clear()
        msg = f"Tài khoản ({email}) đã bị Quản trị viên khóa quyền truy cập hệ thống vĩnh viễn."
        if request.is_json or request.path.startswith('/submit_bulk_booking'):
            return jsonify({'success': False, 'blocked': True, 'message': msg}), 403
        flash(msg, "danger")
        return redirect(url_for('login_page'))
    return None


@app.context_processor
def inject_user_info():
    email = get_current_user_email()
    is_verified = is_current_user_verified() if email else False
    is_gmail = email.lower().endswith('@gmail.com') if email else False
    return dict(
        ADMIN_EMAILS=ADMIN_EMAILS,
        current_user_email=email,
        current_user_name=get_current_user_name(),
        is_user_verified=is_verified,
        is_user_gmail=is_gmail
    )


def login_required(f):
    """Yêu cầu đăng nhập Google trước khi thực hiện thao tác.
    Đồng thời kiểm tra tài khoản bị khóa và yêu cầu tài khoản (@gmail.com) xác minh Mã Passcode Giảng viên."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        email = get_current_user_email()
        if not email:
            if request.is_json or request.path.startswith('/submit_bulk_booking'):
                return jsonify({'success': False, 'message': 'Vui lòng đăng nhập bằng tài khoản Google trước khi gửi yêu cầu.'}), 401
            flash("Vui lòng đăng nhập bằng tài khoản Google để tiếp tục.", "warning")
            return redirect(url_for('login_page', next=request.url))

        # Kiểm tra tài khoản có bị khóa vĩnh viễn hay không
        if is_user_blocked(email):
            session.clear()
            msg = f"Tài khoản ({email}) đã bị Quản trị viên khóa quyền truy cập hệ thống."
            if request.is_json or request.path.startswith('/submit_bulk_booking'):
                return jsonify({'success': False, 'blocked': True, 'message': msg}), 403
            flash(msg, "danger")
            return redirect(url_for('login_page'))

        # Kiểm tra xác thực Giảng viên đối với tài khoản Gmail cá nhân
        if not is_current_user_verified():
            if request.is_json or request.path.startswith('/submit_bulk_booking'):
                return jsonify({
                    'success': False,
                    'require_verification': True,
                    'message': 'Tài khoản cá nhân chưa xác minh Mã Passcode Giảng viên. Vui lòng xác minh để tiếp tục.'
                }), 403
            flash("Tài khoản Gmail của Quý Thầy/Cô cần xác minh Mã Passcode Giảng viên trước khi thực hiện thao tác.", "warning")
            return redirect(url_for('verify_passcode_page', next=request.url))

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
# HỆ THỐNG GỬI EMAIL THÔNG BÁO KẾT QUẢ TỰ ĐỘNG CHO USER (SMTP)
# ==========================================

def get_email_settings(conn=None):
    """Đọc cấu hình SMTP gửi email tự động từ bảng system_settings hoặc biến môi trường."""
    close_after = False
    if conn is None:
        conn = get_db_connection()
        close_after = True

    rows = conn.execute(
        "SELECT key, value FROM system_settings WHERE key IN ('smtp_enabled', 'smtp_email', 'smtp_password', 'smtp_sender_name')"
    ).fetchall()
    if close_after:
        conn.close()

    cfg = {r['key']: (r['value'] or '') for r in rows}
    smtp_enabled = cfg.get('smtp_enabled', '1') == '1'
    smtp_email = (cfg.get('smtp_email') or os.environ.get('SMTP_EMAIL', 'mlhang@hcmus.edu.vn')).strip()
    smtp_password = (cfg.get('smtp_password') or os.environ.get('SMTP_PASSWORD', '')).strip()
    smtp_sender_name = (cfg.get('smtp_sender_name') or 'Hệ thống Phòng máy B4.2 - Phòng QTTB').strip()

    return {
        'enabled': smtp_enabled,
        'email': smtp_email,
        'password': smtp_password,
        'sender_name': smtp_sender_name,
        'has_password': bool(smtp_password),
        'ready': bool(smtp_enabled and smtp_email and smtp_password)
    }


def send_smtp_email_sync(to_email, subject, html_body, text_body, email_cfg=None):
    """Thực hiện gửi email qua Gmail SMTP (TLS port 587). Trả về (success, message)."""
    if email_cfg is None:
        email_cfg = get_email_settings()

    to_email = (to_email or '').strip()
    if not to_email or '@' not in to_email:
        return False, "Yêu cầu không có địa chỉ email người nhận hợp lệ."

    if not email_cfg.get('enabled'):
        return False, "Tính năng gửi email tự động đang tắt."

    sender_email = email_cfg.get('email', '').strip()
    sender_pass = email_cfg.get('password', '').replace(' ', '').strip()
    sender_name = email_cfg.get('sender_name') or 'Hệ thống Phòng máy B4.2 - Phòng QTTB'

    if not sender_email or not sender_pass:
        return False, "Chưa cấu hình Mật khẩu ứng dụng (Google App Password) cho email gửi đi."

    msg = MIMEMultipart('alternative')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = formataddr((str(Header(sender_name, 'utf-8')), sender_email))
    msg['To'] = to_email

    part_text = MIMEText(text_body, 'plain', 'utf-8')
    part_html = MIMEText(html_body, 'html', 'utf-8')
    msg.attach(part_text)
    msg.attach(part_html)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587, timeout=15) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender_email, sender_pass)
            server.sendmail(sender_email, [to_email], msg.as_string())
        return True, f"Đã gửi email tới {to_email}"
    except Exception as e:
        return False, f"Lỗi SMTP: {e}"


def build_request_notification_email(req_type, record, new_status, admin_note='', dashboard_url=''):
    """Tạo tiêu đề, nội dung HTML và Plain-text trang trọng (không dùng icon) cho email trả kết quả yêu cầu của User."""
    rec = dict(record)
    room = str(rec.get('room') or '').replace('Phòng ', '').replace('P.', '').strip()
    user_email = (rec.get('email') or '').strip()
    note_clean = (admin_note or rec.get('admin_note') or '').strip()
    now_str = datetime.datetime.now().strftime('%H:%M %d/%m/%Y')

    if req_type == 'booking':
        user_name = rec.get('requester_name') or user_email.split('@')[0]
        type_title = "YÊU CẦU ĐĂNG KÝ MƯỢN PHÒNG MÁY"
        day_str = f"Thứ {rec.get('day_of_week')}"
        period_str = f"Tiết {rec.get('start_period'):g} - {rec.get('end_period'):g}" if isinstance(rec.get('start_period'), (int, float)) else f"Tiết {rec.get('start_period')} - {rec.get('end_period')}"
        subj_info = rec.get('subject_name') or rec.get('purpose') or 'Mượn phòng máy'
        subject = f"Phản hồi yêu cầu đăng ký mượn phòng máy P.{room} ({day_str}, {period_str}) - Phòng máy tính B4.2"

        if new_status == 'Processing':
            status_badge = "ĐÃ TIẾP NHẬN & ĐANG XỬ LÝ"
            badge_bg, badge_color = "#E0F2FE", "#0369A1"
            lead_msg = f"Bộ phận Quản lý Phòng máy (Phòng QTTB) đã tiếp nhận yêu cầu mượn <strong>Phòng {room}</strong> của Quý Thầy/Cô và đang tiến hành kiểm tra lịch phòng / chuẩn bị phần mềm."
        elif new_status == 'Approved':
            status_badge = "ĐÃ DUYỆT MƯỢN PHÒNG"
            badge_bg, badge_color = "#DCFCE7", "#15803D"
            lead_msg = f"Yêu cầu đăng ký mượn <strong>Phòng {room}</strong> ({day_str}, {period_str}) của Quý Thầy/Cô đã được <strong>DUYỆT</strong> và cập nhật vào Thời khóa biểu vận hành phòng máy."
        else:
            status_badge = "TỪ CHỐI YÊU CẦU"
            badge_bg, badge_color = "#FEE2E2", "#B91C1C"
            lead_msg = f"Rất tiếc, yêu cầu đăng ký mượn <strong>Phòng {room}</strong> ({day_str}, {period_str}) hiện chưa thể đáp ứng. Quý Thầy/Cô vui lòng xem chi tiết phản hồi bên dưới."

        detail_rows = [
            ("Người đăng ký", f"{user_name} ({rec.get('department') or 'Đơn vị'})"),
            ("Phòng máy", f"Phòng {room} - Tòa B4.2"),
            ("Thời gian đăng ký", f"{day_str} | {period_str} (Tuần/Ngày: {rec.get('study_period') or 'Hiện tại'})"),
            ("Môn học / Lớp", f"{subj_info} {(' - Lớp ' + str(rec.get('class_code'))) if rec.get('class_code') else ''}"),
            ("Mục đích & Phần mềm", f"{rec.get('purpose') or 'Giảng dạy'}{(' - Phần mềm yêu cầu: ' + str(rec.get('software_required'))) if rec.get('software_required') else ''}")
        ]

    elif req_type == 'issue':
        user_name = rec.get('reporter_name') or user_email.split('@')[0]
        type_title = "BÁO CÁO SỰ CỐ KỸ THUẬT PHÒNG MÁY"
        machine_no = rec.get('machine_number') or 'N/A'
        subject = f"Phản hồi yêu cầu báo cáo sự cố kỹ thuật P.{room} (Máy {machine_no}) - Phòng máy tính B4.2"

        if new_status == 'Processing':
            status_badge = "ĐÃ TIẾP NHẬN & ĐANG XỬ LÝ"
            badge_bg, badge_color = "#E0F2FE", "#0369A1"
            lead_msg = f"Kỹ thuật viên Phòng máy B4.2 đã tiếp nhận thông tin báo sự cố tại <strong>Phòng {room} (Máy {machine_no})</strong> và đang tiến hành kiểm tra, xử lý."
        else:
            status_badge = "ĐÃ KHẮC PHỤC / HOÀN TẤT"
            badge_bg, badge_color = "#DCFCE7", "#15803D"
            lead_msg = f"Sự cố kỹ thuật tại <strong>Phòng {room} (Máy {machine_no})</strong> do Quý Thầy/Cô phản ánh đã được Kỹ thuật viên Phòng máy kiểm tra và khắc phục hoàn tất."

        detail_rows = [
            ("Người báo sự cố", f"{user_name} ({rec.get('department') or 'Đơn vị'})"),
            ("Vị trí thiết bị", f"Phòng {room} - Máy số: {machine_no}"),
            ("Phân loại sự cố", str(rec.get('issue_category') or 'Sự cố kỹ thuật')),
            ("Mô tả chi tiết", str(rec.get('description') or ''))
        ]

    else:  # cancellation
        user_name = rec.get('teacher_name') or user_email.split('@')[0]
        type_title = "THÔNG BÁO NGHỈ HỌC / BÁO VẮNG ĐỘT XUẤT"
        cancel_date = rec.get('cancel_date') or ''
        period_str = rec.get('period') or ''
        subject = f"Phản hồi yêu cầu thông báo nghỉ học P.{room} (Ngày {cancel_date}) - Phòng máy tính B4.2"

        if new_status == 'Processing':
            status_badge = "ĐÃ TIẾP NHẬN THÔNG BÁO NGHỈ"
            badge_bg, badge_color = "#E0F2FE", "#0369A1"
            lead_msg = f"Bộ phận Quản lý Phòng máy B4.2 đã tiếp nhận thông báo nghỉ học tại <strong>Phòng {room}</strong> (Ngày {cancel_date}, {period_str})."
        else:
            status_badge = "ĐÃ XÁC NHẬN & CẬP NHẬT LỊCH TRỰC"
            badge_bg, badge_color = "#DCFCE7", "#15803D"
            lead_msg = f"Thông báo nghỉ học tại <strong>Phòng {room}</strong> (Ngày {cancel_date}, {period_str}) của Quý Thầy/Cô đã được xác nhận và chuyển thông tin tới bộ phận Trực đóng/mở cửa phòng máy."

        detail_rows = [
            ("Giảng viên thông báo", str(user_name)),
            ("Phòng máy", f"Phòng {room} - Tòa B4.2"),
            ("Ngày báo nghỉ", str(cancel_date)),
            ("Ca / Tiết nghỉ", str(period_str)),
            ("Lý do", str(rec.get('reason') or 'Báo nghỉ đột xuất'))
        ]

    rows_html = ""
    rows_text = []
    for k, v in detail_rows:
        rows_html += f"""
        <tr>
            <td style="padding: 10px 14px; border-bottom: 1px solid #E2E8F0; color: #475569; font-weight: 600; width: 36%; font-size: 14px;">{html_lib.escape(str(k))}</td>
            <td style="padding: 10px 14px; border-bottom: 1px solid #E2E8F0; color: #0F172A; font-weight: 500; font-size: 14px;">{html_lib.escape(str(v))}</td>
        </tr>
        """
        rows_text.append(f"- {k}: {v}")

    note_block_html = ""
    note_block_text = ""
    if note_clean:
        note_block_html = f"""
        <div style="margin-top: 18px; padding: 14px 16px; background-color: #F0FDFA; border-left: 4px solid #0D9488; border-radius: 6px;">
            <div style="font-size: 13px; font-weight: 700; color: #0F766E; text-transform: uppercase; margin-bottom: 4px;">
                Phản hồi / Ghi chú từ Quản trị viên Phòng máy (QTTB):
            </div>
            <div style="font-size: 14px; color: #134E4A; line-height: 1.5;">
                {html_lib.escape(note_clean)}
            </div>
        </div>
        """
        note_block_text = f"\n[Phản hồi từ QTTB]: {note_clean}\n"

    dashboard_btn_html = ""
    if dashboard_url:
        dashboard_btn_html = f"""
        <div style="text-align: center; margin-top: 22px;">
            <a href="{dashboard_url}" style="display: inline-block; background-color: #0F2942; color: #FFFFFF; text-decoration: none; font-weight: 700; font-size: 14px; padding: 10px 24px; border-radius: 6px;">
                Xem chi tiết trên Hệ thống Quản lý Phòng máy
            </a>
        </div>
        """

    signature_text = "Bộ phận quản lý phòng máy"

    html_body = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin: 0; padding: 20px; background-color: #F1F5F9; font-family: 'Segoe UI', Roboto, Arial, sans-serif; color: #1E293B;">
    <div style="max-width: 640px; margin: 0 auto; background-color: #FFFFFF; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 15px rgba(15, 41, 66, 0.08); border: 1px solid #E2E8F0;">
        <div style="background: linear-gradient(135deg, #0F2942 0%, #1E3A5F 100%); padding: 20px 26px; color: #FFFFFF; text-align: center;">
            <h2 style="margin: 0; font-size: 18px; font-weight: 700; color: #FFFFFF; letter-spacing: 0.5px; text-align: center;">
                PHÒNG MÁY TÍNH B4.2
            </h2>
        </div>

        <div style="padding: 26px;">
            <div style="margin-bottom: 18px;">
                <span style="display: inline-block; background-color: {badge_bg}; color: {badge_color}; font-weight: 700; font-size: 12.5px; padding: 6px 14px; border-radius: 4px; letter-spacing: 0.3px;">
                    TRẠNG THÁI: {status_badge}
                </span>
                <div style="font-size: 12px; color: #64748B; margin-top: 6px;">Thời gian cập nhật: {now_str}</div>
            </div>

            <p style="font-size: 15px; line-height: 1.6; margin-top: 0;">
                Kính gửi Thầy/Cô <strong>{html_lib.escape(str(user_name))}</strong>,
            </p>
            <p style="font-size: 14.5px; line-height: 1.6; color: #334155;">
                {lead_msg}
            </p>

            {note_block_html}

            <div style="margin-top: 20px; border: 1px solid #E2E8F0; border-radius: 8px; overflow: hidden;">
                <div style="background-color: #F8FAFC; padding: 10px 14px; font-size: 12.5px; font-weight: 700; color: #0F2942; border-bottom: 1px solid #E2E8F0; text-transform: uppercase;">
                    THÔNG TIN CHI TIẾT — {type_title}
                </div>
                <table style="width: 100%; border-collapse: collapse;">
                    {rows_html}
                </table>
            </div>

            {dashboard_btn_html}

            <div style="margin-top: 26px; padding-top: 16px; border-top: 1px solid #E2E8F0; font-size: 14px; color: #1E293B; line-height: 1.6;">
                <div>Trân trọng,</div>
                <div style="font-weight: 700; color: #0F2942; margin-top: 4px;">
                    {signature_text}
                </div>
            </div>
        </div>

        <div style="background-color: #0F2942; color: #CBD5E1; padding: 16px 26px; font-size: 12px; line-height: 1.6;">
            <div style="color: #FFFFFF; font-weight: 700;">
                Hệ thống Phòng máy tính - Phòng Quản trị Thiết bị - Tòa B4.2 – Cơ sở 2
            </div>
            <div>Điện thoại hỗ trợ kỹ thuật: 0853212259 – Email: <a href="mailto:mlhang@hcmus.edu.vn" style="color: #2DD4BF; text-decoration: none;">mlhang@hcmus.edu.vn</a> | Trực đóng/mở cửa: 0931437955</div>
        </div>
    </div>
</body>
</html>"""

    text_body = (
        f"Kính gửi Thầy/Cô {user_name},\n\n"
        f"Trạng thái xử lý ({type_title}): {status_badge}\n"
        f"{re.sub(r'<[^>]+>', '', lead_msg)}\n"
        f"{note_block_text}\n"
        f"Thông tin chi tiết:\n" + "\n".join(rows_text) + "\n\n"
        f"Trân trọng,\n"
        f"{signature_text}\n"
    )

    return subject, html_body, text_body


def notify_user_on_status_change(req_type, record, new_status, admin_note='', conn=None):
    """Gửi email thông báo kết quả xử lý yêu cầu tới email đăng nhập của User (chạy ngầm qua Thread)."""
    if not record:
        return ""
    rec = dict(record)
    to_email = (rec.get('email') or '').strip()
    if not to_email or '@' not in to_email:
        return " (Yêu cầu này không có email người gửi)"

    email_cfg = get_email_settings(conn=conn)
    if not email_cfg.get('enabled'):
        return ""

    try:
        dashboard_url = url_for('my_dashboard', _external=True)
    except Exception:
        dashboard_url = ""

    subject, html_body, text_body = build_request_notification_email(
        req_type=req_type,
        record=rec,
        new_status=new_status,
        admin_note=admin_note,
        dashboard_url=dashboard_url
    )

    if not email_cfg.get('ready'):
        return f" (⚠️ Chưa cấu hình App Password gửi mail tự động cho {to_email} — bấm '📧 Cấu hình Email' hoặc nút ✉️ để gửi)"

    def _bg_worker():
        send_smtp_email_sync(to_email, subject, html_body, text_body, email_cfg=email_cfg)

    threading.Thread(target=_bg_worker, daemon=True).start()
    return f" (📧 Đã gửi email thông báo kết quả tới {to_email})"


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

        if is_user_blocked(email):
            flash(f"Tài khoản ({email}) đã bị Quản trị viên khóa quyền truy cập hệ thống vĩnh viễn. Vui lòng liên hệ Phòng QTTB.", "danger")
            return redirect(url_for('login_page'))

        session['user_email'] = email
        session['user_name'] = name or email.split('@')[0]

        next_url = session.pop('next_url', None) or url_for('home_page')
        flash(f"Xin chào {session['user_name']} ({email})!", "success")

        # Kiểm tra nếu tài khoản cá nhân chưa xác minh thì chuyển hướng nhập Passcode
        if not is_current_user_verified():
            flash("Quý Thầy/Cô sử dụng tài khoản cá nhân (@gmail.com) vui lòng nhập Mã Passcode Giảng viên do Phòng QTTB cung cấp (chỉ nhập một lần duy nhất).", "info")
            return redirect(url_for('verify_passcode_page', next=next_url))

        return redirect(next_url)

    except Exception as e:
        flash(f"Lỗi kết nối xác thực Google OAuth: {e}", "danger")
        return redirect(url_for('login_page'))


@app.route('/logout')
def logout():
    session.clear()
    flash("Đã đăng xuất khỏi hệ thống.", "success")
    return redirect(url_for('home_page'))


@app.route('/xac-minh-giang-vien', methods=['GET', 'POST'])
def verify_passcode_page():
    """Trang xác thực mã Passcode dành cho Thầy/Cô sử dụng tài khoản Gmail cá nhân."""
    user_email = get_current_user_email()
    if not user_email:
        flash("Vui lòng đăng nhập Google trước khi thực hiện xác minh.", "warning")
        return redirect(url_for('login_page', next=request.url))

    next_url = request.args.get('next') or request.form.get('next') or url_for('home_page')

    # Nếu tài khoản đã xác minh rồi thì chuyển tiếp luôn
    if is_current_user_verified():
        flash("Tài khoản của Quý Thầy/Cô đã được xác minh thành công trước đó.", "info")
        return redirect(next_url)

    error_message = None

    if request.method == 'POST':
        passcode = (request.form.get('passcode') or '').strip()
        expected_passcode = get_teacher_passcode()

        if passcode == expected_passcode:
            conn = get_db_connection()
            user_name = get_current_user_name()
            conn.execute(
                "INSERT OR REPLACE INTO verified_users (email, full_name, verified_by) VALUES (?, ?, 'passcode')",
                (user_email.lower(), user_name)
            )
            conn.commit()
            conn.close()
            session['is_verified'] = True
            flash("Xác minh thành công! Quý Thầy/Cô hiện có toàn quyền đăng ký mượn phòng máy và gửi thông tin.", "success")
            return redirect(next_url)
        else:
            error_message = "Mã Passcode không chính xác! Quý Thầy/Cô vui lòng kiểm tra lại hoặc liên hệ Bộ phận Phòng máy B4.2 (Phòng QTTB) để nhận mã."

    return render_template(
        'verify_passcode.html',
        user_email=user_email,
        user_name=get_current_user_name(),
        next_url=next_url,
        error_message=error_message
    )


# ==========================================
# PHÂN HỆ USER (PUBLIC & AUTHENTICATED)
# ==========================================

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

@app.route('/lien-he')
def contact_page():
    """Trang thông tin liên hệ hỗ trợ công khai (không cần đăng nhập)"""
    return render_template('contact.html')

@app.route('/bao-cao-su-co')
@login_required
def report_page():
    return render_template('report.html')

@app.route('/submit_issue', methods=['POST'])
@login_required
def submit_issue():
    user_email = get_current_user_email()
    allowed, retry_after = check_submission_rate_limit(user_email)
    if not allowed:
        flash(f"⚠️ Giới hạn tốc độ gửi: Mỗi tài khoản chỉ được gửi tối đa 2 đơn trong vòng 1 phút để chống spam. Vui lòng đợi {retry_after} giây nữa trước khi gửi tiếp.", "danger")
        return redirect(url_for('report_page'))

    reporter_name = request.form.get('reporter_name') or get_current_user_name()
    department = request.form.get('department')
    room = request.form.get('room')
    machine_number = request.form.get('machine_number')
    issue_category = request.form.get('issue_category')
    description = request.form.get('description')
    
    conn = get_db_connection()
    try:
        conn.execute('''
        INSERT INTO issue_reports (reporter_name, department, room, machine_number, issue_category, description, email)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (reporter_name, department, room, machine_number, issue_category, description, user_email))
        record_submission_rate_log(user_email, action_type='issue', conn=conn)
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
    return render_template('booking.html')

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
    user_email = get_current_user_email()
    allowed, retry_after = check_submission_rate_limit(user_email)
    if not allowed:
        flash(f"⚠️ Giới hạn tốc độ gửi: Mỗi tài khoản chỉ được gửi tối đa 2 đơn trong vòng 1 phút để chống spam. Vui lòng đợi {retry_after} giây nữa trước khi gửi tiếp.", "danger")
        return redirect(url_for('booking_page'))

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
    conn.execute('''
        INSERT INTO booking_requests (requester_name, room, day_of_week, start_period, end_period, purpose, subject_code, class_code, subject_name, department, study_period, software_required, status, email, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pending', ?, datetime('now', 'localtime'))
    ''', (requester_name, room, day_of_week, start_period, end_period, purpose, subject_code, class_code, subject_name, department, study_period, software_required, user_email))
    record_submission_rate_log(user_email, action_type='booking', conn=conn)
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
    user_email = get_current_user_email()
    allowed, retry_after = check_submission_rate_limit(user_email)
    if not allowed:
        return jsonify({
            'success': False,
            'rate_limited': True,
            'message': f'⚠️ Giới hạn tốc độ gửi: Mỗi tài khoản chỉ được gửi tối đa 2 đơn trong vòng 1 phút để chống spam. Vui lòng đợi {retry_after} giây nữa trước khi gửi tiếp.'
        }), 429

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
        record_submission_rate_log(user_email, action_type='bulk_booking', conn=conn)
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
    user_email = get_current_user_email()
    allowed, retry_after = check_submission_rate_limit(user_email)
    if not allowed:
        flash(f"⚠️ Giới hạn tốc độ gửi: Mỗi tài khoản chỉ được gửi tối đa 2 đơn trong vòng 1 phút để chống spam. Vui lòng đợi {retry_after} giây nữa trước khi gửi tiếp.", "danger")
        return redirect(url_for('cancellation_page'))

    teacher_name = request.form.get('teacher_name') or get_current_user_name()
    room = request.form.get('room')
    cancel_date = request.form.get('cancel_date')
    period = request.form.get('period')
    reason = request.form.get('reason')
    
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO cancellations (teacher_name, room, cancel_date, period, reason, email)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (teacher_name, room, cancel_date, period, reason, user_email))
    record_submission_rate_log(user_email, action_type='cancellation', conn=conn)
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

    # 1. Lấy toàn bộ danh sách yêu cầu mượn phòng (ưu tiên Chờ xử lý lên đầu, lưu vết các yêu cầu đã xử lý)
    booking_requests = conn.execute("""
        SELECT * FROM booking_requests
        ORDER BY CASE status WHEN 'Pending' THEN 0 WHEN 'Processing' THEN 1 WHEN 'Approved' THEN 2 ELSE 3 END, id DESC
    """).fetchall()
    pending_requests = [r for r in booking_requests if r['status'] == 'Pending']

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
    rooms = ROOMS_LIST
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

    # 4. Tính toán số liệu thống kê cho Tab 📊 Dashboard & B42Lab NetAgent
    ak_row = conn.execute("SELECT value FROM system_settings WHERE key = 'netagent_key'").fetchone()
    netagent_key = ak_row['value'] if (ak_row and ak_row['value']) else "b42lab-secret-2026"

    net_rows = conn.execute("SELECT * FROM room_network_status").fetchall()
    net_map = {row['room']: dict(row) for row in net_rows}

    # Đếm số ca học theo từng phòng (theo chuẩn 1 ca = 2.5 tiết, tối đa 24 ca/tuần/phòng, 192 ca/tuần toàn hệ thống)
    room_stats = {}
    max_room_sessions = 24
    total_ca_used = 0
    day_ca_used = {d: 0 for d in range(2, 8)}

    for r in rooms:
        r_active = [s for s in all_enriched_schedules if s['room'] == r and s['door_open']]
        r_locked = [s for s in all_enriched_schedules if s['room'] == r and not s['door_open']]
        r_active_ca = sum(len(get_schedule_ca_slots(s['start_period'], s['end_period'])) for s in r_active)
        r_locked_ca = sum(len(get_schedule_ca_slots(s['start_period'], s['end_period'])) for s in r_locked)
        r_total_ca = r_active_ca + r_locked_ca
        if r_total_ca > max_room_sessions:
            max_room_sessions = r_total_ca
        for s in (r_active + r_locked):
            num_ca = len(get_schedule_ca_slots(s['start_period'], s['end_period']))
            total_ca_used += num_ca
            dow = int(s['day_of_week'])
            if dow in day_ca_used:
                day_ca_used[dow] += num_ca
        room_stats[r] = {
            'active': r_active_ca,
            'locked': r_locked_ca,
            'total': r_total_ca
        }

    for r in rooms:
        room_stats[r]['height_pct'] = int(round((room_stats[r]['total'] / max_room_sessions) * 85)) if room_stats[r]['total'] > 0 else 8

    # Công suất lấp đầy chuẩn (8 phòng x 24 ca/tuần = 192 ca/tuần; 1 ngày = 32 ca/toàn trường)
    utilization_pct = round((total_ca_used / 192.0) * 100, 2) if total_ca_used > 0 else 0
    peak_dow = max(day_ca_used, key=day_ca_used.get) if total_ca_used > 0 else 2
    peak_dow_pct = min(100, int(round((day_ca_used[peak_dow] / 32.0) * 100))) if total_ca_used > 0 else 0

    # Phân bổ theo Khoa / Đơn vị (Tự động nhận diện động & tính theo Số ca chuẩn 2.5 tiết)
    dept_map = clean_and_group_departments(all_enriched_schedules)
    dept_totals = {}
    for s in all_enriched_schedules:
        raw_dept = (s.get('department') or '').strip()
        label = dept_map.get(raw_dept, 'Khác')
        num_ca = len(get_schedule_ca_slots(s['start_period'], s['end_period']))
        dept_totals[label] = dept_totals.get(label, 0) + num_ca
    sorted_depts = sorted(dept_totals.items(), key=lambda x: (-x[1], x[0]))

    total_sched_ca = max(1, total_ca_used)
    palette = [
        '#2563eb',  # Xanh dương
        '#10b981',  # Xanh lá
        '#f59e0b',  # Cam vàng
        '#8b5cf6',  # Tím
        '#06b6d4',  # Xanh ngọc
        '#ec4899',  # Hồng
        '#f97316',  # Cam đậm
        '#14b8a6',  # Teal
        '#ef4444',  # Đỏ
        '#64748b',  # Xám xanh
    ]

    dept_items = []
    gradient_stops = []
    current_deg = 0.0

    for idx, (label, count) in enumerate(sorted_depts):
        color = palette[idx % len(palette)]
        pct = max(1, int(round((count / total_sched_ca) * 100))) if count > 0 else 0
        deg_span = (count / total_sched_ca) * 360.0
        next_deg = 360.0 if idx == len(sorted_depts) - 1 else (current_deg + deg_span)
        gradient_stops.append(f"{color} {current_deg:.1f}deg {next_deg:.1f}deg")
        current_deg = next_deg

        dept_items.append({
            'name': label,
            'count': count,
            'pct': pct,
            'color': color
        })

    donut_gradient = f"conic-gradient({', '.join(gradient_stops)})" if gradient_stops else "conic-gradient(#e2e8f0 0deg 360deg)"

    # Trạng thái trực tiếp 8 phòng hôm nay & Lịch Mở/Khóa cửa hôm nay
    today_dow = today.weekday() + 2  # Thứ 2 = 2 .. Chủ nhật = 8
    display_dow = today_dow if 2 <= today_dow <= 7 else 2
    today_str_dmy = today.strftime('%d/%m/%Y')
    today_str_iso = today.strftime('%Y-%m-%d')

    # Kiểm tra các phòng có báo nghỉ hôm nay
    cancelled_rooms_today = set()
    for c in cancellations:
        c_date = (c['cancel_date'] or '').strip()
        if c_date in (today_str_dmy, today_str_iso):
            cancelled_rooms_today.add(str(c['room']).replace('Phòng ', '').replace('P.', '').strip())

    pending_issues_by_room = {r: 0 for r in rooms}
    for iss in issues:
        if iss['status'] != 'Resolved':
            r_code = str(iss['room']).strip()
            if r_code in pending_issues_by_room:
                pending_issues_by_room[r_code] += 1

    room_live_cards = []
    online_rooms_count = 0
    offline_rooms_list = []
    last_net_update = ''

    for r in rooms:
        r_today_classes = [s for s in all_enriched_schedules if s['room'] == r and int(s['day_of_week']) == display_dow and s['is_in_selected_week']]

        # Nhãn hoạt động trên thẻ Live Status 8 phòng (luôn ưu tiên hiển thị lịch học/trạng thái cửa)
        open_today = [s for s in r_today_classes if s['door_open']]
        ended_today = [s for s in r_today_classes if not s['door_open']]
        if r in cancelled_rooms_today:
            activity_label = "🔒 Khóa cửa (Lớp báo nghỉ)"
            activity_class = "text-danger fw-semibold"
        elif open_today:
            first_c = open_today[0]
            subj = first_c['subject_name']
            if len(subj) > 20:
                subj = subj[:19] + "…"
            activity_label = f"🔓 T.{first_c['start_period']:g}-{first_c['end_period']:g}: {subj}"
            activity_class = "text-primary fw-semibold"
        elif ended_today:
            activity_label = "🔒 Khóa cửa (Đã KT sớm)"
            activity_class = "text-secondary"
        else:
            activity_label = f"⚪ Trống lịch Thứ {display_dow}"
            activity_class = "text-muted"

        net_info = net_map.get(r, {'target_ip': '8.8.8.8', 'status': 'online', 'ping_ms': 12, 'updated_at': ''})
        net_status = net_info.get('status') or 'online'
        net_ping = net_info.get('ping_ms')
        if net_info.get('updated_at'):
            last_net_update = net_info['updated_at']

        if net_status == 'online':
            online_rooms_count += 1
        else:
            offline_rooms_list.append(f"P.{r}")

        room_live_cards.append({
            'room': r,
            'activity_label': activity_label,
            'activity_class': activity_class,
            'issue_count': pending_issues_by_room.get(r, 0),
            'target_ip': net_info.get('target_ip') or '8.8.8.8',
            'net_status': net_status,
            'ping_ms': net_ping,
            'updated_at': net_info.get('updated_at') or ''
        })

    # Gom nhóm các khung tiết Không học / Đóng cửa hôm nay (Ca Sáng & Ca Chiều)
    today_m_free, today_a_free = get_detailed_free_slots_data(display_dow, ref_date=today)

    def _fmt_p_vn(val):
        f = float(val)
        return str(int(f)) if f.is_integer() else str(f).replace('.', ',')

    def _build_closed_groups(slot_items, is_morning=True):
        grouped = {}
        for item in slot_items:
            st = float(item['start_period'])
            ed = float(item['end_period'])
            key = (st, ed)
            room_label = f"P.{item['room']}"
            if room_label not in grouped.setdefault(key, []):
                grouped[key].append(room_label)

        full_key = (1.0, 5.0) if is_morning else (6.0, 10.0)
        for r_cancel in sorted(cancelled_rooms_today):
            r_lbl = f"P.{r_cancel}"
            already_listed = any(r_lbl in r_list for r_list in grouped.values())
            if not already_listed:
                grouped.setdefault(full_key, []).append(r_lbl)

        sorted_keys = sorted(grouped.keys(), key=lambda k: (1 if (k[1] - k[0]) >= 4.0 else 0, k[0], k[1]))
        result = []
        for st, ed in sorted_keys:
            period_code = f"T{_fmt_p_vn(st)} - {_fmt_p_vn(ed)}"
            if is_morning:
                if st == 1.0 and ed == 5.0:
                    note = "(Đóng nguyên buổi sáng)"
                elif st == 1.0:
                    note = "(Đóng nửa đầu ca sáng)"
                else:
                    note = "(Đóng nửa sau ca sáng)"
            else:
                if st == 6.0 and ed == 10.0:
                    note = "(Đóng nguyên buổi chiều)"
                elif st == 6.0:
                    note = "(Đóng nửa đầu ca chiều)"
                else:
                    note = "(Đóng nửa sau ca chiều)"
            result.append({
                'period_code': period_code,
                'rooms': grouped[(st, ed)],
                'rooms_str': ', '.join(grouped[(st, ed)]),
                'note': note
            })
        return result

    morning_closed_groups = _build_closed_groups(today_m_free, is_morning=True)
    afternoon_closed_groups = _build_closed_groups(today_a_free, is_morning=False)

    dashboard_stats = {
        'room_stats': room_stats,
        'total_ca_used': total_ca_used,
        'total_ca_free': max(0, 192 - total_ca_used),
        'utilization_pct': utilization_pct,
        'peak_dow': peak_dow,
        'peak_dow_pct': peak_dow_pct,
        'dept_items': dept_items,
        'donut_gradient': donut_gradient,
        'room_live_cards': room_live_cards,
        'display_dow': display_dow,
        'morning_closed_groups': morning_closed_groups,
        'afternoon_closed_groups': afternoon_closed_groups,
        'online_rooms_count': online_rooms_count,
        'offline_rooms_list': offline_rooms_list,
        'last_net_update': last_net_update,
        'netagent_key': netagent_key
    }

    semester_summary = summarize_available_semesters(raw_schedules)

    # 5. Dữ liệu Quản lý Lịch trình & Chấm công Tăng ca (OT)
    ot_month = request.args.get('ot_month', 'all').strip()
    ot_filter_status = request.args.get('ot_status', 'all').strip()

    month_query = conn.execute("SELECT DISTINCT substr(ot_date, 1, 7) as m FROM overtime_records WHERE ot_date IS NOT NULL AND length(ot_date) >= 7 ORDER BY m DESC").fetchall()
    available_ot_months = [row['m'] for row in month_query if row['m']]
    current_month_str = today.strftime('%Y-%m')
    if current_month_str not in available_ot_months:
        available_ot_months.append(current_month_str)
        available_ot_months.sort(reverse=True)

    if not ot_month:
        ot_month = 'all'

    if ot_month == 'all':
        all_ot_rows = conn.execute("SELECT * FROM overtime_records ORDER BY ot_date DESC, id DESC").fetchall()
    else:
        all_ot_rows = conn.execute("SELECT * FROM overtime_records WHERE ot_date LIKE ? ORDER BY ot_date DESC, id DESC", (f"{ot_month}%",)).fetchall()

    ot_total_hours = sum(float(r['hours'] or 0) for r in all_ot_rows)
    ot_completed_hours = sum(float(r['hours'] or 0) for r in all_ot_rows if r['status'] == 'Completed')
    ot_completed_count = sum(1 for r in all_ot_rows if r['status'] == 'Completed')
    ot_pending_count = sum(1 for r in all_ot_rows if r['status'] == 'Pending')
    ot_staff_set = set(r['staff_name'] for r in all_ot_rows if r['staff_name'])

    if ot_filter_status == 'Completed':
        display_ot_rows = [r for r in all_ot_rows if r['status'] == 'Completed']
    elif ot_filter_status == 'Pending':
        display_ot_rows = [r for r in all_ot_rows if r['status'] == 'Pending']
    else:
        display_ot_rows = all_ot_rows

    ot_data = {
        'records': display_ot_rows,
        'selected_month': ot_month,
        'selected_status': ot_filter_status,
        'available_months': available_ot_months,
        'total_hours': round(ot_total_hours, 1),
        'completed_hours': round(ot_completed_hours, 1),
        'completed_count': ot_completed_count,
        'pending_count': ot_pending_count,
        'staff_count': len(ot_staff_set)
    }

    # 6. Dữ liệu Tab Nhật ký công việc (Work Log) & Checklist Deploy 8 Phòng Máy (SOP)
    all_work_logs = conn.execute("SELECT * FROM work_logs ORDER BY work_date DESC, id DESC").fetchall()
    wl_total = len(all_work_logs)
    wl_completed = sum(1 for r in all_work_logs if r['status'] == 'completed')
    wl_in_progress = wl_total - wl_completed

    enriched_work_logs = []
    for r in all_work_logs:
        d = dict(r)
        d['duration_info'] = calculate_worklog_duration(d.get('work_date'), d.get('completed_date'), d.get('status'))
        enriched_work_logs.append(d)

    worklog_data = {
        'records': enriched_work_logs,
        'total_count': wl_total,
        'completed_count': wl_completed,
        'in_progress_count': wl_in_progress
    }

    # Chuẩn bị dữ liệu Checklist Deploy 8 Phòng Máy
    default_campaign = "HK1 (2026-2027) - Deploy Hệ thống chuẩn"
    camp_rows = conn.execute(
        "SELECT campaign_name FROM deployment_checklists GROUP BY campaign_name ORDER BY MAX(id) DESC"
    ).fetchall()
    campaigns = [r['campaign_name'] for r in camp_rows if r['campaign_name']]
    if not campaigns:
        now_short = datetime.datetime.now().strftime('%H:%M %d/%m/%Y')
        for r_code in rooms:
            conn.execute('''
                INSERT OR IGNORE INTO deployment_checklists (campaign_name, room, checked_steps, notes, updated_by, updated_at)
                VALUES (?, ?, '[]', '', 'Admin', ?)
            ''', (default_campaign, r_code, now_short))
        conn.commit()
        campaigns = [default_campaign]

    selected_campaign = request.args.get('deploy_campaign', '').strip()
    if not selected_campaign or selected_campaign not in campaigns:
        selected_campaign = campaigns[0]

    # Đảm bảo đủ 8 phòng cho đợt đang chọn
    now_short = datetime.datetime.now().strftime('%H:%M %d/%m/%Y')
    for r_code in rooms:
        conn.execute('''
            INSERT OR IGNORE INTO deployment_checklists (campaign_name, room, checked_steps, notes, updated_by, updated_at)
            VALUES (?, ?, '[]', '', 'Admin', ?)
        ''', (selected_campaign, r_code, now_short))
    conn.commit()

    deploy_rows = conn.execute(
        "SELECT * FROM deployment_checklists WHERE campaign_name = ?",
        (selected_campaign,)
    ).fetchall()

    room_deploy_map = {}
    for dr in deploy_rows:
        try:
            parsed_steps = json.loads(dr['checked_steps'] or '[]')
            if not isinstance(parsed_steps, list):
                parsed_steps = []
        except Exception:
            parsed_steps = []
        room_deploy_map[dr['room']] = {
            'room': dr['room'],
            'checked': parsed_steps,
            'notes': dr['notes'] or '',
            'updated_by': dr['updated_by'] or 'Admin',
            'updated_at': dr['updated_at'] or now_short
        }

    for r_code in rooms:
        if r_code not in room_deploy_map:
            room_deploy_map[r_code] = {
                'room': r_code,
                'checked': [],
                'notes': '',
                'updated_by': 'Admin',
                'updated_at': now_short
            }

    deploy_data = {
        'campaigns': campaigns,
        'selected_campaign': selected_campaign,
        'room_map': room_deploy_map,
        'subview': request.args.get('wl_view', 'daily').strip()
    }

    email_config = get_email_settings(conn=conn)

    # Passcode, Danh sách Giảng viên Gmail đã xác minh & Danh sách Tài khoản bị khóa
    teacher_passcode = get_teacher_passcode(conn=conn)
    verified_users = conn.execute("SELECT * FROM verified_users ORDER BY id DESC").fetchall()
    blocked_users = conn.execute("SELECT * FROM blocked_users ORDER BY id DESC").fetchall()
    blocked_emails = [row['email'].strip().lower() for row in blocked_users if row['email']]

    conn.close()
    return render_template(
        'admin_panel.html',
        booking_requests=booking_requests,
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
        issues=issues,
        dashboard_stats=dashboard_stats,
        semester_summary=semester_summary,
        ot_data=ot_data,
        worklog_data=worklog_data,
        deploy_data=deploy_data,
        email_config=email_config,
        teacher_passcode=teacher_passcode,
        verified_users=verified_users,
        blocked_users=blocked_users,
        blocked_emails=blocked_emails
    )

def measure_host_ping_ms(target_ip):
    """Đo ping (ms) tới 1 địa chỉ IP/Host trên cả Windows và Linux."""
    target_ip = (target_ip or '').strip()
    if not target_ip:
        return False, None

    is_win = platform.system().lower().startswith('win')
    cmd = ['ping', '-n', '1', '-w', '1000', target_ip] if is_win else ['ping', '-c', '1', '-W', '1', target_ip]
    try:
        t0 = time.time()
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=2.5)
        elapsed_ms = max(1, int(round((time.time() - t0) * 1000)))
        if res.returncode == 0:
            m = re.search(r'(?:time|thời gian)[=<](\d+)\s*ms', res.stdout, re.IGNORECASE)
            ping_val = int(m.group(1)) if m else elapsed_ms
            return True, ping_val
        return False, None
    except Exception:
        return False, None

@app.route('/api/network-heartbeat', methods=['POST'])
def api_network_heartbeat():
    """API nhận tín hiệu Heartbeat Push từ B42Lab NetAgent chạy tại phòng máy."""
    data = request.get_json(silent=True) or {}
    provided_key = (data.get('agent_key') or request.headers.get('X-Agent-Key') or '').strip()

    conn = get_db_connection()
    ak_row = conn.execute("SELECT value FROM system_settings WHERE key = 'netagent_key'").fetchone()
    expected_key = ak_row['value'] if ak_row and ak_row['value'] else 'b42lab-secret-2026'

    if provided_key != expected_key:
        conn.close()
        return jsonify({'success': False, 'message': 'Invalid agent_key'}), 403

    now_str = datetime.datetime.now().strftime('%H:%M:%S %d/%m/%Y')
    results = data.get('rooms', {})
    # Hỗ trợ cả trường hợp gửi lẻ 1 phòng: {"room": "5.1", "status": "online", "ping_ms": 11}
    if not results and data.get('room'):
        results = {str(data['room']): {'status': data.get('status', 'online'), 'ping_ms': data.get('ping_ms')}}

    for r_code, r_info in results.items():
        if r_code in ROOMS_LIST:
            st = 'online' if r_info.get('status') in ('online', 1, True, '1') else 'offline'
            p_ms = r_info.get('ping_ms')
            conn.execute('''
                UPDATE room_network_status
                SET status = ?, ping_ms = ?, updated_at = ?
                WHERE room = ?
            ''', (st, p_ms, now_str, r_code))

    conn.commit()
    rows = conn.execute("SELECT room, target_ip FROM room_network_status").fetchall()
    targets = {row['room']: row['target_ip'] for row in rows}
    conn.close()

    return jsonify({
        'success': True,
        'updated_at': now_str,
        'targets': targets
    })

@app.route('/api/network-status')
@admin_required
def api_network_status():
    """API trả về trạng thái mạng 8 phòng máy theo thời gian thực cho Dashboard."""
    conn = get_db_connection()
    rows = conn.execute("SELECT room, target_ip, status, ping_ms, updated_at FROM room_network_status").fetchall()
    conn.close()

    room_map = {}
    online_count = 0
    offline_rooms = []
    last_update = ''

    for row in rows:
        r = row['room']
        st = row['status'] or 'online'
        if st == 'online':
            online_count += 1
        else:
            offline_rooms.append(f"P.{r}")
        if row['updated_at']:
            last_update = row['updated_at']
        room_map[r] = {
            'room': r,
            'target_ip': row['target_ip'],
            'status': st,
            'ping_ms': row['ping_ms'],
            'updated_at': row['updated_at'] or ''
        }

    return jsonify({
        'success': True,
        'online_count': online_count,
        'total_rooms': len(rows),
        'offline_rooms': offline_rooms,
        'last_update': last_update,
        'room_map': room_map
    })

@app.route('/admin/network-config', methods=['POST'])
@admin_required
def save_network_config():
    """Lưu cấu hình IP của 8 phòng máy từ giao diện Admin."""
    conn = get_db_connection()
    for r in ROOMS_LIST:
        ip_val = (request.form.get(f"ip_{r}") or '8.8.8.8').strip()
        conn.execute("UPDATE room_network_status SET target_ip = ? WHERE room = ?", (ip_val, r))
    conn.commit()
    conn.close()
    flash("✅ Đã lưu danh sách địa chỉ IP Router cho 8 phòng máy!", "success")
    return redirect(url_for('admin_panel', tab='dashboard'))

@app.route('/admin/network-ping-now', methods=['POST'])
@admin_required
def admin_network_ping_now():
    """Thực hiện Ping kiểm tra ngay lập tức tới 8 phòng máy."""
    conn = get_db_connection()
    rows = conn.execute("SELECT room, target_ip FROM room_network_status").fetchall()
    now_str = datetime.datetime.now().strftime('%H:%M:%S %d/%m/%Y')

    for row in rows:
        r = row['room']
        ip_addr = row['target_ip'] or '8.8.8.8'
        is_up, ping_ms = measure_host_ping_ms(ip_addr)
        st = 'online' if is_up else 'offline'
        conn.execute(
            "UPDATE room_network_status SET status = ?, ping_ms = ?, updated_at = ? WHERE room = ?",
            (st, ping_ms, now_str, r)
        )
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'updated_at': now_str})

@app.route('/admin/download-netagent')
@admin_required
def download_netagent_script():
    """Tải trực tiếp file script b42_network_agent.py đã điền sẵn 8 IP Router, URL và khóa bảo mật (trên RAM để không kích hoạt Flask reload)."""
    conn = get_db_connection()
    rows = conn.execute("SELECT room, target_ip FROM room_network_status").fetchall()
    conn.close()
    ip_dict = {row['room']: row['target_ip'] for row in rows}

    app_dir = os.path.dirname(os.path.abspath(__file__))
    agent_path = os.path.join(app_dir, 'b42_network_agent.py')
    if os.path.exists(agent_path):
        with open(agent_path, 'r', encoding='utf-8') as f:
            content = f.read()
        lines = ["ROOM_TARGETS = {"]
        for r in ROOMS_LIST:
            lines.append(f'    "{r}": "{ip_dict.get(r, "8.8.8.8")}",')
        lines.append("}")
        new_block = "\n".join(lines)
        content = re.sub(r'ROOM_TARGETS\s*=\s*\{[^}]*\}', new_block, content)
        current_server_url = request.host_url.rstrip('/')
        content = content.replace('DEFAULT_SERVER_URL = "http://127.0.0.1:5000"', f'DEFAULT_SERVER_URL = "{current_server_url}"')
        return Response(
            content,
            mimetype='application/octet-stream',
            headers={'Content-Disposition': 'attachment; filename="b42_network_agent.py"'}
        )
    return "File b42_network_agent.py not found", 404


@app.route('/admin/export-semester-report')
@admin_required
def export_semester_report():
    """Xuất file Excel (.xlsx) Báo cáo Tần suất Khai thác Phòng máy tính Học kỳ (4 Sheet chuẩn QTTB)."""
    semester_mode = request.args.get('semester', 'hk1').strip().lower()
    academic_year = request.args.get('academic_year', '').strip()
    from_date = request.args.get('from_date', '').strip()
    to_date = request.args.get('to_date', '').strip()

    conn = get_db_connection()
    schedules_rows = conn.execute("SELECT * FROM schedules ORDER BY room ASC, day_of_week ASC, start_period ASC").fetchall()
    issues_rows = conn.execute("SELECT * FROM issue_reports ORDER BY created_at DESC").fetchall()
    bookings_rows = conn.execute("SELECT * FROM booking_requests ORDER BY id DESC").fetchall()
    try:
        net_rows = conn.execute("SELECT * FROM room_network_status").fetchall()
    except Exception:
        net_rows = []
    conn.close()

    filter_res = filter_schedules_for_report(
        all_schedules=schedules_rows,
        semester_mode=semester_mode,
        academic_year=academic_year,
        from_date_str=from_date,
        to_date_str=to_date
    )

    if len(filter_res['rows']) == 0:
        flash(
            f"⚠️ Không tìm thấy dữ liệu môn học nào thuộc {filter_res['semester_short']} ({filter_res['academic_year']}) "
            f"trong Thời khóa biểu hiện tại! Vui lòng kiểm tra lại thời gian Bắt đầu – Kết thúc (Giai đoạn) của các môn học "
            f"hoặc Import lịch học cho kỳ này trước khi xuất báo cáo.",
            "warning"
        )
        return redirect(url_for('admin_panel', tab='schedules'))

    excel_buffer, filename = build_semester_report_excel(
        all_schedules=schedules_rows,
        issues_list=issues_rows,
        bookings_list=bookings_rows,
        net_status_list=net_rows,
        semester_mode=semester_mode,
        academic_year=academic_year,
        from_date_str=from_date,
        to_date_str=to_date
    )

    return send_file(
        excel_buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def build_overtime_report_excel(ot_rows, month_str):
    """Xuất file Excel bảng chấm công và theo dõi tăng ca (OT)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Bang Cham Cong OT"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Calibri", size=14, bold=True, color="0F2942")
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="475569")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill("solid", fgColor="0F2942")
    fill_summary = PatternFill("solid", fgColor="F0FDFA")
    font_summary = Font(name="Calibri", size=11, bold=True, color="0D9488")
    font_body = Font(name="Calibri", size=10, color="1F2937")
    font_bold = Font(name="Calibri", size=10, bold=True, color="1F2937")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    ws.merge_cells("A1:G1")
    ws["A1"] = "BẢNG TỔNG HỢP THEO DÕI VÀ CHẤM CÔNG TĂNG CA (OT)"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:G2")
    display_period = f"Tháng {month_str}" if month_str != 'all' else "Toàn bộ thời gian"
    ws["A2"] = f"Hệ thống Phòng máy tính Tòa B4.2 (Cơ sở 2) • Ký theo dõi: {display_period}"
    ws["A2"].font = font_subtitle
    ws["A2"].alignment = align_center

    headers = ["STT", "Ngày OT", "Nhân sự", "Bộ phận / Vị trí", "Khung giờ & Số giờ", "Nội dung / Nhiệm vụ tăng ca", "Trạng thái"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[4].height = 26

    total_hours = 0.0
    for idx, r in enumerate(ot_rows, 1):
        row_num = 4 + idx
        h_val = float(r['hours'] or 0)
        total_hours += h_val
        st_text = "Đã hoàn thành (Đã làm)" if r['status'] == 'Completed' else "Chưa tới hạn"
        time_display = f"{r['start_time']} - {r['end_time']} ({h_val:g}h)" if (r['start_time'] and r['end_time']) else f"{h_val:g} giờ"

        cells_data = [
            (idx, align_center, font_body),
            (r['ot_date'], align_center, font_bold),
            (r['staff_name'], align_left, font_bold),
            (r['role'], align_left, font_body),
            (time_display, align_center, font_body),
            (r['purpose'] or "", align_left, font_body),
            (st_text, align_center, font_body)
        ]
        for c_idx, (val, al, fn) in enumerate(cells_data, 1):
            c = ws.cell(row=row_num, column=c_idx, value=val)
            c.alignment = al
            c.font = fn
            c.border = thin_border
        ws.row_dimensions[row_num].height = 20

    sum_row = 5 + len(ot_rows)
    ws.merge_cells(f"A{sum_row}:D{sum_row}")
    ws[f"A{sum_row}"] = "TỔNG CỘNG SỐ GIỜ TĂNG CA:"
    ws[f"A{sum_row}"].font = font_summary
    ws[f"A{sum_row}"].alignment = align_right
    ws[f"A{sum_row}"].fill = fill_summary

    sum_cell = ws.cell(row=sum_row, column=5, value=f"{total_hours:g} giờ")
    sum_cell.font = font_summary
    sum_cell.alignment = align_center
    sum_cell.fill = fill_summary

    ws.merge_cells(f"F{sum_row}:G{sum_row}")
    ws[f"F{sum_row}"].fill = fill_summary

    for c in range(1, 8):
        ws.cell(row=sum_row, column=c).border = thin_border

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['F'].width = 36

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"Bang_Cham_Cong_OT_{month_str}.xlsx"
    return output, filename


@app.route('/admin/overtime/add', methods=['POST'])
@admin_required
def add_overtime():
    """Thêm một ca tăng ca (OT) mới."""
    ot_date = request.form.get('ot_date', '').strip()
    staff_name = request.form.get('staff_name', '').strip()
    role = request.form.get('role', 'Nhân viên Trực phòng').strip()
    start_time = request.form.get('start_time', '').strip()
    end_time = request.form.get('end_time', '').strip()
    raw_hours = request.form.get('hours', '').strip()
    purpose = request.form.get('purpose', '').strip()
    status = request.form.get('status', 'Pending').strip()
    notes = request.form.get('notes', '').strip()

    hours = 0.0
    if raw_hours:
        try:
            hours = float(raw_hours)
        except ValueError:
            hours = 0.0
    elif start_time and end_time:
        try:
            t1 = datetime.datetime.strptime(start_time, '%H:%M')
            t2 = datetime.datetime.strptime(end_time, '%H:%M')
            diff = (t2 - t1).total_seconds() / 3600.0
            if diff < 0:
                diff += 24.0
            hours = round(diff, 2)
        except Exception:
            hours = 0.0

    if not ot_date or not staff_name or hours <= 0:
        flash("⚠️ Vui lòng nhập đầy đủ Ngày, Tên nhân sự và Số giờ tăng ca hợp lệ!", "warning")
        return redirect(url_for('admin_panel', tab='overtime'))

    conn = get_db_connection()
    conn.execute('''
        INSERT INTO overtime_records (ot_date, staff_name, role, start_time, end_time, hours, purpose, status, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (ot_date, staff_name, role, start_time, end_time, hours, purpose, status, notes))
    conn.commit()
    conn.close()

    target_month = ot_date[:7] if len(ot_date) >= 7 else 'all'
    flash(f"✅ Đã ghi nhận lịch tăng ca cho {staff_name} ({hours:g} giờ - {ot_date}) thành công!", "success")
    return redirect(url_for('admin_panel', tab='overtime', ot_month=target_month))


@app.route('/admin/overtime/toggle-status/<int:ot_id>', methods=['POST'])
@admin_required
def toggle_overtime_status(ot_id):
    """Cập nhật trạng thái Đã làm <-> Chưa tới hạn cho ca tăng ca từ Dropdown."""
    current_month = request.args.get('ot_month', '').strip()
    conn = get_db_connection()
    row = conn.execute("SELECT status, staff_name, ot_date FROM overtime_records WHERE id = ?", (ot_id,)).fetchone()
    if row:
        requested_status = (request.form.get('status') or '').strip()
        if requested_status in ('Pending', 'Completed'):
            new_status = requested_status
        else:
            new_status = 'Completed' if row['status'] != 'Completed' else 'Pending'
        conn.execute("UPDATE overtime_records SET status = ? WHERE id = ?", (new_status, ot_id))
        conn.commit()
        st_text = "Đã làm" if new_status == 'Completed' else "Chưa tới hạn"
        flash(f"✅ Đã cập nhật trạng thái ca tăng ca của {row['staff_name']} sang '{st_text}'!", "success")
    conn.close()
    return redirect(url_for('admin_panel', tab='overtime', ot_month=current_month if current_month else None))


@app.route('/admin/overtime/delete/<int:ot_id>', methods=['POST'])
@admin_required
def delete_overtime(ot_id):
    """Xóa một bản ghi tăng ca."""
    current_month = request.args.get('ot_month', '').strip()
    conn = get_db_connection()
    conn.execute("DELETE FROM overtime_records WHERE id = ?", (ot_id,))
    conn.commit()
    conn.close()
    flash("🗑️ Đã xóa bản ghi tăng ca thành công!", "success")
    return redirect(url_for('admin_panel', tab='overtime', ot_month=current_month if current_month else None))


@app.route('/admin/overtime/export', methods=['GET'])
@admin_required
def export_overtime_excel():
    """Xuất file Excel bảng chấm công OT theo tháng."""
    month = request.args.get('month', datetime.date.today().strftime('%Y-%m')).strip()
    conn = get_db_connection()
    if month == 'all':
        rows = conn.execute("SELECT * FROM overtime_records ORDER BY ot_date ASC, id ASC").fetchall()
    else:
        rows = conn.execute("SELECT * FROM overtime_records WHERE ot_date LIKE ? ORDER BY ot_date ASC, id ASC", (f"{month}%",)).fetchall()
    conn.close()

    excel_buffer, filename = build_overtime_report_excel(rows, month)
    return send_file(
        excel_buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================================
# QUẢN LÝ NHẬT KÝ CÔNG VIỆC (WORK LOG) & CHECKLIST DEPLOY 8 PHÒNG MÁY
# ==========================================

WORKLOG_CATEGORY_LABELS = {
    'server_net': 'Server / Router / Switch',
    'vmware': 'VMware / OS / Image',
    'support': 'Hỗ trợ GV (PM / Phần cứng)',
    'booking': 'Xử lý Mượn phòng máy',
    'process': 'Cải tiến Quy trình QTTB',
    'office': 'Công việc Văn phòng khác'
}

WORKLOG_STATUS_LABELS = {
    'completed': 'Hoàn thành',
    'in_progress': 'Đang xử lý'
}

WORKLOG_PRIORITY_LABELS = {
    'urgent': 'Khẩn cấp',
    'high': 'Ưu tiên cao',
    'normal': 'Bình thường'
}

def calculate_worklog_duration(work_date_str, completed_date_str, status):
    """
    Tính số ngày xử lý công việc giữa work_date và completed_date.
    Trả về dict gồm:
      - days: số ngày (int) hoặc None
      - label: text hiển thị ('Trong ngày', '1 ngày', '3 ngày', 'Đang xử lý (2 ngày)', v.v.)
      - badge_class: class màu badge
    """
    if not work_date_str:
        return {'days': None, 'label': '—', 'badge_class': 'text-muted'}

    try:
        start_d = datetime.datetime.strptime(work_date_str[:10], '%Y-%m-%d').date()
    except Exception:
        return {'days': None, 'label': '—', 'badge_class': 'text-muted'}

    if status == 'completed':
        if completed_date_str:
            try:
                end_d = datetime.datetime.strptime(completed_date_str[:10], '%Y-%m-%d').date()
            except Exception:
                end_d = start_d
        else:
            # Nếu chưa lưu completed_date (dữ liệu cũ), mặc định lấy work_date để là xong trong ngày
            end_d = start_d

        diff = (end_d - start_d).days
        if diff <= 0:
            return {'days': 0, 'label': 'Trong ngày', 'badge_class': 'bg-success bg-opacity-10 text-success border border-success'}
        elif diff == 1:
            return {'days': 1, 'label': '1 ngày', 'badge_class': 'bg-primary bg-opacity-10 text-primary border border-primary'}
        else:
            return {'days': diff, 'label': f"{diff} ngày", 'badge_class': 'bg-info bg-opacity-15 text-dark border border-info'}
    else:
        # Đang làm (in_progress hoặc pending)
        today = datetime.date.today()
        elapsed = (today - start_d).days
        if elapsed <= 0:
            return {'days': 0, 'label': 'Hôm nay', 'badge_class': 'bg-light text-secondary border'}
        elif elapsed == 1:
            return {'days': 1, 'label': '1 ngày qua', 'badge_class': 'bg-warning bg-opacity-15 text-dark border border-warning'}
        else:
            return {'days': elapsed, 'label': f"{elapsed} ngày qua", 'badge_class': 'bg-warning bg-opacity-25 text-danger border border-warning'}


@app.route('/admin/worklog/add', methods=['POST'])
@admin_required
def add_worklog():
    """Thêm công việc mới vào Nhật ký công việc thường nhật."""
    work_date = (request.form.get('work_date') or datetime.date.today().strftime('%Y-%m-%d')).strip()
    shift_name = (request.form.get('shift_name') or 'Ca sáng').strip()
    category = (request.form.get('category') or 'server_net').strip()
    target_rooms = (request.form.get('target_rooms') or 'Tất cả 8 phòng').strip()
    content = (request.form.get('content') or '').strip()
    notes = (request.form.get('notes') or '').strip()
    priority = (request.form.get('priority') or 'normal').strip()
    status = (request.form.get('status') or 'in_progress').strip()
    created_by = get_current_user_email() or 'Admin'
    today_iso = datetime.date.today().strftime('%Y-%m-%d')
    completed_date = today_iso if status == 'completed' else None

    if not content:
        flash("⚠️ Vui lòng nhập nội dung công việc!", "warning")
        return redirect(url_for('admin_panel', tab='worklog'))

    conn = get_db_connection()
    conn.execute('''
        INSERT INTO work_logs (work_date, shift_name, category, target_rooms, content, notes, priority, status, created_by, completed_date)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (work_date, shift_name, category, target_rooms, content, notes, priority, status, created_by, completed_date))
    conn.commit()
    conn.close()

    flash("✅ Đã thêm công việc mới vào Nhật ký công việc!", "success")
    return redirect(url_for('admin_panel', tab='worklog'))


@app.route('/admin/worklog/edit/<int:log_id>', methods=['POST'])
@admin_required
def edit_worklog(log_id):
    """Chỉnh sửa nội dung và ghi chú của một công việc trong Nhật ký."""
    work_date = (request.form.get('work_date') or datetime.date.today().strftime('%Y-%m-%d')).strip()
    shift_name = (request.form.get('shift_name') or 'Ca sáng').strip()
    category = (request.form.get('category') or 'server_net').strip()
    target_rooms = (request.form.get('target_rooms') or 'Tất cả 8 phòng').strip()
    content = (request.form.get('content') or '').strip()
    notes = (request.form.get('notes') or '').strip()
    priority = (request.form.get('priority') or 'normal').strip()
    status = (request.form.get('status') or 'in_progress').strip()

    if not content:
        flash("⚠️ Nội dung công việc không được để trống!", "warning")
        return redirect(url_for('admin_panel', tab='worklog'))

    today_iso = datetime.date.today().strftime('%Y-%m-%d')
    conn = get_db_connection()
    row = conn.execute("SELECT status, completed_date FROM work_logs WHERE id = ?", (log_id,)).fetchone()
    if status == 'completed':
        completed_date = (row['completed_date'] if row and row['completed_date'] else today_iso)
    else:
        completed_date = None

    conn.execute('''
        UPDATE work_logs
        SET work_date = ?, shift_name = ?, category = ?, target_rooms = ?,
            content = ?, notes = ?, priority = ?, status = ?, completed_date = ?
        WHERE id = ?
    ''', (work_date, shift_name, category, target_rooms, content, notes, priority, status, completed_date, log_id))
    conn.commit()
    conn.close()

    flash(f"✏️ Đã cập nhật công việc #{log_id} thành công!", "success")
    return redirect(url_for('admin_panel', tab='worklog'))


@app.route('/admin/worklog/toggle-status/<int:log_id>', methods=['POST'])
@admin_required
def toggle_worklog_status(log_id):
    """Cập nhật trạng thái công việc từ Dropdown (Đang xử lý <-> Hoàn thành)."""
    conn = get_db_connection()
    row = conn.execute("SELECT status, work_date, completed_date FROM work_logs WHERE id = ?", (log_id,)).fetchone()
    if row:
        cur = row['status'] or 'in_progress'
        requested_status = (request.form.get('status') or '').strip()
        today_iso = datetime.date.today().strftime('%Y-%m-%d')

        if requested_status in ('in_progress', 'completed'):
            nxt = requested_status
        else:
            nxt = 'completed' if cur != 'completed' else 'in_progress'

        if nxt == 'completed':
            completed_date = today_iso
        else:
            completed_date = None

        conn.execute("UPDATE work_logs SET status = ?, completed_date = ? WHERE id = ?", (nxt, completed_date, log_id))
        conn.commit()

        if nxt == 'completed':
            dur = calculate_worklog_duration(row['work_date'], completed_date, 'completed')
            flash(f"✅ Đã chuyển trạng thái sang 'Hoàn thành' (Thời gian xử lý: {dur['label']})!", "success")
        else:
            flash(f"🔄 Đã chuyển trạng thái công việc sang '{WORKLOG_STATUS_LABELS.get(nxt, nxt)}'!", "success")
    conn.close()
    return redirect(url_for('admin_panel', tab='worklog'))


@app.route('/admin/worklog/delete/<int:log_id>', methods=['POST'])
@admin_required
def delete_worklog(log_id):
    """Xóa lẻ 1 bản ghi nhật ký công việc."""
    conn = get_db_connection()
    conn.execute("DELETE FROM work_logs WHERE id = ?", (log_id,))
    conn.commit()
    conn.close()
    flash("🗑️ Đã xóa bản ghi công việc thành công!", "success")
    return redirect(url_for('admin_panel', tab='worklog'))


@app.route('/admin/worklog/bulk-delete', methods=['POST'])
@admin_required
def bulk_delete_worklogs():
    """Xóa hàng loạt hoặc Chọn xóa tất cả bản ghi nhật ký công việc."""
    selected_ids = [int(i) for i in request.form.getlist('selected_ids') if str(i).isdigit()]
    if selected_ids:
        conn = get_db_connection()
        placeholders = ','.join(['?'] * len(selected_ids))
        conn.execute(f"DELETE FROM work_logs WHERE id IN ({placeholders})", selected_ids)
        conn.commit()
        conn.close()
        flash(f"🗑️ Đã xóa {len(selected_ids)} công việc được chọn khỏi Nhật ký!", "success")
    else:
        flash("⚠️ Chưa có công việc nào được chọn để xóa.", "warning")
    return redirect(url_for('admin_panel', tab='worklog'))


@app.route('/admin/worklog/export', methods=['GET'])
@admin_required
def export_worklog_excel():
    """Xuất file Excel (.xlsx) Nhật ký công việc thường nhật kèm cột Ghi chú."""
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM work_logs ORDER BY work_date DESC, id DESC").fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Nhat Ky Cong Viec"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Calibri", size=14, bold=True, color="0F2942")
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="475569")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill("solid", fgColor="0F2942")
    fill_note_header = PatternFill("solid", fgColor="0D9488")
    font_body = Font(name="Calibri", size=10, color="1F2937")
    font_bold = Font(name="Calibri", size=10, bold=True, color="1F2937")
    fill_note_cell = PatternFill("solid", fgColor="FFFBEB")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    ws.merge_cells("A1:I1")
    ws["A1"] = "SỔ NHẬT KÝ CÔNG VIỆC QUẢN TRỊ HỆ THỐNG 8 PHÒNG MÁY & VĂN PHÒNG"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:I2")
    ws["A2"] = f"Hệ thống Phòng máy (5.1, 5.2, 5.3, 6.1, 6.2, 6.3, 6.5, 6.6) • Ngày xuất: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = font_subtitle
    ws["A2"].alignment = align_center

    headers = [
        "STT", "Ngày thực hiện", "Ca làm việc", "Danh mục công việc",
        "Phòng máy / Phạm vi", "Nội dung công việc chi tiết",
        "Ghi chú (Notes)", "Mức ưu tiên", "Trạng thái", "Số ngày xử lý"
    ]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_note_header if col_idx == 7 else fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[4].height = 28

    for idx, r in enumerate(rows, 1):
        row_num = 4 + idx
        cat_label = WORKLOG_CATEGORY_LABELS.get(r['category'], r['category'])
        pri_label = WORKLOG_PRIORITY_LABELS.get(r['priority'], r['priority'])
        st_label = WORKLOG_STATUS_LABELS.get(r['status'], r['status'])
        dur_info = calculate_worklog_duration(r['work_date'], r['completed_date'], r['status'])

        row_vals = [
            (idx, align_center, font_body),
            (r['work_date'], align_center, font_bold),
            (r['shift_name'], align_center, font_body),
            (cat_label, align_left, font_bold),
            (r['target_rooms'], align_center, font_bold),
            (r['content'], align_left, font_body),
            (r['notes'] or '', align_left, font_body),
            (pri_label, align_center, font_body),
            (st_label, align_center, font_bold),
            (dur_info['label'], align_center, font_body)
        ]
        for c_idx, (val, al, fn) in enumerate(row_vals, 1):
            c = ws.cell(row=row_num, column=c_idx, value=val)
            c.alignment = al
            c.font = fn
            c.border = thin_border
            if c_idx == 7 and val:
                c.fill = fill_note_cell
        ws.row_dimensions[row_num].height = 26

    col_widths = [6, 14, 13, 28, 20, 44, 38, 14, 16, 16]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"Nhat_Ky_Cong_Viec_Phong_May_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route('/api/deploy-checklist/update', methods=['POST'])
@admin_required
def api_update_deploy_checklist():
    """API lưu tức thì tiến độ tích chọn 23 bước SOP Deploy và Ghi chú của từng phòng máy."""
    data = request.get_json(silent=True) or {}
    campaign_name = (data.get('campaign_name') or 'HK1 (2026-2027) - Deploy Hệ thống chuẩn').strip()
    room = (data.get('room') or '5.1').strip()
    checked_steps = data.get('checked_steps', [])
    if not isinstance(checked_steps, list):
        checked_steps = []
    notes = (data.get('notes') or '').strip()
    updated_by = get_current_user_email() or 'Admin'
    now_short = datetime.datetime.now().strftime('%H:%M %d/%m/%Y')

    conn = get_db_connection()
    conn.execute('''
        INSERT INTO deployment_checklists (campaign_name, room, checked_steps, notes, updated_by, updated_at)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(campaign_name, room) DO UPDATE SET
            checked_steps = excluded.checked_steps,
            notes = excluded.notes,
            updated_by = excluded.updated_by,
            updated_at = excluded.updated_at
    ''', (campaign_name, room, json.dumps(checked_steps), notes, updated_by, now_short))
    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'room': room,
        'checked_count': len(checked_steps),
        'updated_at': now_short,
        'updated_by': updated_by
    })


@app.route('/api/deploy-checklist/copy-phase-a-all', methods=['POST'])
@admin_required
def api_copy_phase_a_all():
    """API đánh dấu hoàn tất Phần A (11 bước Master Image) cho toàn bộ 8 phòng trong đợt triển khai."""
    data = request.get_json(silent=True) or {}
    campaign_name = (data.get('campaign_name') or 'HK1 (2026-2027) - Deploy Hệ thống chuẩn').strip()
    phase_a_ids = ['A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10']
    updated_by = get_current_user_email() or 'Admin'
    now_short = datetime.datetime.now().strftime('%H:%M %d/%m/%Y')

    conn = get_db_connection()
    updated_map = {}

    for r_code in ROOMS_LIST:
        row = conn.execute(
            "SELECT checked_steps, notes FROM deployment_checklists WHERE campaign_name = ? AND room = ?",
            (campaign_name, r_code)
        ).fetchone()
        existing = []
        cur_note = ''
        if row:
            cur_note = row['notes'] or ''
            try:
                existing = json.loads(row['checked_steps'] or '[]')
            except Exception:
                existing = []
        merged = list(dict.fromkeys(existing + phase_a_ids))
        conn.execute('''
            INSERT INTO deployment_checklists (campaign_name, room, checked_steps, notes, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(campaign_name, room) DO UPDATE SET
                checked_steps = excluded.checked_steps,
                updated_by = excluded.updated_by,
                updated_at = excluded.updated_at
        ''', (campaign_name, r_code, json.dumps(merged), cur_note, updated_by, now_short))
        updated_map[r_code] = {
            'room': r_code,
            'checked': merged,
            'notes': cur_note,
            'updated_by': updated_by,
            'updated_at': now_short
        }

    conn.commit()
    conn.close()
    return jsonify({'success': True, 'room_map': updated_map, 'updated_at': now_short})


@app.route('/admin/deploy-checklist/campaign/add', methods=['POST'])
@admin_required
def add_deploy_campaign():
    """Khởi tạo một đợt triển khai (Deploy Campaign) mới cho 8 phòng máy."""
    campaign_name = (request.form.get('campaign_name') or '').strip()
    if not campaign_name:
        flash("⚠️ Vui lòng nhập tên đợt triển khai mới!", "warning")
        return redirect(url_for('admin_panel', tab='worklog', wl_view='deploy'))

    now_short = datetime.datetime.now().strftime('%H:%M %d/%m/%Y')
    updated_by = get_current_user_email() or 'Admin'
    conn = get_db_connection()
    for r_code in ROOMS_LIST:
        conn.execute('''
            INSERT OR IGNORE INTO deployment_checklists (campaign_name, room, checked_steps, notes, updated_by, updated_at)
            VALUES (?, ?, '[]', '', ?, ?)
        ''', (campaign_name, r_code, updated_by, now_short))
    conn.commit()
    conn.close()

    flash(f"🚀 Đã khởi tạo đợt triển khai '{campaign_name}' cho 8 phòng máy!", "success")
    return redirect(url_for('admin_panel', tab='worklog', wl_view='deploy', deploy_campaign=campaign_name))


@app.route('/admin/deploy-checklist/campaign/delete', methods=['POST'])
@admin_required
def delete_deploy_campaign():
    """Xóa một đợt triển khai (Deploy Campaign) khỏi hệ thống."""
    campaign_name = (request.form.get('campaign_name') or '').strip()
    if campaign_name:
        conn = get_db_connection()
        conn.execute("DELETE FROM deployment_checklists WHERE campaign_name = ?", (campaign_name,))
        conn.commit()
        conn.close()
        flash(f"🗑️ Đã xóa đợt triển khai '{campaign_name}'!", "success")
    return redirect(url_for('admin_panel', tab='worklog', wl_view='deploy'))


@app.route('/api/deploy-checklist/log-to-worklog', methods=['POST'])
@admin_required
def api_log_deploy_to_worklog():
    """Ghi nhận nhanh tiến độ Deploy của phòng đang chọn vào Bảng Nhật ký công việc hằng ngày."""
    data = request.get_json(silent=True) or {}
    campaign_name = (data.get('campaign_name') or '').strip()
    room = (data.get('room') or '5.1').strip()
    done_count = int(data.get('done_count') or 0)
    total_count = int(data.get('total_count') or 23)
    notes = (data.get('notes') or '').strip()

    pct = int(round((done_count / max(1, total_count)) * 100))
    status = 'completed' if done_count >= total_count else 'in_progress'
    today_iso = datetime.date.today().strftime('%Y-%m-%d')
    created_by = get_current_user_email() or 'Admin'

    content = (
        f"[Deploy Hệ Thống - {campaign_name}] Triển khai Acronis Master Image & Clonezilla Lite Server "
        f"Phòng {room}: Đạt {done_count}/{total_count} bước ({pct}%)"
    )
    final_note = notes if notes else f"Thực hiện theo SOP 23 bước (Acronis Recovery + Clonezilla Lite Server + Cấu hình máy GV)"

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO work_logs (work_date, shift_name, category, target_rooms, content, notes, priority, status, created_by)
        VALUES (?, 'Cả ngày', 'vmware', ?, ?, ?, 'high', ?, ?)
    ''', (today_iso, room, content, final_note, status, created_by))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'log_id': new_id,
        'message': f"Đã ghi kết quả Deploy Phòng {room} ({done_count}/{total_count} bước) vào Nhật ký công việc!"
    })


@app.route('/admin/email-config', methods=['POST'])
@admin_required
def save_email_config():
    """Lưu cấu hình gửi Email tự động (Gmail SMTP App Password)."""
    smtp_enabled = '1' if request.form.get('smtp_enabled') == '1' else '0'
    smtp_email = (request.form.get('smtp_email') or 'mlhang@hcmus.edu.vn').strip()
    smtp_password = (request.form.get('smtp_password') or '').strip()
    smtp_sender_name = (request.form.get('smtp_sender_name') or 'Hệ thống Phòng máy B4.2 - Phòng QTTB').strip()
    return_tab = (request.form.get('return_tab') or 'requests').strip()

    conn = get_db_connection()
    conn.execute("INSERT OR REPLACE INTO system_settings (key, value) VALUES ('smtp_enabled', ?)", (smtp_enabled,))
    conn.execute("INSERT OR REPLACE INTO system_settings (key, value) VALUES ('smtp_email', ?)", (smtp_email,))
    conn.execute("INSERT OR REPLACE INTO system_settings (key, value) VALUES ('smtp_sender_name', ?)", (smtp_sender_name,))
    if smtp_password:
        conn.execute("INSERT OR REPLACE INTO system_settings (key, value) VALUES ('smtp_password', ?)", (smtp_password,))
    conn.commit()
    conn.close()

    flash("✅ Đã lưu cấu hình Gửi Email tự động trả kết quả cho User!", "success")
    return redirect(url_for('admin_panel', tab=return_tab))


@app.route('/admin/email-test', methods=['POST'])
@admin_required
def test_email_config():
    """API kiểm tra kết nối Gmail SMTP và gửi 1 email test tới hộp thư chỉ định."""
    data = request.get_json(silent=True) or {}
    test_to = (data.get('to_email') or get_current_user_email() or 'mlhang@hcmus.edu.vn').strip()
    smtp_email = (data.get('smtp_email') or '').strip()
    smtp_password = (data.get('smtp_password') or '').strip()
    smtp_sender_name = (data.get('smtp_sender_name') or 'Hệ thống Phòng máy B4.2 - Phòng QTTB').strip()

    conn = get_db_connection()
    saved_cfg = get_email_settings(conn=conn)
    conn.close()

    effective_cfg = {
        'enabled': True,
        'email': smtp_email or saved_cfg['email'],
        'password': smtp_password or saved_cfg['password'],
        'sender_name': smtp_sender_name or saved_cfg['sender_name']
    }

    sample_rec = {
        'requester_name': get_current_user_name() or 'Quản trị viên QTTB',
        'department': 'Phòng Quản trị Thiết bị',
        'room': '5.1',
        'day_of_week': 2,
        'start_period': 1,
        'end_period': 5,
        'subject_name': 'Kiểm tra hệ thống gửi Email tự động',
        'class_code': 'TEST-B42',
        'purpose': 'Kiểm tra kết nối SMTP trả kết quả tự động cho giảng viên',
        'study_period': datetime.date.today().strftime('%d/%m/%Y'),
        'email': test_to
    }
    subject, html_body, text_body = build_request_notification_email(
        req_type='booking',
        record=sample_rec,
        new_status='Processing',
        admin_note='Đây là email kiểm tra kết nối từ Hệ thống Quản lý Phòng máy B4.2. Cấu hình SMTP hoạt động tốt!',
        dashboard_url=url_for('my_dashboard', _external=True)
    )
    ok, msg = send_smtp_email_sync(test_to, subject, html_body, text_body, email_cfg=effective_cfg)
    return jsonify({'success': ok, 'message': msg})


@app.route('/admin/update-teacher-passcode', methods=['POST'])
@admin_required
def update_teacher_passcode():
    """Cập nhật Mã Passcode nội bộ dành cho Thầy/Cô dùng Gmail cá nhân."""
    new_passcode = (request.form.get('teacher_passcode') or '').strip()
    return_tab = (request.form.get('return_tab') or 'dashboard').strip()
    if not new_passcode:
        flash("Mã Passcode không được để trống.", "warning")
        return redirect(url_for('admin_panel', tab=return_tab))

    conn = get_db_connection()
    conn.execute("INSERT OR REPLACE INTO system_settings (key, value) VALUES ('teacher_passcode', ?)", (new_passcode,))
    conn.commit()
    conn.close()
    flash(f"Đã cập nhật Mã Passcode Giảng viên thành công: {new_passcode}", "success")
    return redirect(url_for('admin_panel', tab=return_tab))


@app.route('/admin/delete-verified-user/<int:user_id>', methods=['POST'])
@admin_required
def delete_verified_user(user_id):
    """Thu hồi quyền Giảng viên của một tài khoản Gmail."""
    return_tab = (request.form.get('return_tab') or 'dashboard').strip()
    conn = get_db_connection()
    row = conn.execute("SELECT email FROM verified_users WHERE id = ?", (user_id,)).fetchone()
    if row:
        conn.execute("DELETE FROM verified_users WHERE id = ?", (user_id,))
        conn.commit()
        flash(f"Đã thu hồi xác thực Giảng viên của tài khoản {row['email']}.", "info")
    conn.close()
    return redirect(url_for('admin_panel', tab=return_tab))


@app.route('/admin/add-verified-user', methods=['POST'])
@admin_required
def add_verified_user():
    """Admin chủ động thêm xác thực cho tài khoản Giảng viên dùng Gmail."""
    email = (request.form.get('email') or '').strip().lower()
    full_name = (request.form.get('full_name') or '').strip()
    return_tab = (request.form.get('return_tab') or 'dashboard').strip()
    if not email or '@' not in email:
        flash("Địa chỉ email không hợp lệ.", "warning")
        return redirect(url_for('admin_panel', tab=return_tab))

    conn = get_db_connection()
    try:
        conn.execute(
            "INSERT OR REPLACE INTO verified_users (email, full_name, verified_by) VALUES (?, ?, 'admin')",
            (email, full_name or email.split('@')[0])
        )
        conn.commit()
        flash(f"Đã cấp quyền Giảng viên thành công cho tài khoản {email}.", "success")
    except Exception as e:
        flash(f"Lỗi: {e}", "danger")
    finally:
        conn.close()
    return redirect(url_for('admin_panel', tab=return_tab))


@app.route('/admin/block-user', methods=['POST'])
@admin_required
def block_user_account():
    """Khóa tài khoản vĩnh viễn (1-Click Block User) trên trang Admin."""
    email = (request.form.get('email') or '').strip().lower()
    reason = (request.form.get('reason') or 'Biểu hiện bất thường / Spam hệ thống').strip()
    return_tab = (request.form.get('return_tab') or 'requests').strip()

    if not email or '@' not in email:
        flash("Địa chỉ email cần khóa không hợp lệ.", "warning")
        return redirect(url_for('admin_panel', tab=return_tab))

    if email in [e.lower() for e in ADMIN_EMAILS]:
        flash("Không thể khóa tài khoản Quản trị viên (Admin)!", "danger")
        return redirect(url_for('admin_panel', tab=return_tab))

    admin_email = get_current_user_email() or 'Admin'
    conn = get_db_connection()
    try:
        conn.execute(
            "INSERT OR REPLACE INTO blocked_users (email, reason, blocked_by) VALUES (?, ?, ?)",
            (email, reason, admin_email)
        )
        # Đồng thời thu hồi quyền xác minh Passcode nếu có
        conn.execute("DELETE FROM verified_users WHERE LOWER(email) = LOWER(?)", (email,))
        conn.commit()
        flash(f"🚫 Đã khóa vĩnh viễn tài khoản {email} thành công!", "danger")
    except Exception as e:
        flash(f"Lỗi khi khóa tài khoản: {e}", "danger")
    finally:
        conn.close()

    return redirect(url_for('admin_panel', tab=return_tab))


@app.route('/admin/unblock-user/<int:block_id>', methods=['POST'])
@admin_required
def unblock_user_account(block_id):
    """Mở khóa tài khoản đã bị chặn."""
    return_tab = (request.form.get('return_tab') or 'requests').strip()
    conn = get_db_connection()
    row = conn.execute("SELECT email FROM blocked_users WHERE id = ?", (block_id,)).fetchone()
    if row:
        conn.execute("DELETE FROM blocked_users WHERE id = ?", (block_id,))
        conn.commit()
        flash(f"🔓 Đã mở khóa tài khoản {row['email']} thành công.", "success")
    conn.close()
    return redirect(url_for('admin_panel', tab=return_tab))


@app.route('/admin/approve-cancellation/<int:cancellation_id>', methods=['POST'])
@admin_required
def approve_cancellation(cancellation_id):
    """Cập nhật trạng thái thông báo nghỉ gấp (Chờ xử lý / Đã tiếp nhận / Đã duyệt) từ Dropdown và gửi email thông báo cho giảng viên."""
    new_status = request.form.get('status', 'Approved')
    if new_status not in ('Pending', 'Processing', 'Approved'):
        new_status = 'Approved'
    admin_note = request.form.get('admin_note', '').strip()

    conn = get_db_connection()
    rec = conn.execute("SELECT * FROM cancellations WHERE id = ?", (cancellation_id,)).fetchone()
    email_suffix = ""
    if rec:
        final_note = admin_note if admin_note else (rec['admin_note'] or '')
        conn.execute(
            "UPDATE cancellations SET status = ?, admin_note = ? WHERE id = ?",
            (new_status, final_note, cancellation_id)
        )
        conn.commit()
        if new_status in ('Processing', 'Approved'):
            email_suffix = notify_user_on_status_change('cancellation', rec, new_status, admin_note=final_note, conn=conn)
    conn.close()

    status_label = "Đã tiếp nhận" if new_status == 'Processing' else ("Đã duyệt" if new_status == 'Approved' else "Chờ xử lý")
    flash(f"Đã cập nhật trạng thái báo nghỉ: {status_label}.{email_suffix}", "success")
    return redirect(url_for('admin_panel', tab='cancellations'))

@app.route('/admin/delete-cancellation/<int:cancellation_id>', methods=['POST'])
@admin_required
def delete_cancellation(cancellation_id):
    """Xóa một thông báo nghỉ gấp."""
    conn = get_db_connection()
    conn.execute("DELETE FROM cancellations WHERE id = ?", (cancellation_id,))
    conn.commit()
    conn.close()
    flash("Đã xóa thông báo nghỉ thành công.", "success")
    return redirect(url_for('admin_panel', tab='cancellations'))

@app.route('/admin/bulk-delete-cancellations', methods=['POST'])
@admin_required
def bulk_delete_cancellations():
    """Xóa hàng loạt thông báo nghỉ gấp đã chọn."""
    selected_ids = [int(i) for i in request.form.getlist('selected_ids') if str(i).isdigit()]
    if selected_ids:
        conn = get_db_connection()
        placeholders = ','.join(['?'] * len(selected_ids))
        conn.execute(f"DELETE FROM cancellations WHERE id IN ({placeholders})", selected_ids)
        conn.commit()
        conn.close()
        flash(f"Đã xóa {len(selected_ids)} thông báo nghỉ được chọn.", "success")
    else:
        flash("Chưa có mục nào được chọn để xóa.", "warning")
    return redirect(url_for('admin_panel', tab='cancellations'))

@app.route('/admin/update-request-status/<int:request_id>', methods=['POST'])
@admin_required
def update_request_status(request_id):
    """Cập nhật trạng thái yêu cầu mượn phòng (Chờ xử lý / Đã tiếp nhận / Đã duyệt / Từ chối) từ Dropdown và gửi email."""
    new_status = (request.form.get('status') or 'Processing').strip()
    if new_status not in ('Pending', 'Processing', 'Approved', 'Rejected'):
        new_status = 'Processing'
    admin_note = request.form.get('admin_note', '').strip()

    conn = get_db_connection()
    req = conn.execute("SELECT * FROM booking_requests WHERE id = ?", (request_id,)).fetchone()
    email_suffix = ""
    if req:
        if new_status == 'Approved' and req['status'] != 'Approved':
            class_code = req['class_code'] if req['class_code'] else 'MƯỢN PHÒNG'
            subject_name = req['subject_name'] if req['subject_name'] else req['purpose']
            department = req['department'] if req['department'] else req['requester_name']
            study_period = normalize_study_period_str(req['study_period']) if req['study_period'] else datetime.date.today().strftime('%d/%m/%Y')

            conn.execute('''
                INSERT INTO schedules (room, day_of_week, start_period, end_period, class_code, subject_name, department, study_period, num_weeks, is_ended_early)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, 0)
            ''', (req['room'], req['day_of_week'], req['start_period'], req['end_period'], class_code, subject_name, department, study_period))

        final_note = admin_note if admin_note else (req['admin_note'] or '')
        conn.execute(
            "UPDATE booking_requests SET status = ?, admin_note = ? WHERE id = ?",
            (new_status, final_note, request_id)
        )
        conn.commit()
        if new_status in ('Processing', 'Approved', 'Rejected'):
            email_suffix = notify_user_on_status_change('booking', req, new_status, admin_note=final_note, conn=conn)

    conn.close()
    status_map = {
        'Pending': ('Chờ xử lý', 'secondary'),
        'Processing': ('Đã tiếp nhận', 'info'),
        'Approved': ('Đã duyệt', 'success'),
        'Rejected': ('Từ chối', 'warning')
    }
    st_label, flash_cat = status_map.get(new_status, ('Đã cập nhật', 'info'))
    flash(f"Đã cập nhật trạng thái yêu cầu mượn phòng sang '{st_label}'.{email_suffix}", flash_cat)
    return redirect(url_for('admin_panel', tab='requests'))

@app.route('/admin/delete-request/<int:request_id>', methods=['POST'])
@admin_required
def delete_request(request_id):
    """Xóa một yêu cầu mượn phòng khỏi danh sách."""
    conn = get_db_connection()
    conn.execute("DELETE FROM booking_requests WHERE id = ?", (request_id,))
    conn.commit()
    conn.close()
    flash("Đã xóa yêu cầu mượn phòng thành công.", "success")
    return redirect(url_for('admin_panel', tab='requests'))

@app.route('/admin/bulk-delete-requests', methods=['POST'])
@admin_required
def bulk_delete_requests():
    """Xóa hàng loạt yêu cầu mượn phòng đã chọn."""
    selected_ids = [int(i) for i in request.form.getlist('selected_ids') if str(i).isdigit()]
    if selected_ids:
        conn = get_db_connection()
        placeholders = ','.join(['?'] * len(selected_ids))
        conn.execute(f"DELETE FROM booking_requests WHERE id IN ({placeholders})", selected_ids)
        conn.commit()
        conn.close()
        flash(f"Đã xóa {len(selected_ids)} yêu cầu mượn phòng được chọn.", "success")
    else:
        flash("Chưa có mục nào được chọn để xóa.", "warning")
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
    schedule_filter = request.form.get('schedule_filter', 'all')
    conn = get_db_connection()
    row = conn.execute("SELECT subject_name, class_code, room FROM schedules WHERE id = ?", (schedule_id,)).fetchone()
    conn.execute("DELETE FROM schedules WHERE id = ?", (schedule_id,))
    conn.commit()
    conn.close()
    
    if row:
        cls_info = f" - Lớp {row['class_code']}" if row['class_code'] else ""
        flash(f"🗑️ Đã xóa môn '{row['subject_name']}'{cls_info} (P.{row['room']}) khỏi Lịch học.", "success")
    else:
        flash("Đã xóa lịch học thành công.", "success")
    return redirect(url_for('admin_panel', tab='schedules', schedule_filter=schedule_filter))


@app.route('/admin/bulk-delete-schedules', methods=['POST'])
@admin_required
def bulk_delete_schedules():
    """Xóa hàng loạt các môn học được chọn (ví dụ khi Giáo vụ báo hủy môn)."""
    schedule_filter = request.form.get('schedule_filter', 'all')
    selected_ids = [int(i) for i in request.form.getlist('selected_ids') if str(i).isdigit()]
    if selected_ids:
        conn = get_db_connection()
        placeholders = ','.join(['?'] * len(selected_ids))
        conn.execute(f"DELETE FROM schedules WHERE id IN ({placeholders})", selected_ids)
        conn.commit()
        conn.close()
        flash(f"🗑️ Đã xóa {len(selected_ids)} lịch môn học được chọn khỏi Thời khóa biểu!", "success")
    else:
        flash("⚠️ Chưa có môn học nào được chọn để xóa.", "warning")
    return redirect(url_for('admin_panel', tab='schedules', schedule_filter=schedule_filter))

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
    new_status = (request.form.get('status') or 'Pending').strip()
    if new_status == 'Approved':
        new_status = 'Resolved'
    admin_note = request.form.get('admin_note', '').strip()
    if new_status in ('Pending', 'Processing', 'Resolved'):
        conn = get_db_connection()
        iss = conn.execute("SELECT * FROM issue_reports WHERE id = ?", (issue_id,)).fetchone()
        email_suffix = ""
        if iss:
            final_note = admin_note if admin_note else (iss['admin_note'] or '')
            conn.execute('UPDATE issue_reports SET status = ?, admin_note = ? WHERE id = ?', (new_status, final_note, issue_id))
            conn.commit()
            if new_status in ('Processing', 'Resolved'):
                email_suffix = notify_user_on_status_change('issue', iss, new_status, admin_note=final_note, conn=conn)
        conn.close()
        status_label = "Đã tiếp nhận" if new_status == 'Processing' else ("Đã duyệt" if new_status == 'Resolved' else "Chờ xử lý")
        flash(f"Đã cập nhật trạng thái sự cố: {status_label}.{email_suffix}", "success")
    
    return redirect(url_for('admin_panel', tab='issues'))

@app.route('/admin/delete-issue/<int:issue_id>', methods=['POST'])
@admin_required
def delete_issue(issue_id):
    """Xóa một báo cáo sự cố khỏi danh sách."""
    conn = get_db_connection()
    conn.execute("DELETE FROM issue_reports WHERE id = ?", (issue_id,))
    conn.commit()
    conn.close()
    flash("Đã xóa báo cáo sự cố thành công.", "success")
    return redirect(url_for('admin_panel', tab='issues'))

@app.route('/admin/bulk-delete-issues', methods=['POST'])
@admin_required
def bulk_delete_issues():
    """Xóa hàng loạt báo cáo sự cố đã chọn."""
    selected_ids = [int(i) for i in request.form.getlist('selected_ids') if str(i).isdigit()]
    if selected_ids:
        conn = get_db_connection()
        placeholders = ','.join(['?'] * len(selected_ids))
        conn.execute(f"DELETE FROM issue_reports WHERE id IN ({placeholders})", selected_ids)
        conn.commit()
        conn.close()
        flash(f"Đã xóa {len(selected_ids)} báo cáo sự cố được chọn.", "success")
    else:
        flash("Chưa có mục nào được chọn để xóa.", "warning")
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
    import sys
    is_pythonw = sys.executable.lower().endswith('pythonw.exe') or sys.stdout is None
    if is_pythonw:
        sys.stdout = open(os.devnull, 'w', encoding='utf-8')
        sys.stderr = open(os.devnull, 'w', encoding='utf-8')
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
    app.run(debug=debug_mode, use_reloader=(debug_mode and not is_pythonw), host='127.0.0.1', port=5000)
