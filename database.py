import sqlite3
import os
import base64
import hashlib
import hmac
import secrets
from timezone_utils import get_vn_now_str, get_vn_today_iso

# Khai báo đường dẫn tuyệt đối tới file database (tương thích cả Windows và PythonAnywhere Linux)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, 'lab_schedule.db')
ENV_FILE_PATH = os.path.join(BASE_DIR, '.env')
_SECRET_PREFIX = 'enc:v1:'


def get_or_create_env_secret(env_var):
    """Lấy hoặc tự sinh khóa bí mật ngẫu nhiên trong biến môi trường / file .env (đã nằm trong .gitignore)."""
    key_str = (os.environ.get(env_var) or '').strip()
    if not key_str and os.path.exists(ENV_FILE_PATH):
        try:
            with open(ENV_FILE_PATH, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#') or '=' not in line:
                        continue
                    k, v = line.split('=', 1)
                    if k.strip() == env_var and v.strip():
                        key_str = v.strip().strip('"').strip("'")
                        os.environ[env_var] = key_str
                        break
        except Exception:
            pass

    if not key_str:
        key_str = secrets.token_urlsafe(32)
        os.environ[env_var] = key_str
        try:
            need_nl = os.path.exists(ENV_FILE_PATH) and os.path.getsize(ENV_FILE_PATH) > 0
            with open(ENV_FILE_PATH, 'a', encoding='utf-8') as f:
                if need_nl:
                    f.write('\n')
                f.write(f'{env_var}={key_str}\n')
        except Exception:
            pass

    return key_str


def _get_router_master_key():
    """Lấy hoặc tự sinh khóa đối xứng 32-byte từ biến môi trường / file .env (đã nằm trong .gitignore)."""
    key_str = get_or_create_env_secret('ROUTER_SECRET_KEY')
    return hashlib.sha256(key_str.encode('utf-8')).digest()


def _keystream_xor(enc_key, nonce, data_bytes):
    """Sinh keystream SHA-256 CTR và XOR với dữ liệu."""
    out = bytearray()
    counter = 0
    for idx in range(0, len(data_bytes), 32):
        block_key = hashlib.sha256(enc_key + nonce + counter.to_bytes(4, 'big')).digest()
        chunk = data_bytes[idx:idx + 32]
        for j, b in enumerate(chunk):
            out.append(b ^ block_key[j])
        counter += 1
    return bytes(out)


def encrypt_router_secret(plaintext):
    """Mã hóa mật khẩu router (Encrypt-then-MAC với SHA-256 + HMAC-SHA256) trước khi lưu SQLite."""
    text = str(plaintext or '')
    if not text:
        return ''
    if text.startswith(_SECRET_PREFIX):
        return text

    master = _get_router_master_key()
    nonce = secrets.token_bytes(16)
    k_enc = hmac.new(master, b'enc:' + nonce, hashlib.sha256).digest()
    k_mac = hmac.new(master, b'mac:' + nonce, hashlib.sha256).digest()

    cipher = _keystream_xor(k_enc, nonce, text.encode('utf-8'))
    tag = hmac.new(k_mac, nonce + cipher, hashlib.sha256).digest()
    return (
        _SECRET_PREFIX
        + base64.urlsafe_b64encode(nonce).decode('ascii') + ':'
        + base64.urlsafe_b64encode(cipher).decode('ascii') + ':'
        + base64.urlsafe_b64encode(tag).decode('ascii')
    )


def decrypt_router_secret(stored_val):
    """Giải mã mật khẩu router đã lưu trong SQLite (hỗ trợ tương thích ngược nếu chưa mã hóa)."""
    text = str(stored_val or '')
    if not text:
        return ''
    if not text.startswith(_SECRET_PREFIX):
        return text

    try:
        parts = text[len(_SECRET_PREFIX):].split(':')
        if len(parts) != 3:
            return ''
        nonce = base64.urlsafe_b64decode(parts[0].encode('ascii'))
        cipher = base64.urlsafe_b64decode(parts[1].encode('ascii'))
        tag = base64.urlsafe_b64decode(parts[2].encode('ascii'))

        master = _get_router_master_key()
        k_enc = hmac.new(master, b'enc:' + nonce, hashlib.sha256).digest()
        k_mac = hmac.new(master, b'mac:' + nonce, hashlib.sha256).digest()

        expected_tag = hmac.new(k_mac, nonce + cipher, hashlib.sha256).digest()
        if not hmac.compare_digest(tag, expected_tag):
            return ''
        plain = _keystream_xor(k_enc, nonce, cipher)
        return plain.decode('utf-8')
    except Exception:
        return ''

def get_db_connection():
    """Hàm tạo kết nối tới database (bật WAL mode + timeout 15s để chống lỗi database is locked)."""
    conn = sqlite3.connect(DB_NAME, timeout=15.0)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        conn.execute("PRAGMA busy_timeout=15000;")
    except Exception:
        pass
    return conn

def init_db():
    """Hàm khởi tạo cơ sở dữ liệu và các bảng cần thiết."""
    # Kiểm tra xem file database đã tồn tại chưa
    db_exists = os.path.exists(DB_NAME)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Bảng lưu Thời khóa biểu chính thức
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            room TEXT NOT NULL,
            day_of_week INTEGER NOT NULL,
            start_period REAL NOT NULL,
            end_period REAL NOT NULL,
            class_code TEXT,
            subject_name TEXT NOT NULL,
            department TEXT,
            study_period TEXT
        )
    ''')
    
    # 2. Bảng lưu Yêu cầu mượn phòng
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS booking_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            requester_name TEXT NOT NULL,
            room TEXT NOT NULL,
            day_of_week INTEGER NOT NULL,
            start_period REAL NOT NULL,
            end_period REAL NOT NULL,
            purpose TEXT,
            subject_code TEXT,
            class_code TEXT,
            subject_name TEXT,
            department TEXT,
            study_period TEXT,
            software_required TEXT,
            status TEXT DEFAULT 'Processing',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 3. Bảng lưu Thông báo nghỉ gấp
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cancellations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_name TEXT NOT NULL,
            department TEXT DEFAULT '',
            room TEXT NOT NULL,
            cancel_date TEXT NOT NULL,
            period TEXT NOT NULL,
            reason TEXT,
            status TEXT DEFAULT 'Processing',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 4. Bảng báo cáo sự cố (Ticketing)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS issue_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            reporter_name TEXT NOT NULL,
            department TEXT NOT NULL,
            room TEXT NOT NULL,
            machine_number TEXT NOT NULL,
            issue_category TEXT NOT NULL,
            description TEXT NOT NULL,
            status TEXT DEFAULT 'Processing',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Add email & admin_note columns if they don't exist
    try:
        cursor.execute("ALTER TABLE booking_requests ADD COLUMN email TEXT")
    except Exception:
        pass
        
    try:
        cursor.execute("ALTER TABLE issue_reports ADD COLUMN email TEXT")
    except Exception:
        pass
        
    try:
        cursor.execute("ALTER TABLE cancellations ADD COLUMN email TEXT")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE cancellations ADD COLUMN department TEXT DEFAULT ''")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE booking_requests ADD COLUMN admin_note TEXT")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE issue_reports ADD COLUMN admin_note TEXT")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE cancellations ADD COLUMN admin_note TEXT")
    except Exception:
        pass

    # Migration an toàn: Thêm cột created_at cho DB cũ nếu chưa có
    try:
        cursor.execute("ALTER TABLE booking_requests ADD COLUMN created_at TIMESTAMP")
        cursor.execute("UPDATE booking_requests SET created_at = datetime('now', '+7 hours') WHERE created_at IS NULL")
    except Exception:
        pass

    # Migration đồng bộ Múi giờ Việt Nam (UTC+7): Khắc phục các bản ghi cũ bị SQLite lưu ngầm theo UTC
    try:
        cursor.execute('''
            UPDATE cancellations
            SET created_at = datetime(created_at, '+7 hours')
            WHERE created_at LIKE '2026-11-12 04:%' OR created_at LIKE '2026-11-12 05:%'
        ''')
    except Exception:
        pass

    try:
        cursor.execute('''
            UPDATE issue_reports
            SET created_at = datetime(created_at, '+7 hours')
            WHERE created_at LIKE '2026-11-12 04:%' OR created_at LIKE '2026-11-12 05:%'
        ''')
    except Exception:
        pass

    try:
        cursor.execute('''
            UPDATE admin_permissions
            SET created_at = datetime(created_at, '+7 hours')
            WHERE created_at LIKE '2026-11-12 05:%'
        ''')
    except Exception:
        pass

    # Migration an toàn cho bảng schedules: num_weeks, is_ended_early, ended_early_date
    try:
        cursor.execute("ALTER TABLE schedules ADD COLUMN num_weeks INTEGER DEFAULT 15")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE schedules ADD COLUMN is_ended_early INTEGER DEFAULT 0")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE schedules ADD COLUMN ended_early_date TEXT DEFAULT NULL")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE schedules ADD COLUMN notes TEXT DEFAULT ''")
    except Exception:
        pass

    cursor.execute("UPDATE schedules SET num_weeks = 15 WHERE num_weeks IS NULL")
    cursor.execute("UPDATE schedules SET is_ended_early = 0 WHERE is_ended_early IS NULL")
    cursor.execute("UPDATE schedules SET notes = '' WHERE notes IS NULL")
    cursor.execute("UPDATE schedules SET study_period = '28/09/2026' WHERE study_period LIKE '2026-09-28%'")
    cursor.execute("UPDATE booking_requests SET status = 'Processing' WHERE status = 'Pending' OR status IS NULL")
    cursor.execute("UPDATE issue_reports SET status = 'Processing' WHERE status = 'Pending' OR status IS NULL")
    cursor.execute("UPDATE cancellations SET status = 'Processing' WHERE status = 'Pending' OR status IS NULL")

    # 6. Bảng lưu cấu hình hệ thống & Trạng thái mạng 8 phòng máy (B42Lab NetAgent)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS system_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS room_network_status (
            room TEXT PRIMARY KEY,
            target_ip TEXT DEFAULT '8.8.8.8',
            status TEXT DEFAULT 'online',
            ping_ms INTEGER DEFAULT 12,
            updated_at TEXT,
            router_model TEXT DEFAULT 'GWN7062G',
            https_port INTEGER DEFAULT 443,
            lan_ip TEXT DEFAULT '',
            room_state TEXT DEFAULT 'wan_ping_only',
            room_state_label TEXT DEFAULT '',
            online_clients INTEGER DEFAULT 0,
            arp_count INTEGER DEFAULT 0,
            rx_rate_bps INTEGER DEFAULT 0,
            tx_rate_bps INTEGER DEFAULT 0,
            total_rx_bytes INTEGER DEFAULT 0,
            total_tx_bytes INTEGER DEFAULT 0,
            cpu_usage INTEGER DEFAULT 0,
            mem_usage INTEGER DEFAULT 0,
            temp_c INTEGER DEFAULT 0,
            telemetry_json TEXT DEFAULT '{}'
        )
    ''')

    # [#QTTB-03] Migration an toàn cho bảng room_network_status (Grandstream GCC6010 & GWN7062G)
    router_extra_cols = [
        ("router_model", "TEXT DEFAULT 'GWN7062G'"),
        ("https_port", "INTEGER DEFAULT 443"),
        ("lan_ip", "TEXT DEFAULT ''"),
        ("room_state", "TEXT DEFAULT 'wan_ping_only'"),
        ("room_state_label", "TEXT DEFAULT ''"),
        ("online_clients", "INTEGER DEFAULT 0"),
        ("arp_count", "INTEGER DEFAULT 0"),
        ("rx_rate_bps", "INTEGER DEFAULT 0"),
        ("tx_rate_bps", "INTEGER DEFAULT 0"),
        ("total_rx_bytes", "INTEGER DEFAULT 0"),
        ("total_tx_bytes", "INTEGER DEFAULT 0"),
        ("cpu_usage", "INTEGER DEFAULT 0"),
        ("mem_usage", "INTEGER DEFAULT 0"),
        ("temp_c", "INTEGER DEFAULT 0"),
        ("telemetry_json", "TEXT DEFAULT '{}'"),
    ]
    for col_name, col_def in router_extra_cols:
        try:
            cursor.execute(f"ALTER TABLE room_network_status ADD COLUMN {col_name} {col_def}")
        except Exception:
            pass

    # 7. Bảng quản lý Lịch trình & Giờ Tăng ca (OT)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS overtime_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ot_date TEXT NOT NULL,
            staff_name TEXT NOT NULL,
            role TEXT NOT NULL,
            start_time TEXT,
            end_time TEXT,
            hours REAL NOT NULL,
            purpose TEXT,
            status TEXT DEFAULT 'Pending',
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # 8. Bảng Nhật ký công việc thường nhật của Admin (Work Log)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS work_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            work_date TEXT NOT NULL,
            shift_name TEXT NOT NULL,
            category TEXT NOT NULL,
            target_rooms TEXT NOT NULL,
            content TEXT NOT NULL,
            notes TEXT DEFAULT '',
            priority TEXT DEFAULT 'normal',
            status TEXT DEFAULT 'in_progress',
            created_by TEXT DEFAULT 'Admin',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            completed_date TEXT DEFAULT NULL
        )
    ''')

    # Migration an toàn: Thêm cột completed_date cho work_logs nếu chưa có
    try:
        cursor.execute("ALTER TABLE work_logs ADD COLUMN completed_date TEXT DEFAULT NULL")
    except Exception:
        pass
    cursor.execute("UPDATE work_logs SET status = 'in_progress' WHERE status = 'pending'")

    # 9. Bảng Checklist Deploy hệ thống 8 phòng máy theo từng đợt (SOP Acronis & Clonezilla)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS deployment_checklists (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            campaign_name TEXT NOT NULL,
            room TEXT NOT NULL,
            checked_steps TEXT DEFAULT '[]',
            notes TEXT DEFAULT '',
            updated_by TEXT DEFAULT 'Admin',
            updated_at TEXT,
            UNIQUE(campaign_name, room)
        )
    ''')

    # 10. Bảng Danh sách tài khoản Gmail đã xác minh Mã Passcode Giảng viên
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS verified_users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            full_name TEXT,
            verified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            verified_by TEXT DEFAULT 'passcode'
        )
    ''')

    # 11. Bảng Danh sách tài khoản bị khóa vĩnh viễn (Block User)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS blocked_users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            reason TEXT DEFAULT 'Vi phạm quy định / Spam hệ thống',
            blocked_by TEXT DEFAULT 'Admin',
            blocked_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # 12. Bảng Lưu vết tần suất gửi đơn để giới hạn tốc độ (Rate Limiting: tối đa 2 đơn / 60 giây)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS submission_rate_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL,
            action_type TEXT DEFAULT 'submit',
            created_ts REAL NOT NULL
        )
    ''')

    # 13. Bảng Danh sách Phân quyền Quản trị (Super Admin & Contributor)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin_permissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            role TEXT NOT NULL DEFAULT 'contributor',
            note TEXT DEFAULT '',
            added_by TEXT DEFAULT 'mlhang@hcmus.edu.vn',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('''
        INSERT OR IGNORE INTO admin_permissions (email, role, note, added_by)
        VALUES ('mlhang@hcmus.edu.vn', 'super_admin', 'Super Admin / Administrator - Toàn quyền quản trị', 'System')
    ''')
    cursor.execute('''
        UPDATE admin_permissions SET role = 'super_admin' WHERE LOWER(email) = 'mlhang@hcmus.edu.vn'
    ''')

    # Khởi tạo cấu hình mặc định cho B42Lab NetAgent & IP Router 8 phòng máy
    ak_row = cursor.execute("SELECT value FROM system_settings WHERE key = 'netagent_key'").fetchone()
    if not ak_row or not ak_row['value']:
        cursor.execute(
            "INSERT OR REPLACE INTO system_settings (key, value) VALUES ('netagent_key', 'b42lab-secret-2026')"
        )

    # Khởi tạo Mã Passcode nội bộ dành cho Thầy/Cô dùng Gmail cá nhân
    cursor.execute(
        "INSERT OR IGNORE INTO system_settings (key, value) VALUES ('teacher_passcode', 'QttB@2026')"
    )

    # Khởi tạo cấu hình mặc định cho Hệ thống Gửi Email Thông báo Kết quả (SMTP)
    default_smtp_settings = {
        'smtp_enabled': '1',
        'smtp_email': 'mlhang@hcmus.edu.vn',
        'smtp_password': '',
        'smtp_sender_name': 'Hệ thống Phòng máy B4.2 - Phòng QTTB',
        'admin_notify_enabled': '1'
    }
    for s_key, s_val in default_smtp_settings.items():
        cursor.execute(
            "INSERT OR IGNORE INTO system_settings (key, value) VALUES (?, ?)",
            (s_key, s_val)
        )

    # Khởi tạo trạng thái Đăng ký mượn phòng (Khóa / Mở cổng đăng ký)
    default_booking_settings = {
        'booking_enabled': '1',
        'booking_disabled_notice': 'Hệ thống đang tạm khóa đăng ký mượn phòng để Quản trị viên sắp xếp Thời khóa biểu. Quý Thầy/Cô vui lòng quay lại sau!'
    }
    for b_key, b_val in default_booking_settings.items():
        cursor.execute(
            "INSERT OR IGNORE INTO system_settings (key, value) VALUES (?, ?)",
            (b_key, b_val)
        )

    # [#QTTB-03] Khởi tạo cấu hình mặc định cho 8 Router (không lưu mật khẩu thật trong mã nguồn)
    default_router_settings = {
        'router_admin_user': 'admin',
        'router_admin_pass': '',
        'router_default_port': '443',
        'airpro_admin_user': 'admin',
        'airpro_admin_pass': '',
        'airpro_default_port': '80',
        'pc_mac_ouis': '4C:D7:17,00:68:EB,80:E8:2C',
        'ca1_start': '07:30',
        'ca1_end': '10:00',
        'ca2_start': '10:05',
        'ca2_end': '12:00',
        'ca3_start': '12:30',
        'ca3_end': '14:50',
        'ca4_start': '15:05',
        'ca4_end': '17:00',
    }
    for r_key, r_val in default_router_settings.items():
        cursor.execute(
            "INSERT OR IGNORE INTO system_settings (key, value) VALUES (?, ?)",
            (r_key, r_val)
        )

    # Bảng theo dõi phiên sử dụng PC theo Phòng x Ngày x Ca (Tuần -> Tháng -> Học kỳ)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS room_ca_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            log_date TEXT NOT NULL,
            day_of_week INTEGER NOT NULL,
            room TEXT NOT NULL,
            ca_slot INTEGER NOT NULL,
            ca_start_cfg TEXT DEFAULT '',
            ca_end_cfg TEXT DEFAULT '',
            subject_name TEXT DEFAULT '',
            class_code TEXT DEFAULT '',
            department TEXT DEFAULT '',
            first_online_at TEXT DEFAULT '',
            last_online_at TEXT DEFAULT '',
            last_sampled_at TEXT DEFAULT '',
            active_minutes INTEGER DEFAULT 0,
            idle_minutes INTEGER DEFAULT 0,
            last_pc_count INTEGER DEFAULT 0,
            peak_pc_count INTEGER DEFAULT 0,
            avg_pc_count REAL DEFAULT 0.0,
            sample_count INTEGER DEFAULT 0,
            UNIQUE(log_date, room, ca_slot)
        )
    ''')

    # Tự động mã hóa các mật khẩu router cũ đang lưu plaintext trong SQLite (nếu có)
    for sec_key in ('router_admin_pass', 'airpro_admin_pass'):
        sec_row = cursor.execute(
            "SELECT value FROM system_settings WHERE key = ?",
            (sec_key,)
        ).fetchone()
        if sec_row and sec_row['value'] and not str(sec_row['value']).startswith(_SECRET_PREFIX):
            enc_val = encrypt_router_secret(sec_row['value'])
            cursor.execute(
                "UPDATE system_settings SET value = ? WHERE key = ?",
                (enc_val, sec_key)
            )

    default_router_ips = {
        '5.1': '10.29.37.53',
        '5.2': '10.29.37.54',
        '5.3': '10.29.37.52',
        '6.1': '10.29.37.60',
        '6.2': '10.29.37.62',
        '6.3': '10.29.37.64',
        '6.5': '10.29.37.63',
        '6.6': '10.29.37.61',
    }
    default_router_models = {
        '5.1': 'AIRPRO',
        '5.2': 'GCC6010',
        '5.3': 'AIRPRO',
        '6.1': 'AIRPRO',
        '6.2': 'AIRPRO',
        '6.3': 'GCC6010',
        '6.5': 'AIRPRO',
        '6.6': 'AIRPRO',
    }
    default_lan_ips = {
        '5.1': '192.168.51.1',
        '5.2': '192.168.54.1',
        '5.3': '192.168.53.1',
        '6.1': '192.168.60.1',
        '6.2': '192.168.62.1',
        '6.3': '192.168.63.1',
        '6.5': '192.168.65.1',
        '6.6': '192.168.66.1',
    }
    cursor.execute("UPDATE room_network_status SET lan_ip = '192.168.60.1' WHERE room = '6.1' AND lan_ip = '192.168.61.1'")
    now_str = get_vn_now_str('%H:%M:%S %d/%m/%Y')
    airpro_migrated = cursor.execute("SELECT value FROM system_settings WHERE key = 'airpro_migrated_v1'").fetchone()
    if not airpro_migrated:
        for ar_room in ('5.1', '5.3', '6.1', '6.2', '6.5', '6.6'):
            cursor.execute(
                "UPDATE room_network_status SET router_model = 'AIRPRO', https_port = 80 WHERE room = ?",
                (ar_room,)
            )
        cursor.execute("INSERT OR REPLACE INTO system_settings (key, value) VALUES ('airpro_migrated_v1', '1')")

    # Đảm bảo không còn bản ghi cũ mang nhãn GWN7062G và đồng bộ đúng GCC6010 cho P5.2, P6.3
    cursor.execute("UPDATE room_network_status SET router_model = 'GCC6010', https_port = 443 WHERE room IN ('5.2', '6.3') AND COALESCE(router_model, '') != 'GCC6010'")
    cursor.execute("UPDATE room_network_status SET router_model = 'AIRPRO', https_port = 80 WHERE room IN ('5.1', '5.3', '6.1', '6.2', '6.5', '6.6') AND router_model = 'GWN7062G'")

    for r_code, real_ip in default_router_ips.items():
        r_model = default_router_models.get(r_code, 'AIRPRO')
        r_port = 80 if r_model == 'AIRPRO' else 443
        r_lan = default_lan_ips.get(r_code, '192.168.1.1')
        existing_net = cursor.execute(
            "SELECT room, target_ip, router_model, https_port, lan_ip FROM room_network_status WHERE room = ?", (r_code,)
        ).fetchone()
        if not existing_net:
            cursor.execute('''
                INSERT INTO room_network_status (
                    room, target_ip, status, ping_ms, updated_at, router_model, https_port, lan_ip
                ) VALUES (?, ?, 'online', 46, ?, ?, ?, ?)
            ''', (r_code, real_ip, now_str, r_model, r_port, r_lan))
        else:
            if existing_net['target_ip'] and existing_net['target_ip'].startswith('192.168.1.'):
                cursor.execute(
                    "UPDATE room_network_status SET target_ip = ? WHERE room = ?",
                    (real_ip, r_code)
                )
            if not existing_net['router_model']:
                cursor.execute(
                    "UPDATE room_network_status SET router_model = ? WHERE room = ?",
                    (r_model, r_code)
                )
            if not existing_net['https_port']:
                cursor.execute(
                    "UPDATE room_network_status SET https_port = ? WHERE room = ?",
                    (r_port, r_code)
                )
            if not existing_net['lan_ip']:
                cursor.execute(
                    "UPDATE room_network_status SET lan_ip = ? WHERE room = ?",
                    (r_lan, r_code)
                )

    # Khởi tạo dữ liệu mẫu lần đầu cho Nhật ký công việc
    seeded_row = cursor.execute("SELECT value FROM system_settings WHERE key = 'work_logs_seeded'").fetchone()
    if not seeded_row:
        existing_wl_count = cursor.execute("SELECT COUNT(*) as c FROM work_logs").fetchone()['c']
        if existing_wl_count == 0:
            today_iso = get_vn_today_iso()
            sample_logs = [
                (
                    today_iso, 'Ca sáng', 'server_net', '6.1, 6.2, Server/Rack',
                    'Cấu hình VLAN & kiểm tra băng thông Switch tầng 6 cho cụm phòng 6.1 - 6.6',
                    'Port Gi1/0/1-24 tốc độ ổn định 1Gbps, đã lưu cấu hình running-config vào startup-config',
                    'high', 'completed', 'Admin'
                ),
                (
                    today_iso, 'Ca sáng', 'vmware', '5.1, 5.2, 5.3',
                    'Kiểm thử hệ điều hành Windows 11 & phần mềm học kỳ mới trên VMware trước khi bung Image',
                    'Snapshot v2.1 chạy mượt Visual Studio 2022, MATLAB & Veyon Client; sẵn sàng tạo file Acronis chuẩn',
                    'high', 'in_progress', 'Admin'
                ),
                (
                    today_iso, 'Ca chiều', 'support', '5.2, 6.5',
                    'Hỗ trợ giảng viên xử lý phần mềm giảng dạy, kiểm tra cáp HDMI & cấu hình quyền ổ Z:',
                    'Đã chép thư mục t3h và kiểm tra quyền truy cập ổ Z: (pass: giaovien) trên máy GV',
                    'normal', 'completed', 'Admin'
                ),
                (
                    today_iso, 'Cả ngày', 'process', 'Tất cả 8 phòng',
                    'Cải tiến quy trình quản lý phòng máy & chuẩn hóa SOP triển khai Acronis + Clonezilla Lite Server',
                    'Tích hợp bảng Checklist 23 bước trực tiếp lên hệ thống Web quản trị QTTB',
                    'normal', 'completed', 'Admin'
                )
            ]
            cursor.executemany('''
                INSERT INTO work_logs (work_date, shift_name, category, target_rooms, content, notes, priority, status, created_by)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', sample_logs)
        cursor.execute("INSERT OR REPLACE INTO system_settings (key, value) VALUES ('work_logs_seeded', '1')")

    # Dọn sạch bản ghi đợt Deploy bị lỗi mã hóa ký tự (nếu có) và khởi tạo đợt mặc định nếu bảng đang trống
    cursor.execute("DELETE FROM deployment_checklists WHERE campaign_name LIKE '%?%'")
    has_campaign = cursor.execute("SELECT 1 FROM deployment_checklists LIMIT 1").fetchone()
    if not has_campaign:
        default_campaign = "HK1 (2026-2027) - Deploy Hệ thống chuẩn"
        now_short = get_vn_now_str('%H:%M %d/%m/%Y')
        for r_code in default_router_ips.keys():
            cursor.execute('''
                INSERT OR IGNORE INTO deployment_checklists (campaign_name, room, checked_steps, notes, updated_by, updated_at)
                VALUES (?, ?, '[]', '', 'Admin', ?)
            ''', (default_campaign, r_code, now_short))

    # =========================================================================
    # 14. [#QTTB-02] KÊNH THÔNG BÁO PHÒNG MÁY & ĐỀ XUẤT PHẦN MỀM KỲ TỚI
    # =========================================================================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS room_announcements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            room TEXT NOT NULL,
            category TEXT NOT NULL DEFAULT 'info',
            content TEXT NOT NULL,
            is_pinned INTEGER DEFAULT 0,
            created_by TEXT DEFAULT 'QTTB - Kỹ thuật phòng máy',
            created_at TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS room_software_specs (
            room TEXT PRIMARY KEY,
            spec_summary TEXT NOT NULL,
            updated_by TEXT DEFAULT 'Admin',
            updated_at TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS room_software_proposals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            room TEXT NOT NULL,
            proposal_type TEXT NOT NULL DEFAULT 'software_next_sem',
            target_scope TEXT NOT NULL DEFAULT 'next_semester',
            title_software TEXT NOT NULL,
            description TEXT DEFAULT '',
            proposer_name TEXT NOT NULL,
            proposer_email TEXT NOT NULL,
            upvotes INTEGER DEFAULT 1,
            upvoted_emails TEXT DEFAULT '[]',
            status TEXT DEFAULT 'recorded',
            admin_reply TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT
        )
    ''')

    cursor.execute(
        "INSERT OR IGNORE INTO system_settings (key, value) VALUES ('semester_survey_banner', '0')"
    )

    default_room_specs = {
        'ALL': 'Hệ thống 8 phòng máy B4.2 (328 máy tính học viên + 8 máy GV) • Windows 11 Pro 64-bit • Mạng nội bộ 1Gbps & Veyon Master.',
        '5.1': '41 máy • Win 11 Pro • Visual Studio 2022, VS Code, Python 3.11, MATLAB R2024a, Dev-C++, Git, Veyon.',
        '5.2': '41 máy • Win 11 Pro • Visual Studio 2022, VS Code, Python 3.11, SQL Server 2022, Docker Desktop, Postman.',
        '5.3': '41 máy • Win 11 Pro • Visual Studio 2022, Android Studio, JDK 17/21, Node.js, Python 3.11, Git.',
        '6.1': '41 máy • Win 11 Pro • MATLAB R2024a, Proteus 8, AutoCAD 2023, Keil C, Visual Studio 2022, Python 3.11.',
        '6.2': '41 máy • Win 11 Pro • Anaconda Python, RStudio, SPSS, PowerBI Desktop, Visual Studio 2022, VS Code.',
        '6.3': '41 máy • Win 11 Pro • VMware Workstation Pro, Cisco Packet Tracer, Wireshark, Linux Ubuntu VM, Python 3.11.',
        '6.5': '41 máy • Win 11 Pro • Visual Studio 2022, IntelliJ IDEA, SQL Server, MySQL Workbench, Python 3.11.',
        '6.6': '41 máy • Win 11 Pro • Visual Studio 2022, MATLAB R2024a, Python 3.11 (PyTorch/OpenCV), VS Code, Git.'
    }
    now_iso_str = get_vn_now_str('%Y-%m-%d %H:%M:%S')
    for r_key, spec_txt in default_room_specs.items():
        cursor.execute('''
            INSERT OR IGNORE INTO room_software_specs (room, spec_summary, updated_by, updated_at)
            VALUES (?, ?, 'Admin', ?)
        ''', (r_key, spec_txt, now_iso_str))

    hub_seeded = cursor.execute("SELECT value FROM system_settings WHERE key = 'room_hub_seeded'").fetchone()
    if not hub_seeded:
        ann_count = cursor.execute("SELECT COUNT(*) as c FROM room_announcements").fetchone()['c']
        if ann_count == 0:
            sample_anns = [
                (
                    'ALL', 'info',
                    'Chào mừng Quý Thầy/Cô đến với Kênh Thông Báo & Đề Xuất Phần Mềm Phòng Máy B4.2. Quý Thầy/Cô có thể chọn từng phòng (5.1 -> 6.6) để theo dõi thông báo vận hành hoặc gửi đề xuất nâng cấp phần mềm/thư viện cho Học kỳ tới.',
                    1, 'QTTB - Kỹ thuật phòng máy', now_iso_str
                ),
                (
                    '5.1', 'maintenance',
                    'Máy số 39 cuối phòng đang chờ bảo trì bộ nguồn. Quý Thầy/Cô vui lòng hướng dẫn sinh viên sử dụng các máy dự phòng 40 - 41.',
                    1, 'QTTB - Kỹ thuật phòng máy', now_iso_str
                ),
                (
                    '5.1', 'info',
                    'Đã cập nhật đầy đủ bộ thư viện chuẩn C/C++ và Python (numpy, pandas, matplotlib) cho toàn bộ 41 máy tại Phòng 5.1.',
                    0, 'QTTB - Kỹ thuật phòng máy', now_iso_str
                )
            ]
            cursor.executemany('''
                INSERT INTO room_announcements (room, category, content, is_pinned, created_by, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', sample_anns)

        prop_count = cursor.execute("SELECT COUNT(*) as c FROM room_software_proposals").fetchone()['c']
        if prop_count == 0:
            sample_props = [
                (
                    '5.1', 'software_next_sem', 'next_semester',
                    'Nâng cấp Python 3.12 & bổ sung thư viện PyTorch, OpenCV',
                    'Học kỳ tới bộ môn có triển khai học phần Thị giác máy tính và Học máy tại P.5.1, nhờ QTTB tích hợp sẵn vào file Image mới.',
                    'GV. Bộ môn KHMT', 'teacher_demo@hcmus.edu.vn', 2,
                    '["teacher_demo@hcmus.edu.vn", "mlhang@hcmus.edu.vn"]',
                    'planned_image',
                    'QTTB đã ghi nhận và đưa vào Checklist tạo Image mẫu cho Học kỳ tới.',
                    now_iso_str, now_iso_str
                )
            ]
            cursor.executemany('''
                INSERT INTO room_software_proposals (
                    room, proposal_type, target_scope, title_software, description,
                    proposer_name, proposer_email, upvotes, upvoted_emails, status,
                    admin_reply, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', sample_props)

        cursor.execute("INSERT OR REPLACE INTO system_settings (key, value) VALUES ('room_hub_seeded', '1')")

    # Nếu là lần đầu chạy và tạo db, ta insert 1 dòng dữ liệu mẫu để test giao diện
    if not db_exists:
        cursor.execute("""
            INSERT INTO schedules (room, day_of_week, start_period, end_period, class_code, subject_name, department, study_period, num_weeks, is_ended_early) 
            VALUES ('5.1', 3, 1, 5, '26KDL1A', 'Kỹ thuật lập trình', 'VLKT', '28/09/2026', 15, 0)
        """)
        
    conn.commit()
    conn.close()

# Đoạn code này chỉ chạy khi bạn thực thi trực tiếp file database.py để test
if __name__ == '__main__':
    init_db()
