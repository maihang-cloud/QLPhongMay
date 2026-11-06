import sqlite3
import datetime
import os

# Khai báo đường dẫn tuyệt đối tới file database (tương thích cả Windows và PythonAnywhere Linux)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, 'lab_schedule.db')

def get_db_connection():
    """Hàm tạo kết nối tới database. Trả về object connection."""
    conn = sqlite3.connect(DB_NAME)
    # Lệnh này giúp trả về dữ liệu dưới dạng Dictionary thay vì Tuple, dễ xử lý bên Frontend hơn
    conn.row_factory = sqlite3.Row 
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
            status TEXT DEFAULT 'Pending',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 3. Bảng lưu Thông báo nghỉ gấp
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cancellations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_name TEXT NOT NULL,
            room TEXT NOT NULL,
            cancel_date TEXT NOT NULL,
            period TEXT NOT NULL,
            reason TEXT,
            status TEXT DEFAULT 'Pending',
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
            status TEXT DEFAULT 'Pending',
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
        cursor.execute("UPDATE booking_requests SET created_at = datetime('now', 'localtime') WHERE created_at IS NULL")
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

    cursor.execute("UPDATE schedules SET num_weeks = 15 WHERE num_weeks IS NULL")
    cursor.execute("UPDATE schedules SET is_ended_early = 0 WHERE is_ended_early IS NULL")
    cursor.execute("UPDATE schedules SET study_period = '28/09/2026' WHERE study_period LIKE '2026-09-28%'")

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
            updated_at TEXT
        )
    ''')

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
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

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

    # Khởi tạo cấu hình mặc định cho B42Lab NetAgent & IP Router 8 phòng máy
    ak_row = cursor.execute("SELECT value FROM system_settings WHERE key = 'netagent_key'").fetchone()
    if not ak_row or not ak_row['value']:
        cursor.execute(
            "INSERT OR REPLACE INTO system_settings (key, value) VALUES ('netagent_key', 'b42lab-secret-2026')"
        )

    # Khởi tạo cấu hình mặc định cho Hệ thống Gửi Email Thông báo Kết quả (SMTP)
    default_smtp_settings = {
        'smtp_enabled': '1',
        'smtp_email': 'mlhang@hcmus.edu.vn',
        'smtp_password': '',
        'smtp_sender_name': 'Hệ thống Phòng máy B4.2 - Phòng QTTB'
    }
    for s_key, s_val in default_smtp_settings.items():
        cursor.execute(
            "INSERT OR IGNORE INTO system_settings (key, value) VALUES (?, ?)",
            (s_key, s_val)
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
    now_str = datetime.datetime.now().strftime('%H:%M:%S %d/%m/%Y')
    for r_code, real_ip in default_router_ips.items():
        existing_net = cursor.execute(
            "SELECT room, target_ip FROM room_network_status WHERE room = ?", (r_code,)
        ).fetchone()
        if not existing_net:
            cursor.execute(
                "INSERT INTO room_network_status (room, target_ip, status, ping_ms, updated_at) VALUES (?, ?, 'online', 46, ?)",
                (r_code, real_ip, now_str)
            )
        elif existing_net['target_ip'] and existing_net['target_ip'].startswith('192.168.1.'):
            cursor.execute(
                "UPDATE room_network_status SET target_ip = ? WHERE room = ?",
                (real_ip, r_code)
            )

    # Khởi tạo dữ liệu mẫu lần đầu cho Nhật ký công việc
    seeded_row = cursor.execute("SELECT value FROM system_settings WHERE key = 'work_logs_seeded'").fetchone()
    if not seeded_row:
        existing_wl_count = cursor.execute("SELECT COUNT(*) as c FROM work_logs").fetchone()['c']
        if existing_wl_count == 0:
            today_iso = datetime.date.today().strftime('%Y-%m-%d')
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
        now_short = datetime.datetime.now().strftime('%H:%M %d/%m/%Y')
        for r_code in default_router_ips.keys():
            cursor.execute('''
                INSERT OR IGNORE INTO deployment_checklists (campaign_name, room, checked_steps, notes, updated_by, updated_at)
                VALUES (?, ?, '[]', '', 'Admin', ?)
            ''', (default_campaign, r_code, now_short))

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
