import sqlite3
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
    
    # 5. Bảng lưu tài khoản Admin
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            email TEXT NOT NULL,
            reset_token TEXT
        )
    ''')
    
    # Tạo tài khoản admin mặc định nếu chưa có
    from werkzeug.security import generate_password_hash
    admin_exists = cursor.execute("SELECT * FROM admin WHERE username = 'admin'").fetchone()
    if not admin_exists:
        default_hash = generate_password_hash('admin123')
        cursor.execute("INSERT INTO admin (username, password_hash, email) VALUES (?, ?, ?)", 
                       ('admin', default_hash, 'mlhang@hcmus.edu.vn'))


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
