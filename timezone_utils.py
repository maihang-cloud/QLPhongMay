"""Module quản lý Múi giờ Việt Nam (UTC+7 / Asia/Ho_Chi_Minh) thống nhất cho toàn bộ hệ thống.
Đảm bảo khi triển khai trên localhost, Docker hay Server đám mây (PythonAnywhere, Linux VPS - mặc định UTC),
toàn bộ ngày giờ hiển thị, lưu vết Database, gửi email và đối soát TKB đều chuẩn 100% theo giờ Việt Nam.
"""

from datetime import datetime, date, timedelta, timezone

# Múi giờ Việt Nam: UTC+7 (không có giờ mùa hè DST)
VN_TZ = timezone(timedelta(hours=7), name='Asia/Ho_Chi_Minh')


def get_vn_now(naive: bool = True):
    """Lấy đối tượng datetime hiện tại theo múi giờ Việt Nam.
    naive=True (mặc định): Trả về offset-naive datetime (đã mang giá trị giờ VN) để tương thích
    an toàn với mọi phép so sánh datetime trong code hiện tại mà không bị lỗi type mismatch.
    """
    dt = datetime.now(VN_TZ)
    return dt.replace(tzinfo=None) if naive else dt


def get_vn_today() -> date:
    """Lấy đối tượng date ngày hôm nay theo múi giờ Việt Nam."""
    return datetime.now(VN_TZ).date()


def get_vn_now_str(fmt: str = '%H:%M:%S %d/%m/%Y') -> str:
    """Trả về chuỗi ngày giờ hiện tại theo định dạng chỉ định (múi giờ VN)."""
    return datetime.now(VN_TZ).strftime(fmt)


def get_vn_now_iso() -> str:
    """Trả về chuỗi ngày giờ chuẩn SQLite 'YYYY-MM-DD HH:MM:SS' theo múi giờ Việt Nam để lưu DB."""
    return datetime.now(VN_TZ).strftime('%Y-%m-%d %H:%M:%S')


def get_vn_today_iso() -> str:
    """Trả về chuỗi ngày chuẩn 'YYYY-MM-DD' theo múi giờ Việt Nam."""
    return datetime.now(VN_TZ).strftime('%Y-%m-%d')


def format_vn_datetime(dt_val, fmt: str = '%H:%M %d/%m/%Y') -> str:
    """Chuyển đổi chuỗi timestamp trong database (YYYY-MM-DD HH:MM:SS) thành định dạng dễ đọc."""
    if not dt_val:
        return ''
    s = str(dt_val).strip()
    try:
        # Nếu đã có định dạng 'HH:MM:SS DD/MM/YYYY' hoặc 'HH:MM DD/MM/YYYY'
        if '/' in s:
            return s
        # Nếu là dạng ISO 'YYYY-MM-DD HH:MM:SS' hoặc 'YYYY-MM-DDTHH:MM:SS'
        s_clean = s.replace('T', ' ')
        dt = datetime.strptime(s_clean[:19], '%Y-%m-%d %H:%M:%S')
        return dt.strftime(fmt)
    except Exception:
        return s
