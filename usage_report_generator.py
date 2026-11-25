"""
Module theo dõi phiên sử dụng PC theo 4 Ca học (kế thừa TKB + Đơn vị) & Xuất Báo cáo Excel
(Tuần -> Tháng -> Học kỳ) cho 8 Phòng máy B4.2.
"""

import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from report_exporter import (
    get_schedule_ca_slots,
    parse_report_date,
    get_course_date_span,
    classify_schedule_semester
)
from timezone_utils import get_vn_now, get_vn_today

DEFAULT_CA_TIMES = {
    1: ('07:30', '10:00'),
    2: ('10:05', '12:00'),
    3: ('12:30', '14:50'),
    4: ('15:05', '17:00'),
}


def get_ca_time_config(conn):
    """Đọc cấu hình giờ Bắt đầu / Kết thúc của 4 Ca học từ system_settings."""
    keys = [f'ca{i}_{part}' for i in range(1, 5) for part in ('start', 'end')]
    placeholders = ','.join('?' for _ in keys)
    rows = conn.execute(
        f"SELECT key, value FROM system_settings WHERE key IN ({placeholders})",
        keys
    ).fetchall()
    smap = {r['key']: (r['value'] or '').strip() for r in rows}

    cfg = {}
    for slot in range(1, 5):
        def_s, def_e = DEFAULT_CA_TIMES[slot]
        s_val = smap.get(f'ca{slot}_start') or def_s
        e_val = smap.get(f'ca{slot}_end') or def_e
        cfg[slot] = {'start': s_val, 'end': e_val}
    return cfg


def _hhmm_to_minutes(hhmm):
    try:
        parts = str(hhmm).strip().split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except Exception:
        return None


def detect_current_ca_slot(now_dt, ca_cfg):
    """Xác định Ca học hiện tại (1..4 hoặc 0 = Ngoài giờ) dựa vào giờ hệ thống và cấu hình 4 Ca."""
    cur_min = now_dt.hour * 60 + now_dt.minute
    for slot in range(1, 5):
        s_min = _hhmm_to_minutes(ca_cfg[slot]['start'])
        e_min = _hhmm_to_minutes(ca_cfg[slot]['end'])
        if s_min is not None and e_min is not None and s_min <= cur_min <= e_min:
            return slot, ca_cfg[slot]['start'], ca_cfg[slot]['end']
    return 0, '', ''


def format_duration_vn(minutes):
    """Định dạng số phút sang dạng 'X giờ Y phút' hoặc 'Y phút'."""
    m = max(0, int(minutes or 0))
    if m >= 60:
        h = m // 60
        rem = m % 60
        return f"{h} giờ {rem} phút" if rem else f"{h} giờ"
    return f"{m} phút"


def _is_schedule_active_on_date(sched_row, ref_date, approved_cancels):
    """Kiểm tra học phần có đang học trong ngày ref_date (trong chu kỳ tuần, chưa KT sớm, chưa báo vắng)."""
    if int(sched_row['is_ended_early'] or 0) == 1:
        return False

    sp_str = (sched_row['study_period'] or '').strip()
    if sp_str:
        # Hỗ trợ định dạng 'dd/mm/yyyy - dd/mm/yyyy' hoặc 'dd/mm/yy'
        parts = [p.strip() for p in sp_str.replace('–', '-').split('-') if p.strip()]
        start_d = parse_report_date(parts[0]) if parts else None
        if start_d:
            try:
                num_weeks = max(1, int(sched_row['num_weeks'] or 15))
            except Exception:
                num_weeks = 15
            end_d = parse_report_date(parts[1]) if len(parts) > 1 else (start_d + datetime.timedelta(weeks=num_weeks))
            if ref_date < start_d or (end_d and ref_date > end_d):
                return False

    # Kiểm tra báo vắng đã duyệt cho đúng ngày
    ref_iso = ref_date.strftime('%Y-%m-%d')
    for c in approved_cancels:
        if str(c['room']).strip() == str(sched_row['room']).strip() and str(c['cancel_date']).strip() == ref_iso:
            # Nếu khớp ca hoặc báo nghỉ cùng phòng ngày đó
            return False
    return True


def resolve_room_ca_schedule(conn, room, day_of_week, ca_slot, ref_date):
    """Kế thừa Môn học, Mã lớp và Khoa / Đơn vị từ TKB cho (room, day_of_week, ca_slot)."""
    if ca_slot not in (1, 2, 3, 4) or not (2 <= day_of_week <= 7):
        return '', '', ''

    rows = conn.execute(
        "SELECT * FROM schedules WHERE room = ? AND day_of_week = ? ORDER BY start_period ASC",
        (room, day_of_week)
    ).fetchall()
    if not rows:
        return '', '', ''

    cancels = conn.execute(
        "SELECT * FROM cancellations WHERE status = 'Approved' AND room = ? AND cancel_date = ?",
        (room, ref_date.strftime('%Y-%m-%d'))
    ).fetchall()

    for r in rows:
        slots = get_schedule_ca_slots(r['start_period'], r['end_period'])
        if ca_slot in slots and _is_schedule_active_on_date(r, ref_date, cancels):
            return (
                (r['subject_name'] or '').strip(),
                (r['class_code'] or '').strip(),
                (r['department'] or '').strip(),
            )
    return '', '', ''


ROOM_STUDENT_PC_CAPACITY = 40
AUTO_POLL_INTERVAL_MIN = 10
MIN_SAMPLE_GAP_SEC = 90
EDGE_WARMUP_COOLDOWN_MIN = 10


def _is_in_ca_edge_window(now_dt, ca_start, ca_end, margin_min=EDGE_WARMUP_COOLDOWN_MIN):
    """Kiểm tra thời điểm quét có rơi vào 10 phút đầu ca (SV mới vào) hoặc 10 phút cuối ca (SV ra về) không."""
    cur_min = now_dt.hour * 60 + now_dt.minute
    s_min = _hhmm_to_minutes(ca_start)
    e_min = _hhmm_to_minutes(ca_end)
    if s_min is None or e_min is None or e_min <= s_min:
        return False
    return (cur_min - s_min) < margin_min or (e_min - cur_min) < margin_min


def record_room_ca_session_sample(conn, poll_res, now_dt=None):
    """
    Ghi nhận & cộng dồn phiên sử dụng PC theo Phòng x Ngày x Ca mỗi lần quét/heartbeat.
    Quy chuẩn sạch dữ liệu:
    1. Bỏ qua mẫu lỗi kết nối (offline / wan_ping_only / connected_api=False) để không tính oan 0 PC.
    2. Bỏ qua ngoài giờ (ca_slot=0) và phòng có ca trống (không có môn học trên TKB).
    3. Chống nhiễu bấm quét liên tục (< 90s không tăng sample_count / active_minutes).
    4. Tách ngưỡng Máy giảng viên (<= 1 PC) vs Máy sinh viên (>= 2 PC) khi cộng dồn thời lượng sử dụng.
    5. Khử nhiễu 10 phút đầu ca / 10 phút cuối ca khi pc_cnt <= 1 để không kéo tụt trung bình PC.
    """
    now_dt = now_dt or get_vn_now()
    r_state = poll_res.get('room_state') or 'wan_ping_only'
    telem = poll_res.get('telemetry') or {}
    # 1. Không ghi nhận mẫu nếu mất kết nối SSH/API tới Router
    if r_state in ('offline', 'wan_ping_only') or not telem.get('connected_api', False):
        return

    ca_cfg = get_ca_time_config(conn)
    ca_slot, ca_start, ca_end = detect_current_ca_slot(now_dt, ca_cfg)
    if ca_slot not in (1, 2, 3, 4):
        return

    room = poll_res['room']
    log_date = now_dt.strftime('%Y-%m-%d')
    day_of_week = now_dt.weekday() + 2  # Thứ 2=2 .. CN=8
    subj_name, class_code, dept = resolve_room_ca_schedule(conn, room, day_of_week, ca_slot, now_dt.date())
    # 2. Phòng có ca trống (không có môn học trong TKB) -> không ghi vào báo cáo đánh giá sử dụng PC
    if not subj_name:
        return

    now_hm = now_dt.strftime('%H:%M')
    now_iso = now_dt.isoformat()
    pc_cnt = max(0, int(poll_res.get('online_clients') or 0))
    has_student_pcs = (pc_cnt >= 2)  # >= 2 PC: có máy SV sử dụng (trừ trường hợp chỉ 1 máy GV chiếu slide)
    in_edge = _is_in_ca_edge_window(now_dt, ca_start, ca_end)

    existing = conn.execute(
        "SELECT * FROM room_ca_sessions WHERE log_date = ? AND room = ? AND ca_slot = ?",
        (log_date, room, ca_slot)
    ).fetchone()

    if not existing:
        # Nếu rơi vào 10 phút đầu/cuối ca mà chưa có máy SV bật -> bỏ qua mẫu đầu tiên này để tránh nhiễu
        if in_edge and not has_student_pcs:
            return
        first_on = now_hm if has_student_pcs else ''
        last_on = now_hm if has_student_pcs else ''
        init_active = AUTO_POLL_INTERVAL_MIN if has_student_pcs else 0
        init_idle = AUTO_POLL_INTERVAL_MIN if (not has_student_pcs and r_state in ('lan_up_idle', 'open_active')) else 0
        conn.execute('''
            INSERT INTO room_ca_sessions (
                log_date, day_of_week, room, ca_slot, ca_start_cfg, ca_end_cfg,
                subject_name, class_code, department,
                first_online_at, last_online_at, last_sampled_at,
                active_minutes, idle_minutes, last_pc_count, peak_pc_count, avg_pc_count, sample_count
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
        ''', (
            log_date, day_of_week, room, ca_slot, ca_start, ca_end,
            subj_name, class_code, dept,
            first_on, last_on, now_iso,
            init_active, init_idle,
            pc_cnt, pc_cnt, float(pc_cnt)
        ))
        return

    ex = dict(existing)
    diff_sec = None
    if ex.get('last_sampled_at'):
        try:
            prev_dt = datetime.datetime.fromisoformat(ex['last_sampled_at'])
            diff_sec = max(0.0, (now_dt - prev_dt).total_seconds())
        except Exception:
            diff_sec = None

    first_on = ex.get('first_online_at') or ''
    last_on = ex.get('last_online_at') or ''
    active_min = int(ex.get('active_minutes') or 0)
    idle_min = int(ex.get('idle_minutes') or 0)
    old_n = int(ex.get('sample_count') or 0)
    old_avg = float(ex.get('avg_pc_count') or 0.0)
    new_peak = max(int(ex.get('peak_pc_count') or 0), pc_cnt)

    if has_student_pcs:
        if not first_on:
            first_on = now_hm
        last_on = now_hm

    # 3. Chống spam bấm quét liên tiếp (< 90s): chỉ cập nhật trạng thái hiện tại & đỉnh, không tăng sample_count/phút
    if diff_sec is not None and diff_sec < MIN_SAMPLE_GAP_SEC:
        conn.execute('''
            UPDATE room_ca_sessions
            SET last_pc_count = ?, peak_pc_count = ?, first_online_at = ?, last_online_at = ?, last_sampled_at = ?
            WHERE id = ?
        ''', (pc_cnt, new_peak, first_on, last_on, now_iso, ex['id']))
        return

    # 4. Nếu rơi vào 10 phút đầu/cuối ca mà pc_cnt <= 1 -> không kéo tụt avg_pc_count
    if in_edge and not has_student_pcs:
        conn.execute('''
            UPDATE room_ca_sessions
            SET last_pc_count = ?, last_sampled_at = ?
            WHERE id = ?
        ''', (pc_cnt, now_iso, ex['id']))
        return

    delta_min = AUTO_POLL_INTERVAL_MIN if diff_sec is None else min(15, max(1, int(round(diff_sec / 60.0))))
    if has_student_pcs:
        active_min += delta_min
    elif r_state in ('lan_up_idle', 'open_active'):
        idle_min += delta_min

    new_n = old_n + 1
    new_avg = round(((old_avg * old_n) + pc_cnt) / new_n, 1)

    conn.execute('''
        UPDATE room_ca_sessions
        SET ca_start_cfg = ?,
            ca_end_cfg = ?,
            subject_name = CASE WHEN subject_name = '' THEN ? ELSE subject_name END,
            class_code = CASE WHEN class_code = '' THEN ? ELSE class_code END,
            department = CASE WHEN department = '' THEN ? ELSE department END,
            first_online_at = ?,
            last_online_at = ?,
            last_sampled_at = ?,
            active_minutes = ?,
            idle_minutes = ?,
            last_pc_count = ?,
            peak_pc_count = ?,
            avg_pc_count = ?,
            sample_count = ?
        WHERE id = ?
    ''', (
        ca_start or ex.get('ca_start_cfg') or '',
        ca_end or ex.get('ca_end_cfg') or '',
        subj_name, class_code, dept,
        first_on, last_on, now_iso,
        active_min, idle_min,
        pc_cnt, new_peak, new_avg, new_n,
        ex['id']
    ))


def get_live_room_session_summary(conn, room, now_dt=None):
    """Lấy tóm tắt phiên hiện tại (Ca, Giờ, Môn, Đơn vị, Thời gian đã dùng) để hiển thị trên Dashboard."""
    now_dt = now_dt or get_vn_now()
    ca_cfg = get_ca_time_config(conn)
    ca_slot, ca_start, ca_end = detect_current_ca_slot(now_dt, ca_cfg)
    today_iso = now_dt.strftime('%Y-%m-%d')
    dow = now_dt.weekday() + 2
    full_set = _get_full_session_schedule_set(conn)

    def _build_live_ca_lbl(subj_txt, s_cfg, e_cfg):
        if ca_slot <= 0:
            return "Ngoài giờ"
        if subj_txt:
            if ca_slot in (1, 2) and (str(room).strip(), dow, subj_txt, (1, 2)) in full_set:
                return f"Ca 1 + Ca 2 ({ca_cfg[1]['start']}-{ca_cfg[2]['end']})"
            if ca_slot in (3, 4) and (str(room).strip(), dow, subj_txt, (3, 4)) in full_set:
                return f"Ca 3 + Ca 4 ({ca_cfg[3]['start']}-{ca_cfg[4]['end']})"
        return f"Ca {ca_slot} ({s_cfg or ca_start}-{e_cfg or ca_end})"

    row = conn.execute(
        "SELECT * FROM room_ca_sessions WHERE log_date = ? AND room = ? AND ca_slot = ?",
        (today_iso, room, ca_slot)
    ).fetchone()

    if row:
        d = dict(row)
        subj = (d.get('subject_name') or '').strip()
        dept = (d.get('department') or '').strip()
        act_m = int(d.get('active_minutes') or 0)
        ca_lbl = _build_live_ca_lbl(subj, d.get('ca_start_cfg'), d.get('ca_end_cfg'))
        return {
            'ca_slot': ca_slot,
            'ca_label': ca_lbl,
            'subject_name': subj,
            'department': dept,
            'first_online_at': d.get('first_online_at') or '',
            'active_minutes': act_m,
            'active_duration_str': format_duration_vn(act_m),
        }

    subj, _, dept = resolve_room_ca_schedule(conn, room, dow, ca_slot, now_dt.date())
    ca_lbl = _build_live_ca_lbl(subj, ca_start, ca_end)
    return {
        'ca_slot': ca_slot,
        'ca_label': ca_lbl,
        'subject_name': subj,
        'department': dept,
        'first_online_at': '',
        'active_minutes': 0,
        'active_duration_str': '0 phút',
    }


def _date_range_for_mode(mode, ref_date=None, conn=None):
    ref_date = ref_date or get_vn_today()
    if mode == 'month':
        start_d = ref_date.replace(day=1)
        if ref_date.month == 12:
            end_d = ref_date.replace(year=ref_date.year + 1, month=1, day=1) - datetime.timedelta(days=1)
        else:
            end_d = ref_date.replace(month=ref_date.month + 1, day=1) - datetime.timedelta(days=1)
        label = f"Tháng {ref_date.month:02d}/{ref_date.year}"
    elif mode == 'semester':
        sem_starts, sem_ends = [], []
        sem_name = "Học kỳ"
        if conn is not None:
            try:
                sch_rows = conn.execute("SELECT * FROM schedules").fetchall()
                target_sem = 'hk1' if ref_date.month in (8, 9, 10, 11, 12, 1) else 'hk2'
                for sr in sch_rows:
                    sc, yr, s_dt, e_dt = classify_schedule_semester(sr)
                    if s_dt and e_dt and sc == target_sem:
                        sem_starts.append(s_dt)
                        sem_ends.append(e_dt)
                if not sem_starts:
                    for sr in sch_rows:
                        s_dt, e_dt = get_course_date_span(sr)
                        if s_dt and e_dt:
                            sem_starts.append(s_dt)
                            sem_ends.append(e_dt)
                sem_name = "Học kỳ I" if target_sem == 'hk1' else "Học kỳ II"
            except Exception:
                pass
        if sem_starts and sem_ends:
            start_d = min(sem_starts)
            end_d = max(sem_ends)
        else:
            start_d = ref_date - datetime.timedelta(days=120)
            end_d = ref_date
        label = f"Tổng hợp {sem_name} ({start_d.strftime('%d/%m/%Y')} – {end_d.strftime('%d/%m/%Y')})"
    else:
        start_d = ref_date - datetime.timedelta(days=ref_date.weekday())
        end_d = start_d + datetime.timedelta(days=6)
        label = f"Tuần {start_d.strftime('%d/%m/%Y')} – {end_d.strftime('%d/%m/%Y')}"
    return start_d.strftime('%Y-%m-%d'), end_d.strftime('%Y-%m-%d'), label


def _get_full_session_schedule_set(conn):
    """
    Trả về tập (room, day_of_week, subject_name, session_pair) của các môn học thực sự kéo dài cả buổi
    (chiếm 2 ca trên cùng 1 dòng TKB: Tiết 1-5/1-4 -> (1, 2) hoặc Tiết 6-10/6-9 -> (3, 4)).
    Các môn tách 2 dòng TKB rời nhau (VD: T1-3.5 và T3.5-5) sẽ KHÔNG nằm trong tập này.
    """
    full_set = set()
    for r in conn.execute("SELECT room, day_of_week, start_period, end_period, subject_name FROM schedules").fetchall():
        slots = get_schedule_ca_slots(r['start_period'], r['end_period'])
        subj = (r['subject_name'] or '').strip()
        if not subj:
            continue
        rm = str(r['room']).strip()
        dow = int(r['day_of_week'] or 0)
        if 1 in slots and 2 in slots:
            full_set.add((rm, dow, subj, (1, 2)))
        if 3 in slots and 4 in slots:
            full_set.add((rm, dow, subj, (3, 4)))
    return full_set


def _merge_two_ca_rows(r1, r2, ca_label, ca_start, ca_end):
    """Ghép 2 bản ghi Ca liền kề trong cùng buổi (Ca 1+2 hoặc Ca 3+4) thành 1 buổi học thống nhất."""
    n1 = int(r1.get('sample_count') or 0)
    n2 = int(r2.get('sample_count') or 0)
    tot_n = max(1, n1 + n2)
    avg_pc = round(((float(r1.get('avg_pc_count') or 0.0) * n1) + (float(r2.get('avg_pc_count') or 0.0) * n2)) / tot_n, 1)

    f1 = (r1.get('first_online_at') or '').strip()
    f2 = (r2.get('first_online_at') or '').strip()
    first_on = f1 or f2

    return {
        'log_date': r1['log_date'],
        'day_of_week': r1['day_of_week'],
        'room': r1['room'],
        'ca_slot': r1['ca_slot'],
        'ca_label': ca_label,
        'ca_time': f"{ca_start} - {ca_end}",
        'subject_name': r1.get('subject_name') or r2.get('subject_name') or '',
        'class_code': r1.get('class_code') or r2.get('class_code') or '',
        'department': r1.get('department') or r2.get('department') or '',
        'first_online_at': first_on,
        'active_minutes': int(r1.get('active_minutes') or 0) + int(r2.get('active_minutes') or 0),
        'idle_minutes': int(r1.get('idle_minutes') or 0) + int(r2.get('idle_minutes') or 0),
        'last_pc_count': int(r2.get('last_pc_count') if n2 > 0 else (r1.get('last_pc_count') or 0)),
        'peak_pc_count': max(int(r1.get('peak_pc_count') or 0), int(r2.get('peak_pc_count') or 0)),
        'avg_pc_count': avg_pc,
        'sample_count': tot_n,
    }


def merge_full_session_rows(conn, raw_rows, exclude_empty_ca=True):
    """
    Chuẩn hóa danh sách phiên theo Ca học để đưa vào Báo cáo đánh giá sử dụng PC:
    1. Loại bỏ các phòng có Ca trống (không có lịch môn học trong TKB hoặc ngoài 4 ca chuẩn) khi exclude_empty_ca=True.
    2. Chỉ gộp Ca 1+2 (Sáng) hoặc Ca 3+4 (Chiều) đối với môn học thực sự kéo dài cả buổi trong TKB (full_set).
    3. Nếu môn học kéo dài cả buổi (VD: Tiết 1-5) nhưng hệ thống mới quét 1 ca đầu (Ca 1), vẫn chuẩn hóa nhãn
       thành 'Ca 1 + Ca 2 (Sáng)' với khung giờ cả buổi để báo cáo phản ánh đúng thời lượng môn học.
    """
    ca_cfg = get_ca_time_config(conn)
    full_set = _get_full_session_schedule_set(conn)

    by_day_room = {}
    for raw in raw_rows:
        r = dict(raw)
        sl = int(r.get('ca_slot') or 0)
        rm = str(r.get('room') or '').strip()
        dow = int(r.get('day_of_week') or 0)
        subj = (r.get('subject_name') or '').strip()

        # Đồng bộ lại từ TKB nếu bản ghi cũ chưa lưu tên môn học
        if not subj and sl in (1, 2, 3, 4):
            ref_d = parse_report_date(r.get('log_date')) or get_vn_today()
            s_name, c_code, dept_name = resolve_room_ca_schedule(conn, rm, dow, sl, ref_d)
            if s_name:
                r['subject_name'] = s_name
                r['class_code'] = r.get('class_code') or c_code
                r['department'] = r.get('department') or dept_name
                subj = s_name

        # Bỏ qua các phòng có ca trống (không có môn học trong TKB ở ca đó) hoặc ngoài giờ
        if exclude_empty_ca and (sl not in (1, 2, 3, 4) or not subj):
            continue

        by_day_room.setdefault((r['log_date'], rm), {})[sl] = r

    merged_list = []
    session_defs = (
        (1, 2, 'Ca 1 + Ca 2 (Sáng)'),
        (3, 4, 'Ca 3 + Ca 4 (Chiều)'),
    )
    for (log_date, room), slot_map in sorted(by_day_room.items(), key=lambda x: (x[0][0], x[0][1]), reverse=True):
        consumed = set()
        for s_first, s_second, period_lbl in session_defs:
            r1 = slot_map.get(s_first)
            r2 = slot_map.get(s_second)
            if r1 and r2:
                sj1 = (r1.get('subject_name') or '').strip()
                sj2 = (r2.get('subject_name') or '').strip()
                cc1 = (r1.get('class_code') or '').strip()
                cc2 = (r2.get('class_code') or '').strip()
                dow = int(r1.get('day_of_week') or 0)
                is_full = bool(
                    sj1 and sj1 == sj2 and cc1 == cc2
                    and (room, dow, sj1, (s_first, s_second)) in full_set
                )
                if is_full:
                    c_start = r1.get('ca_start_cfg') or ca_cfg[s_first]['start']
                    c_end = r2.get('ca_end_cfg') or ca_cfg[s_second]['end']
                    merged_list.append(_merge_two_ca_rows(r1, r2, period_lbl, c_start, c_end))
                    consumed.update((s_first, s_second))

        for sl in sorted(slot_map.keys()):
            if sl in consumed:
                continue
            r = dict(slot_map[sl])
            sj = (r.get('subject_name') or '').strip()
            dow = int(r.get('day_of_week') or 0)
            # Kiểm tra nếu môn này trong TKB là môn học kéo dài cả buổi (2 ca) nhưng mới quét 1 ca
            matched_full = None
            for s_first, s_second, period_lbl in session_defs:
                if sl in (s_first, s_second) and (room, dow, sj, (s_first, s_second)) in full_set:
                    matched_full = (s_first, s_second, period_lbl)
                    break

            if matched_full:
                sf, ss, period_lbl = matched_full
                r['ca_label'] = period_lbl
                r['ca_time'] = f"{ca_cfg[sf]['start']} - {ca_cfg[ss]['end']}"
            else:
                c_s = r.get('ca_start_cfg') or (ca_cfg[sl]['start'] if sl in ca_cfg else '')
                c_e = r.get('ca_end_cfg') or (ca_cfg[sl]['end'] if sl in ca_cfg else '')
                r['ca_label'] = f"Ca {sl}" if sl > 0 else "Ngoài giờ"
                r['ca_time'] = f"{c_s} - {c_e}".strip(' -') if sl > 0 else "Ngoài 4 ca"
            merged_list.append(r)

    merged_list.sort(key=lambda x: (x['log_date'], x['room'], -int(x['ca_slot'])), reverse=True)
    return merged_list


def classify_usage_and_recommendation(active_pc, total_samples, sessions_count=1, low_pc_threshold=10, capacity=ROOM_STUDENT_PC_CAPACITY):
    """
    Phân loại mức độ sử dụng PC sinh viên & Đề xuất phương án (Chuyển phòng lý thuyết / Theo dõi / Giữ phòng máy)
    dựa trên Số PC hoạt động (chuẩn trần 40 PC/ca) và Tỷ lệ sử dụng (% / 40 PC).
    Trả về tuple: (active_pc_int, util_pct, usage_level, recommendation, severity_code)
    """
    cap = max(1, int(capacity or ROOM_STUDENT_PC_CAPACITY))
    active_i = max(0, min(cap, int(round(float(active_pc or 0.0)))))
    samples_i = max(0, int(total_samples or 0))
    util_pct = round((active_i / float(cap)) * 100.0, 1)

    if samples_i < 2:
        return (
            active_i,
            util_pct,
            "Chưa đủ mẫu (< 2 lần quét)",
            "Chưa đủ dữ liệu — Cần quét thêm",
            "gray"
        )
    if active_i <= 1:
        return (
            active_i,
            util_pct,
            "Không dùng máy SV (0–1 PC)",
            "Đề xuất chuyển phòng lý thuyết",
            "red"
        )
    if active_i < low_pc_threshold:
        rec = (
            "Xem xét chuyển phòng lý thuyết / ghép lớp"
            if sessions_count >= 2
            else "Cảnh báo dùng rất ít — Theo dõi thêm"
        )
        return (
            active_i,
            util_pct,
            f"Sử dụng rất ít (< {round(low_pc_threshold * 100 / cap)}%)",
            rec,
            "red" if sessions_count >= 2 else "orange"
        )
    if active_i < int(cap * 0.5):
        return (
            active_i,
            util_pct,
            "Sử dụng trung bình thấp (25%–49%)",
            "Theo dõi thêm",
            "amber"
        )
    return (
        active_i,
        util_pct,
        "Khai thác tốt (≥ 50%)",
        "Giữ nguyên phòng máy",
        "green"
    )


def build_usage_excel_report(conn, mode='week', ref_date=None, low_pc_threshold=10):
    """
    Xuất file Excel báo cáo tần suất sử dụng PC 8 phòng (Tuần / Tháng / Học kỳ) phục vụ
    đánh giá tỷ lệ trung bình hoạt động PC theo từng Mã lớp học khác nhau (bất kể trùng tên môn)
    để xem xét chuyển sang phòng học lý thuyết.
    - Tự động loại bỏ các phòng có ca trống (không có lịch trong TKB).
    - Gom gọn còn 2 trường PC chuẩn: 'Số PC hoạt động' (số nguyên, trần 40 PC) và 'Tỷ lệ sử dụng (% / 40 PC)'.
    - Gồm 3 Sheet:
      1. De xuat Chuyen phong (Tổng hợp theo từng Mã lớp học khác nhau bất kể trùng tên)
      2. Chi tiet Buoi hoc (Bằng chứng chi tiết từng Ngày x Phòng x Ca)
      3. Canh bao Lang phi (Chỉ lọc các buổi đủ mẫu >= 2 lần quét và Số PC hoạt động < ngưỡng)
    """
    start_iso, end_iso, period_label = _date_range_for_mode(mode, ref_date=ref_date, conn=conn)
    raw_rows = [dict(r) for r in conn.execute('''
        SELECT * FROM room_ca_sessions
        WHERE log_date >= ? AND log_date <= ?
        ORDER BY log_date DESC, room ASC, ca_slot ASC
    ''', (start_iso, end_iso)).fetchall()]
    rows = merge_full_session_rows(conn, raw_rows, exclude_empty_ca=True)

    wb = Workbook()
    navy = PatternFill("solid", fgColor="0F2942")
    teal = PatternFill("solid", fgColor="0D9488")
    crimson = PatternFill("solid", fgColor="991B1B")
    sub_fill = PatternFill("solid", fgColor="E2E8F0")
    green = PatternFill("solid", fgColor="DCFCE7")
    amber = PatternFill("solid", fgColor="FEF3C7")
    orange = PatternFill("solid", fgColor="FFEDD5")
    red = PatternFill("solid", fgColor="FEE2E2")
    gray_fill = PatternFill("solid", fgColor="F1F5F9")
    zebra = PatternFill("solid", fgColor="F8FAFC")
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    sev_fill_map = {
        'red': red,
        'orange': orange,
        'amber': amber,
        'green': green,
        'gray': gray_fill,
    }
    sev_font_map = {
        'red': Font(name="Calibri", size=10, bold=True, color="991B1B"),
        'orange': Font(name="Calibri", size=10, bold=True, color="C2410C"),
        'amber': Font(name="Calibri", size=10, bold=True, color="92400E"),
        'green': Font(name="Calibri", size=10, bold=True, color="166534"),
        'gray': Font(name="Calibri", size=10, italic=True, color="475569"),
    }

    def _banner(ws, title, sub, cols):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
        c1 = ws.cell(row=1, column=1, value=title)
        c1.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        c1.fill = navy
        c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 28
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
        c2 = ws.cell(row=2, column=1, value=sub)
        c2.font = Font(name="Calibri", size=10, italic=True, color="334155")
        c2.fill = sub_fill
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[2].height = 20

    def _hdr(ws, r_idx, headers, fill=teal):
        ws.row_dimensions[r_idx].height = 26
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=r_idx, column=ci, value=h)
            c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c.fill = fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border

    def _autofit(ws):
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            ml = 0
            for c in col:
                if c.row <= 2:
                    continue
                ml = max(ml, len(str(c.value or "")))
            ws.column_dimensions[letter].width = max(10, min(42, ml + 4))

    # =========================================================================
    # SHEET 1: 1. De xuat Chuyen phong (Tổng hợp theo từng Mã lớp học độc lập)
    # =========================================================================
    ws1 = wb.active
    ws1.title = "1. De xuat Chuyen phong"
    _banner(
        ws1,
        "ĐÁNH GIÁ TỶ LỆ TRUNG BÌNH SỬ DỤNG PC THEO MÃ LỚP HỌC — ĐỀ XUẤT CHUYỂN PHÒNG LÝ THUYẾT",
        f"Kỳ báo cáo: {period_label} • Chuẩn 1 ca: {ROOM_STUDENT_PC_CAPACITY} PC SV • Phân tách độc lập theo từng Mã lớp học bất kể trùng tên môn • Đã loại ca trống",
        11
    )
    h1 = [
        "STT", "Mã lớp học", "Tên môn học (TKB)", "Đơn vị", "Phòng & Ca học",
        "Số buổi theo dõi", "Tổng mẫu quét", "Số PC hoạt động",
        "Tỷ lệ sử dụng (% / 40 PC)", "Phân loại mức dùng", "Đề xuất phương án"
    ]
    _hdr(ws1, 4, h1, fill=navy)

    grouped = {}
    for r in rows:
        ca_lbl = r.get('ca_label') or (f"Ca {r['ca_slot']}" if r['ca_slot'] > 0 else "Ngoài giờ")
        sj = (r.get('subject_name') or '').strip()
        cc = (r.get('class_code') or '').strip()
        dp = (r.get('department') or '').strip() or '—'
        rm = str(r.get('room') or '').strip()

        # Phân tách độc lập theo từng Mã lớp học bất kể trùng tên môn;
        # nếu Mã lớp học trống (VD: Tin học cơ sở), phân tách theo Phòng & Ca để không gộp lẫn.
        if cc:
            key = ('class', cc, sj, dp)
        else:
            key = ('room_ca', '', sj, dp, rm, ca_lbl)

        g = grouped.setdefault(key, {
            'class_code': cc or '—',
            'subject': sj,
            'dept': dp,
            'room_cas': [],
            'sessions': 0,
            'samples': 0,
            'session_pc_sum': 0,
        })
        rc_tag = f"P{rm} • {ca_lbl}"
        if rc_tag not in g['room_cas']:
            g['room_cas'].append(rc_tag)
        sc = max(1, int(r.get('sample_count') or 1))
        sess_pc = max(0, min(ROOM_STUDENT_PC_CAPACITY, int(r.get('peak_pc_count') or 0)))
        g['sessions'] += 1
        g['samples'] += sc
        g['session_pc_sum'] += sess_pc

    summary_items = []
    sev_priority = {'red': 0, 'orange': 1, 'amber': 2, 'gray': 3, 'green': 4}
    for g in grouped.values():
        avg_sess_pc = g['session_pc_sum'] / float(max(1, g['sessions']))
        active_pc, util_pct, lvl_txt, rec_txt, sev = classify_usage_and_recommendation(
            avg_sess_pc, g['samples'], g['sessions'], low_pc_threshold=low_pc_threshold
        )
        summary_items.append({
            'class_code': g['class_code'],
            'subject': g['subject'],
            'dept': g['dept'],
            'room_ca_str': "; ".join(g['room_cas']),
            'sessions': g['sessions'],
            'samples': g['samples'],
            'active_pc': active_pc,
            'util_pct': util_pct,
            'lvl_txt': lvl_txt,
            'rec_txt': rec_txt,
            'sev': sev,
        })

    # Ưu tiên đưa các mã lớp đề xuất chuyển phòng lý thuyết (red/orange) lên đầu bảng
    summary_items.sort(key=lambda x: (sev_priority.get(x['sev'], 9), x['active_pc'], x['subject'], x['class_code']))

    for idx, item in enumerate(summary_items, start=1):
        r_idx = idx + 4
        vals = [
            idx,
            item['class_code'],
            item['subject'],
            item['dept'],
            item['room_ca_str'],
            item['sessions'],
            item['samples'],
            item['active_pc'],
            f"{item['util_pct']}%",
            item['lvl_txt'],
            item['rec_txt'],
        ]
        for ci, val in enumerate(vals, start=1):
            c = ws1.cell(row=r_idx, column=ci, value=val)
            c.font = Font(name="Calibri", size=10, color="0F172A")
            c.border = border
            c.alignment = Alignment(horizontal="left" if ci in (3, 4, 5, 10, 11) else "center", vertical="center")
            if ci in (10, 11):
                c.fill = sev_fill_map[item['sev']]
                c.font = sev_font_map[item['sev']]
            elif r_idx % 2 == 1:
                c.fill = zebra
    ws1.freeze_panes = "A5"
    _autofit(ws1)

    # =========================================================================
    # SHEET 2: 2. Chi tiet Buoi hoc (Bằng chứng chi tiết từng Ngày x Phòng x Ca)
    # =========================================================================
    ws2 = wb.create_sheet("2. Chi tiet Buoi hoc")
    _banner(
        ws2,
        "CHI TIẾT TÌNH TRẠNG SỬ DỤNG PC THEO TỪNG BUỔI HỌC (8 PHÒNG MÁY B4.2)",
        f"Kỳ báo cáo: {period_label} • Chuẩn 1 ca: {ROOM_STUDENT_PC_CAPACITY} PC SV • Chỉ hiển thị ca có lịch TKB (đã bỏ ca trống & mẫu lỗi kết nối)",
        12
    )
    h2 = [
        "Ngày", "Phòng", "Ca học", "Khung giờ", "Mã lớp học", "Môn học (TKB)",
        "Đơn vị", "Số mẫu quét", "Số PC hoạt động", "Tỷ lệ sử dụng (% / 40 PC)",
        "Thời lượng dùng PC SV", "Đánh giá buổi học"
    ]
    _hdr(ws2, 4, h2, fill=teal)

    for idx, r in enumerate(rows, start=5):
        ca_lbl = r.get('ca_label') or (f"Ca {r['ca_slot']}" if r['ca_slot'] > 0 else "Ngoài giờ")
        ca_time = r.get('ca_time') or (f"{r.get('ca_start_cfg') or ''} - {r.get('ca_end_cfg') or ''}".strip(' -'))
        samples = int(r.get('sample_count') or 0)
        raw_peak = int(r.get('peak_pc_count') or 0)
        act_min = int(r.get('active_minutes') or 0)
        act_str = format_duration_vn(act_min) if act_min > 0 else (f"~{AUTO_POLL_INTERVAL_MIN} phút (1 mẫu)" if raw_peak >= 2 else "0 phút")

        active_pc, util_pct, lvl_txt, _, sev = classify_usage_and_recommendation(
            raw_peak, samples, sessions_count=1, low_pc_threshold=low_pc_threshold
        )

        vals = [
            r['log_date'],
            f"P{r['room']}",
            ca_lbl,
            ca_time,
            r.get('class_code') or '—',
            r.get('subject_name') or '—',
            r.get('department') or '—',
            samples,
            active_pc,
            f"{util_pct}%",
            act_str,
            lvl_txt
        ]
        for ci, val in enumerate(vals, start=1):
            c = ws2.cell(row=idx, column=ci, value=val)
            c.font = Font(name="Calibri", size=10, color="0F172A")
            c.border = border
            c.alignment = Alignment(horizontal="left" if ci in (6, 7, 12) else "center", vertical="center")
            if ci == 12:
                c.fill = sev_fill_map[sev]
                c.font = sev_font_map[sev]
            elif idx % 2 == 1:
                c.fill = zebra
    ws2.freeze_panes = "A5"
    _autofit(ws2)

    # =========================================================================
    # SHEET 3: 3. Canh bao Lang phi (Chỉ lọc buổi đủ mẫu >= 2 & PC < ngưỡng)
    # =========================================================================
    ws3 = wb.create_sheet("3. Canh bao Lang phi")
    _banner(
        ws3,
        "NHẬT KÝ CẢNH BÁO CÁC BUỔI HỌC KHÔNG DÙNG / DÙNG RẤT ÍT MÁY TÍNH",
        f"Kỳ báo cáo: {period_label} • Điều kiện lọc: Số mẫu quét ≥ 2 lần (đủ tin cậy) và Số PC hoạt động < {low_pc_threshold} PC (< 25% công suất 40 PC)",
        9
    )
    h3 = [
        "Ngày & Ca học", "Phòng", "Mã lớp học", "Môn học", "Đơn vị",
        "Số mẫu quét", "Số PC hoạt động", "Tỷ lệ sử dụng (% / 40 PC)", "Mức độ cảnh báo"
    ]
    _hdr(ws3, 4, h3, fill=crimson)

    waste_rows = [
        r for r in rows
        if (r.get('subject_name') or '').strip()
        and int(r.get('sample_count') or 0) >= 2
        and min(ROOM_STUDENT_PC_CAPACITY, int(r.get('peak_pc_count') or 0)) < low_pc_threshold
    ]
    waste_rows.sort(key=lambda r: (int(r.get('peak_pc_count') or 0), r['log_date']), reverse=False)

    if not waste_rows:
        ws3.merge_cells("A5:I5")
        c_empty = ws3.cell(row=5, column=1, value="Chưa ghi nhận buổi học nào đủ điều kiện cảnh báo lãng phí (≥ 2 mẫu quét và < 10 PC) trong kỳ báo cáo.")
        c_empty.font = Font(name="Calibri", size=10, italic=True, color="475569")
        c_empty.alignment = Alignment(horizontal="center", vertical="center")
        ws3.row_dimensions[5].height = 24
    else:
        for idx, r in enumerate(waste_rows, start=5):
            ca_lbl = r.get('ca_label') or (f"Ca {r['ca_slot']}" if r['ca_slot'] > 0 else "Ngoài giờ")
            ca_time = r.get('ca_time') or (f"{r.get('ca_start_cfg') or ''} - {r.get('ca_end_cfg') or ''}".strip(' -'))
            time_str = f"{r['log_date']} • {ca_lbl} ({ca_time})" if ca_time else f"{r['log_date']} • {ca_lbl}"
            samples = int(r.get('sample_count') or 0)
            active_pc = max(0, min(ROOM_STUDENT_PC_CAPACITY, int(r.get('peak_pc_count') or 0)))
            util_pct = round((active_pc / float(ROOM_STUDENT_PC_CAPACITY)) * 100.0, 1)
            if active_pc <= 1:
                warn_txt = "Không dùng máy SV (0–1 PC) — Đề xuất chuyển phòng lý thuyết"
                w_fill, w_font = red, sev_font_map['red']
            else:
                warn_txt = f"Dùng rất ít ({active_pc}/{ROOM_STUDENT_PC_CAPACITY} PC) — Cần theo dõi"
                w_fill, w_font = orange, sev_font_map['orange']

            vals = [
                time_str,
                f"P{r['room']}",
                r.get('class_code') or '—',
                r.get('subject_name') or '—',
                r.get('department') or '—',
                samples,
                active_pc,
                f"{util_pct}%",
                warn_txt,
            ]
            for ci, val in enumerate(vals, start=1):
                c = ws3.cell(row=idx, column=ci, value=val)
                c.font = Font(name="Calibri", size=10, color="0F172A")
                c.border = border
                c.alignment = Alignment(horizontal="left" if ci in (1, 4, 5, 9) else "center", vertical="center")
                if ci == 9:
                    c.fill = w_fill
                    c.font = w_font
    ws3.freeze_panes = "A5"
    _autofit(ws3)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, period_label
