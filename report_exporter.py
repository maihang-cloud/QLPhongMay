import io
import re
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOMS_LIST = ['5.1', '5.2', '5.3', '6.1', '6.2', '6.3', '6.5', '6.6']
MAX_CA_PER_ROOM_DAY = 4       # 2 ca sáng + 2 ca chiều
MAX_CA_PER_ROOM_WEEK = 24     # 6 ngày (Thứ 2 -> Thứ 7) x 4 ca/ngày
MAX_CA_SYSTEM_WEEK = 192      # 8 phòng x 24 ca/tuần
MAX_CA_SYSTEM_SESSION = 16    # 8 phòng x 2 ca/buổi (Sáng hoặc Chiều)


def parse_report_date(raw_val):
    """Phân tích chuỗi ngày bắt đầu (study_period) thành datetime.date."""
    if not raw_val:
        return None
    if isinstance(raw_val, datetime.datetime):
        return raw_val.date()
    if isinstance(raw_val, datetime.date):
        return raw_val
    s = str(raw_val).strip()
    if not s:
        return None
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d'):
        try:
            part = s[:10] if fmt in ('%Y-%m-%d', '%Y/%m/%d') and len(s) >= 10 else s
            return datetime.datetime.strptime(part, fmt).date()
        except ValueError:
            pass
    m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', s)
    if m:
        try:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    return None


def get_schedule_ca_slots(start_period, end_period):
    """
    Quy đổi khoảng tiết (start_period -> end_period) thành danh sách các Ca chuẩn trong ngày:
    - Ca 1 (Sáng 1): Tiết 1 – 3.5 (07h30 – 10h00)
    - Ca 2 (Sáng 2): Tiết 3.5 – 5 (10h05 – 12h00)
    - Ca 3 (Chiều 1): Tiết 6 – 8.5 (12h30 – 14h50)
    - Ca 4 (Chiều 2): Tiết 8.5 – 10 (15h05 – 17h00)
    Mỗi ca chuẩn = 2,5 tiết. Lớp kéo dài cả buổi (VD: Tiết 1-5 hoặc Tiết 6-10) chiếm 2 ca.
    """
    try:
        start = float(start_period)
        end = float(end_period)
    except (ValueError, TypeError):
        return [1]

    slots = []
    # Xét buổi sáng (Tiết 1 -> 5)
    if start < 5.5:
        if start <= 1.5 and end >= 4.5:
            slots.extend([1, 2])
        elif start >= 3.0:
            slots.append(2)
        else:
            slots.append(1)

    # Xét buổi chiều (Tiết 6 -> 10)
    if end > 5.5:
        if start <= 6.5 and end >= 9.5:
            slots.extend([3, 4])
        elif start >= 8.0:
            slots.append(4)
        elif start >= 5.5:
            slots.append(3)
        elif end >= 8.5:
            slots.extend([3, 4])
        else:
            slots.append(3)

    return slots if slots else [1]


def get_course_date_span(row):
    """
    Tính toán chính xác ngày Bắt đầu (start_dt) và ngày Kết thúc (end_dt) của một môn học
    dựa trên cột study_period (Giai_Doan), chu kỳ lặp num_weeks và ngày kết thúc sớm (nếu có).
    """
    keys = row.keys() if hasattr(row, 'keys') else row
    raw_period = row['study_period'] if 'study_period' in keys else ''
    start_dt = parse_report_date(raw_period)
    if not start_dt:
        return None, None

    num_weeks = int(row['num_weeks'] or 15) if 'num_weeks' in keys else 15
    start_monday = start_dt - datetime.timedelta(days=start_dt.weekday())
    end_dt = start_monday + datetime.timedelta(days=num_weeks * 7 - 1)

    if 'is_ended_early' in keys and int(row['is_ended_early'] or 0):
        early_dt = parse_report_date(row['ended_early_date'] if 'ended_early_date' in keys else '')
        if early_dt and start_dt <= early_dt <= end_dt:
            end_dt = early_dt

    return start_dt, end_dt


def classify_schedule_semester(row):
    """
    Phân loại Học kỳ (hk1 / hk2) và Năm học của môn học dựa trên thời gian BĐ (start_dt) và KT (end_dt):
    - Học kỳ 1 (hk1): Bắt đầu từ tháng 8 -> tháng 12 (kết thúc khoảng tháng 1 năm sau),
      hoặc bắt đầu trong tháng 1 và kết thúc trước ngày 15/02.
    - Học kỳ 2 (hk2): Bắt đầu từ tháng 2 trở đi -> tháng 7 (kết thúc khoảng tháng 5 - tháng 7).
    Trả về tuple: (semester_code, academic_year_str, start_dt, end_dt)
    """
    start_dt, end_dt = get_course_date_span(row)
    if start_dt:
        if start_dt.month in (8, 9, 10, 11, 12):
            return 'hk1', f"{start_dt.year} - {start_dt.year + 1}", start_dt, end_dt
        if start_dt.month == 1:
            if end_dt and end_dt.month >= 3:
                return 'hk2', f"{start_dt.year - 1} - {start_dt.year}", start_dt, end_dt
            return 'hk1', f"{start_dt.year - 1} - {start_dt.year}", start_dt, end_dt
        # Tháng 2 -> Tháng 7: Học kỳ 2
        return 'hk2', f"{start_dt.year - 1} - {start_dt.year}", start_dt, end_dt

    keys = row.keys() if hasattr(row, 'keys') else row
    s = str((row['study_period'] if 'study_period' in keys else '') or '').strip().lower()
    if 'hk2' in s or 'học kỳ 2' in s or 'hoc ky 2' in s or 'kỳ 2' in s:
        return 'hk2', '', None, None
    if 'hk1' in s or 'học kỳ 1' in s or 'hoc ky 1' in s or 'kỳ 1' in s:
        return 'hk1', '', None, None
    return 'unclassified', '', None, None


def format_date_span(rows_with_dates):
    """Định dạng chuỗi khoảng thời gian BĐ sớm nhất -> KT muộn nhất từ danh sách (row, start_dt, end_dt)."""
    starts = [s for _, s, _ in rows_with_dates if s is not None]
    ends = [e for _, _, e in rows_with_dates if e is not None]
    if starts and ends:
        return f"{min(starts).strftime('%d/%m/%Y')} – {max(ends).strftime('%d/%m/%Y')}"
    return "Chưa có lịch"


def summarize_available_semesters(all_schedules):
    """
    Tổng hợp nhanh số lượng lớp, số ca và khoảng thời gian BĐ – KT của Học kỳ 1 và Học kỳ 2
    hiện có trong Database để hiển thị trực quan trên giao diện trang Lịch học.
    """
    hk1_items = []
    hk2_items = []
    hk1_years = {}
    hk2_years = {}

    for r in all_schedules:
        sem, yr, s_dt, e_dt = classify_schedule_semester(r)
        if sem == 'hk1':
            hk1_items.append((r, s_dt, e_dt))
            if yr:
                hk1_years[yr] = hk1_years.get(yr, 0) + 1
        elif sem == 'hk2':
            hk2_items.append((r, s_dt, e_dt))
            if yr:
                hk2_years[yr] = hk2_years.get(yr, 0) + 1

    hk1_ca = sum(len(get_schedule_ca_slots(r['start_period'], r['end_period'])) for r, _, _ in hk1_items)
    hk2_ca = sum(len(get_schedule_ca_slots(r['start_period'], r['end_period'])) for r, _, _ in hk2_items)

    best_hk1_yr = max(hk1_years.items(), key=lambda x: x[1])[0] if hk1_years else ''
    best_hk2_yr = max(hk2_years.items(), key=lambda x: x[1])[0] if hk2_years else best_hk1_yr

    all_starts = [s for _, s, _ in (hk1_items + hk2_items) if s is not None]
    all_ends = [e for _, _, e in (hk1_items + hk2_items) if e is not None]

    return {
        'hk1_classes': len(hk1_items),
        'hk1_ca': hk1_ca,
        'hk1_span': format_date_span(hk1_items),
        'hk1_year': best_hk1_yr,
        'hk2_classes': len(hk2_items),
        'hk2_ca': hk2_ca,
        'hk2_span': format_date_span(hk2_items),
        'hk2_year': best_hk2_yr,
        'min_start_iso': min(all_starts).strftime('%Y-%m-%d') if all_starts else '',
        'max_end_iso': max(all_ends).strftime('%Y-%m-%d') if all_ends else '',
    }


def filter_schedules_for_report(
    all_schedules,
    semester_mode='hk1',
    academic_year='',
    from_date_str='',
    to_date_str=''
):
    """
    Lọc danh sách lịch học chính xác theo Học kỳ (hk1 / hk2) hoặc Khoảng thời gian BĐ – KT (custom).
    Tuyệt đối KHÔNG tự ý lấy dữ liệu của Học kỳ 1 thế vào khi chọn Học kỳ 2.
    """
    semester_mode = (semester_mode or 'hk1').strip().lower()
    req_year = (academic_year or '').strip()
    norm_req_year = re.sub(r'\s+', '', req_year)

    from_dt = parse_report_date(from_date_str)
    to_dt = parse_report_date(to_date_str)

    matched_tuples = []
    detected_years = {}

    if semester_mode == 'custom' and (from_dt or to_dt):
        for r in all_schedules:
            sem, yr, s_dt, e_dt = classify_schedule_semester(r)
            if not s_dt:
                continue
            # Môn học thuộc khoảng lọc nếu thời gian BĐ nằm từ from_dt trở đi và thời gian BĐ/KT nằm trong to_dt
            if from_dt and s_dt < from_dt:
                continue
            if to_dt and s_dt > to_dt:
                continue
            matched_tuples.append((r, s_dt, e_dt))
            if yr:
                detected_years[yr] = detected_years.get(yr, 0) + 1

        span_lbl = format_date_span(matched_tuples)
        f_str = from_dt.strftime('%d/%m/%Y') if from_dt else '...'
        t_str = to_dt.strftime('%d/%m/%Y') if to_dt else '...'
        semester_label = f"GIAI ĐOẠN ({f_str} – {t_str})"
        semester_short = f"Giai đoạn {f_str} – {t_str}"
    else:
        target_sem = 'hk2' if semester_mode == 'hk2' else 'hk1'
        for r in all_schedules:
            sem, yr, s_dt, e_dt = classify_schedule_semester(r)
            if sem != target_sem:
                continue
            if norm_req_year and yr and re.sub(r'\s+', '', yr) != norm_req_year:
                continue
            matched_tuples.append((r, s_dt, e_dt))
            if yr:
                detected_years[yr] = detected_years.get(yr, 0) + 1

        if target_sem == 'hk2':
            semester_label = "HỌC KỲ II"
            semester_short = "Học kỳ 2"
        else:
            semester_label = "HỌC KỲ I"
            semester_short = "Học kỳ 1"
        span_lbl = format_date_span(matched_tuples)

    if not req_year:
        if detected_years:
            req_year = max(detected_years.items(), key=lambda x: x[1])[0]
        else:
            today = datetime.date.today()
            req_year = f"{today.year} - {today.year + 1}" if today.month >= 8 else f"{today.year - 1} - {today.year}"

    target_rows = [item[0] for item in matched_tuples]
    return {
        'rows': target_rows,
        'semester_label': semester_label,
        'semester_short': semester_short,
        'academic_year': req_year,
        'date_span_str': span_lbl,
    }


def clean_and_group_departments(schedule_rows):
    """
    Chuẩn hóa tên Đơn vị (loại bỏ tiền tố 'Khoa ', gộp các tên tương đương như
    'Khoa Điện tử Viễn thông' và 'Điện tử - Viễn thông' về tên chuẩn phổ biến nhất).
    Trả về dict ánh xạ từ raw_department -> canonical_department_name.
    """
    norm_Votes = {}
    raw_to_norm = {}

    for r in schedule_rows:
        raw_dept = (r['department'] if 'department' in r.keys() else '') or ''
        raw_dept = raw_dept.strip()
        if not raw_dept:
            clean_dept = 'Khác'
        else:
            clean_dept = re.sub(r'^[Kk]hoa\s+', '', raw_dept).strip()
            if not clean_dept:
                clean_dept = raw_dept
        norm_key = re.sub(r'[\s\-–—/]+', ' ', clean_dept.lower()).strip()
        raw_to_norm[raw_dept] = norm_key
        if norm_key not in norm_Votes:
            norm_Votes[norm_key] = {}
        norm_Votes[norm_key][clean_dept] = norm_Votes[norm_key].get(clean_dept, 0) + 1

    norm_to_best = {}
    for norm_key, candidates in norm_Votes.items():
        # Ưu tiên tên có dấu gạch nối chuẩn (như 'Điện tử - Viễn thông') hoặc xuất hiện nhiều nhất
        best_label = max(candidates.items(), key=lambda x: (x[1], '-' in x[0], len(x[0])))[0]
        norm_to_best[norm_key] = best_label

    return {raw: norm_to_best[norm] for raw, norm in raw_to_norm.items()}


def compute_semester_statistics(schedule_rows, rooms=None):
    """
    Tính toán toàn bộ chỉ số thống kê tần suất khai thác phòng máy theo chuẩn:
    - 1 ca = 2.5 tiết
    - 1 ngày = 4 ca (2 Sáng, 2 Chiều)
    - 1 tuần = 24 ca / phòng -> 8 phòng = 192 ca / tuần
    """
    if rooms is None:
        rooms = ROOMS_LIST

    raw_to_canonical = clean_and_group_departments(schedule_rows)

    room_ca_used = {r: 0 for r in rooms}
    room_class_count = {r: 0 for r in rooms}
    room_day_session = {r: {(d, s): 0 for d in range(2, 8) for s in ('S', 'C')} for r in rooms}

    dept_ca_total = {}
    dept_class_count = {}
    dept_room_matrix = {}
    dept_day_session_matrix = {}

    day_session_totals = {(d, s): 0 for d in range(2, 8) for s in ('S', 'C')}
    day_ca_totals = {d: 0 for d in range(2, 8)}

    total_ca_used = 0
    total_classes = len(schedule_rows)

    for row in schedule_rows:
        room = str(row['room'] or '').strip()
        try:
            dow = int(row['day_of_week'])
        except (ValueError, TypeError):
            dow = 2
        if dow < 2 or dow > 7:
            dow = 2

        raw_dept = (row['department'] if 'department' in row.keys() else '') or ''
        raw_dept = raw_dept.strip()
        dept = raw_to_canonical.get(raw_dept, 'Khác')

        slots = get_schedule_ca_slots(row['start_period'], row['end_period'])
        num_ca = len(slots)

        total_ca_used += num_ca

        if room in room_ca_used:
            room_ca_used[room] += num_ca
            room_class_count[room] += 1

        dept_ca_total[dept] = dept_ca_total.get(dept, 0) + num_ca
        dept_class_count[dept] = dept_class_count.get(dept, 0) + 1

        if dept not in dept_room_matrix:
            dept_room_matrix[dept] = {r: 0 for r in rooms}
        if room in dept_room_matrix[dept]:
            dept_room_matrix[dept][room] += num_ca

        if dept not in dept_day_session_matrix:
            dept_day_session_matrix[dept] = {(d, s): 0 for d in range(2, 8) for s in ('S', 'C')}

        for sl in slots:
            sess = 'S' if sl in (1, 2) else 'C'
            dept_day_session_matrix[dept][(dow, sess)] += 1
            day_session_totals[(dow, sess)] += 1
            day_ca_totals[dow] += 1
            if room in room_day_session:
                room_day_session[room][(dow, sess)] += 1

    # Sắp xếp danh sách đơn vị giảm dần theo tổng số ca sử dụng
    sorted_departments = sorted(
        dept_ca_total.keys(),
        key=lambda d: (-dept_ca_total[d], d)
    )

    max_system_ca = len(rooms) * MAX_CA_PER_ROOM_WEEK
    total_ca_free = max(0, max_system_ca - total_ca_used)
    utilization_rate = round((total_ca_used / max_system_ca) * 100, 2) if max_system_ca > 0 else 0.0
    free_rate = round(100.0 - utilization_rate, 2) if max_system_ca > 0 else 0.0

    peak_dow = max(day_ca_totals, key=day_ca_totals.get) if total_ca_used > 0 else 2
    peak_dow_ca = day_ca_totals.get(peak_dow, 0)
    peak_dow_pct = round((peak_dow_ca / (len(rooms) * MAX_CA_PER_ROOM_DAY)) * 100, 1) if total_ca_used > 0 else 0.0

    return {
        'rooms': rooms,
        'total_classes': total_classes,
        'total_ca_used': total_ca_used,
        'total_ca_free': total_ca_free,
        'max_system_ca': max_system_ca,
        'utilization_rate': utilization_rate,
        'free_rate': free_rate,
        'room_ca_used': room_ca_used,
        'room_class_count': room_class_count,
        'room_day_session': room_day_session,
        'sorted_departments': sorted_departments,
        'dept_ca_total': dept_ca_total,
        'dept_class_count': dept_class_count,
        'dept_room_matrix': dept_room_matrix,
        'dept_day_session_matrix': dept_day_session_matrix,
        'day_session_totals': day_session_totals,
        'day_ca_totals': day_ca_totals,
        'peak_dow': peak_dow,
        'peak_dow_ca': peak_dow_ca,
        'peak_dow_pct': peak_dow_pct,
    }


def build_semester_report_excel(
    all_schedules,
    issues_list,
    bookings_list,
    net_status_list,
    semester_mode='hk1',
    academic_year='',
    from_date_str='',
    to_date_str=''
):
    """
    Khởi tạo Workbook Excel (.xlsx) gồm 4 Sheet chuyên nghiệp theo chuẩn báo cáo
    của Phòng Quản trị Thiết bị - Trường ĐH Khoa học Tự nhiên, ĐHQG-HCM.
    """
    filter_res = filter_schedules_for_report(
        all_schedules=all_schedules,
        semester_mode=semester_mode,
        academic_year=academic_year,
        from_date_str=from_date_str,
        to_date_str=to_date_str
    )
    target_rows = filter_res['rows']
    semester_label = filter_res['semester_label']
    semester_short = filter_res['semester_short']
    academic_year = filter_res['academic_year']
    date_span_str = filter_res['date_span_str']

    stats = compute_semester_statistics(target_rows, ROOMS_LIST)
    today_str = datetime.date.today().strftime('%d/%m/%Y')

    wb = openpyxl.Workbook()
    default_ws = wb.active

    # =========================================================================
    # ĐỊNH NGHĨA BẢNG MÀU & STYLE CHUẨN CÔNG VỤ ĐẠI HỌC KHTN
    # =========================================================================
    FONT_NAME = "Times New Roman"

    f_org_bold = Font(name=FONT_NAME, size=11, bold=True, color="1E3A8A")
    f_org_sub = Font(name=FONT_NAME, size=11, bold=True, underline="single", color="1E293B")
    f_date_italic = Font(name=FONT_NAME, size=11, italic=True, color="334155")
    f_title = Font(name=FONT_NAME, size=15, bold=True, color="1E3A8A")
    f_subtitle = Font(name=FONT_NAME, size=11, italic=True, color="475569")
    f_sec_hdr = Font(name=FONT_NAME, size=12, bold=True, color="1E3A8A")

    f_tbl_hdr = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    f_tbl_subhdr = Font(name=FONT_NAME, size=10, bold=True, color="1E3A8A")
    f_data = Font(name=FONT_NAME, size=11, bold=False, color="0F172A")
    f_data_bold = Font(name=FONT_NAME, size=11, bold=True, color="0F172A")
    f_data_blue = Font(name=FONT_NAME, size=11, bold=True, color="1D4ED8")
    f_data_green = Font(name=FONT_NAME, size=11, bold=True, color="15803D")
    f_data_amber = Font(name=FONT_NAME, size=11, bold=True, color="B45309")
    f_data_red = Font(name=FONT_NAME, size=11, bold=True, color="B91C1C")
    f_total = Font(name=FONT_NAME, size=11, bold=True, color="1E3A8A")
    f_note = Font(name=FONT_NAME, size=11, italic=False, color="1E293B")

    fill_navy = PatternFill("solid", fgColor="1E3A8A")
    fill_blue_sub = PatternFill("solid", fgColor="DBEAFE")
    fill_zebra = PatternFill("solid", fgColor="F8FAFC")
    fill_white = PatternFill("solid", fgColor="FFFFFF")
    fill_total = PatternFill("solid", fgColor="E0E7FF")
    fill_soft_green = PatternFill("solid", fgColor="DCFCE7")
    fill_soft_amber = PatternFill("solid", fgColor="FEF3C7")
    fill_soft_red = PatternFill("solid", fgColor="FEE2E2")
    fill_section_banner = PatternFill("solid", fgColor="EFF6FF")

    thin_side = Side(border_style="thin", color="CBD5E1")
    med_navy = Side(border_style="medium", color="1E3A8A")
    double_bottom = Side(border_style="double", color="1E3A8A")

    b_grid = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    b_header = Border(left=thin_side, right=thin_side, top=med_navy, bottom=med_navy)
    b_total = Border(left=thin_side, right=thin_side, top=med_navy, bottom=double_bottom)
    b_section = Border(left=med_navy, right=thin_side, top=thin_side, bottom=thin_side)

    al_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def add_official_header(ws, max_col, sheet_title, sheet_subtitle):
        """Tạo phần Quốc hiệu, Tiêu ngữ và Tiêu đề báo cáo chuẩn hành chính."""
        ws.views.sheetView[0].showGridLines = True
        mid_col = max(3, max_col // 2)

        # Dòng 1 & 2: Quốc hiệu - Tiêu ngữ
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=mid_col)
        ws.cell(1, 1, "TRƯỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN").font = f_org_bold
        ws.cell(1, 1).alignment = al_center

        ws.merge_cells(start_row=1, start_column=mid_col + 1, end_row=1, end_column=max_col)
        ws.cell(1, mid_col + 1, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM").font = f_org_bold
        ws.cell(1, mid_col + 1).alignment = al_center

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=mid_col)
        ws.cell(2, 1, "PHÒNG QUẢN TRỊ THIẾT BỊ (CSLT)").font = f_org_sub
        ws.cell(2, 1).alignment = al_center

        ws.merge_cells(start_row=2, start_column=mid_col + 1, end_row=2, end_column=max_col)
        ws.cell(2, mid_col + 1, "Độc lập - Tự do - Hạnh phúc").font = f_org_sub
        ws.cell(2, mid_col + 1).alignment = al_center

        ws.merge_cells(start_row=3, start_column=mid_col + 1, end_row=3, end_column=max_col)
        now_dt = datetime.date.today()
        ws.cell(3, mid_col + 1, f"TP. Hồ Chí Minh, ngày {now_dt.day:02d} tháng {now_dt.month:02d} năm {now_dt.year}").font = f_date_italic
        ws.cell(3, mid_col + 1).alignment = al_center

        ws.row_dimensions[1].height = 20
        ws.row_dimensions[2].height = 20
        ws.row_dimensions[3].height = 18

        # Dòng 5 & 6: Tiêu đề lớn
        ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=max_col)
        ws.cell(5, 1, sheet_title).font = f_title
        ws.cell(5, 1).alignment = al_center
        ws.row_dimensions[5].height = 28

        ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=max_col)
        ws.cell(6, 1, sheet_subtitle).font = f_subtitle
        ws.cell(6, 1).alignment = al_center
        ws.row_dimensions[6].height = 20

    def add_section_banner(ws, row_idx, max_col, text):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=max_col)
        for c in range(1, max_col + 1):
            cell = ws.cell(row_idx, c)
            cell.fill = fill_section_banner
            cell.border = b_section if c == 1 else b_grid
        ws.cell(row_idx, 1, text).font = f_sec_hdr
        ws.cell(row_idx, 1).alignment = al_left
        ws.row_dimensions[row_idx].height = 25

    def style_range_border(ws, min_row, max_row, min_col, max_col, border_obj):
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(r, c).border = border_obj

    # =========================================================================
    # SHEET 1: 1. TỔNG QUAN HIỆU SUẤT (Bảng 3, Bảng 6 & Bảng 7 trong Word)
    # =========================================================================
    ws1 = wb.create_sheet(title="1. Tổng quan Hiệu suất")
    max_c1 = 6
    add_official_header(
        ws1,
        max_c1,
        f"BÁO CÁO TẦN SUẤT KHAI THÁC PHÒNG MÁY TÍNH — {semester_label} ({academic_year})",
        f"Cơ sở 2 (Linh Trung) • Giai đoạn môn học (BĐ – KT): {date_span_str} • Ngày xuất báo cáo: {today_str}"
    )

    # --- MỤC I: QUY CHUẨN CA HỌC & KHUNG GIỜ ---
    r_curr = 8
    add_section_banner(ws1, r_curr, max_c1, "I. QUY CHUẨN CA HỌC & CÔNG SUẤT THIẾT KẾ TẠI PHÒNG MÁY TÍNH")

    r_curr += 1
    ws1.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=max_c1)
    ws1.cell(
        r_curr, 1,
        "• Quy định: Mỗi ca học chuẩn tại phòng máy tính có thời lượng 2,5 tiết. Một ngày có 4 ca học (2 ca Sáng, 2 ca Chiều). "
        "Trong 1 tuần (Thứ 2 đến Thứ 7) mỗi phòng có tối đa 24 ca học và tổng 8 phòng máy cung cấp tối đa 192 ca học/tuần."
    ).font = f_note
    ws1.cell(r_curr, 1).alignment = al_left
    ws1.row_dimensions[r_curr].height = 24

    r_curr += 1
    slot_headers = ["Khung ca học", "Ca 1 (Buổi Sáng 1)", "Ca 2 (Buổi Sáng 2)", "Ca 3 (Buổi Chiều 1)", "Ca 4 (Buổi Chiều 2)", "Tổng công suất / ngày"]
    for idx, h_text in enumerate(slot_headers, 1):
        c = ws1.cell(r_curr, idx, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_navy
        c.alignment = al_center
        c.border = b_header
    ws1.row_dimensions[r_curr].height = 24

    r_curr += 1
    slot_periods = ["Tiết học chuẩn", "Tiết 1 – 3.5 (2,5 tiết)", "Tiết 3.5 – 5 (2,5 tiết)", "Tiết 6 – 8.5 (2,5 tiết)", "Tiết 8.5 – 10 (2,5 tiết)", "10 tiết / 4 ca chuẩn"]
    for idx, val in enumerate(slot_periods, 1):
        c = ws1.cell(r_curr, idx, val)
        c.font = f_data_bold if idx in (1, 6) else f_data
        c.fill = fill_blue_sub if idx == 1 else fill_white
        c.alignment = al_center
        c.border = b_grid
    ws1.row_dimensions[r_curr].height = 22

    r_curr += 1
    slot_times = ["Khung giờ học", "07h30 – 10h00", "10h05 – 12h00", "12h30 – 14h50", "15h05 – 17h00", "07h30 – 17h00"]
    for idx, val in enumerate(slot_times, 1):
        c = ws1.cell(r_curr, idx, val)
        c.font = f_data_bold if idx == 1 else f_data_blue
        c.fill = fill_blue_sub if idx == 1 else fill_zebra
        c.alignment = al_center
        c.border = b_grid
    ws1.row_dimensions[r_curr].height = 22

    # --- MỤC II: BẢNG TỔNG KẾT TỶ LỆ KHAI THÁC VÀ DƯ ĐỊA (BẢNG 6 MẪU WORD) ---
    r_curr += 2
    add_section_banner(ws1, r_curr, max_c1, f"II. BẢNG TỔNG KẾT TỶ LỆ KHAI THÁC VÀ DƯ ĐỊA CỦA PHÒNG MÁY TÍNH ({semester_label})")

    r_curr += 1
    b6_headers = ["STT", "Chỉ số thống kê khai thác", f"Số liệu {semester_short}", "Định mức tối đa", "Đơn vị tính", "Đánh giá thực trạng & Dư địa"]
    for idx, h_text in enumerate(b6_headers, 1):
        c = ws1.cell(r_curr, idx, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_navy
        c.alignment = al_center
        c.border = b_header
    ws1.row_dimensions[r_curr].height = 26

    util_eval = (
        "Đạt ngưỡng khai thác rất cao (gần tối đa công suất)"
        if stats['utilization_rate'] >= 80
        else ("Khai thác ở mức khá, cân đối" if stats['utilization_rate'] >= 60 else "Còn nhiều dư địa ca trống trong tuần")
    )
    free_eval = (
        f"Dư địa {stats['total_ca_free']} ca trống phục vụ đăng ký mượn phòng, dạy bù và bảo trì định kỳ"
        if stats['total_ca_free'] > 0
        else "Đã khai thác kín 100% công suất toàn hệ thống"
    )

    b6_rows = [
        (1, "Tổng số ca sử dụng trong tuần", stats['total_ca_used'], stats['max_system_ca'], "Ca / tuần", f"Tương đương {stats['total_classes']} lớp học phần trên Thời khóa biểu"),
        (2, "Tỷ lệ khai thác công suất", f"{stats['utilization_rate']:.2f}%", "100.00%", "% công suất", util_eval),
        (3, "Tổng số ca trống (Dư địa)", stats['total_ca_free'], 0, "Ca / tuần", free_eval),
        (4, "Tỷ lệ ca trống", f"{stats['free_rate']:.2f}%", "0.00%", "% công suất", f"Ngày cao điểm nhất: Thứ {stats['peak_dow']} ({stats['peak_dow_ca']}/32 ca - đạt {stats['peak_dow_pct']}%)"),
    ]

    for row_data in b6_rows:
        r_curr += 1
        is_even = (row_data[0] % 2 == 0)
        row_fill = fill_zebra if is_even else fill_white
        for c_idx, val in enumerate(row_data, 1):
            cell = ws1.cell(r_curr, c_idx, val)
            cell.fill = row_fill
            cell.border = b_grid
            if c_idx == 1:
                cell.alignment = al_center
                cell.font = f_data
            elif c_idx == 2:
                cell.alignment = al_left
                cell.font = f_data_bold
            elif c_idx == 3:
                cell.alignment = al_center
                cell.font = f_data_blue if row_data[0] in (1, 2) else f_data_amber
            elif c_idx in (4, 5):
                cell.alignment = al_center
                cell.font = f_data
            else:
                cell.alignment = al_left
                cell.font = f_data
        ws1.row_dimensions[r_curr].height = 24

    # --- MỤC III: BẢNG TỔNG KẾT SỐ CA CỦA MỖI ĐƠN VỊ (BẢNG 7 MẪU WORD) ---
    r_curr += 2
    add_section_banner(ws1, r_curr, max_c1, f"III. BẢNG TỔNG KẾT SỐ CA VÀ TỶ TRỌNG SỬ DỤNG CỦA MỖI ĐƠN VỊ ({semester_label})")

    r_curr += 1
    b7_headers = ["STT", "Đơn vị đào tạo (Khoa / Trung tâm)", "Số lớp TKB", "Tổng số ca / tuần", "Tỷ trọng sử dụng (%)", "Phân loại mức độ khai thác"]
    for idx, h_text in enumerate(b7_headers, 1):
        c = ws1.cell(r_curr, idx, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_navy
        c.alignment = al_center
        c.border = b_header
    ws1.row_dimensions[r_curr].height = 26

    for idx, dept in enumerate(stats['sorted_departments'], 1):
        r_curr += 1
        ca_cnt = stats['dept_ca_total'][dept]
        cls_cnt = stats['dept_class_count'][dept]
        pct = round((ca_cnt / stats['total_ca_used']) * 100, 2) if stats['total_ca_used'] > 0 else 0.0

        if pct >= 25.0:
            role_label = "Đơn vị sử dụng chủ lực (Trọng điểm)"
            role_font = f_data_blue
        elif pct >= 7.0:
            role_label = "Đơn vị sử dụng thường xuyên"
            role_font = f_data_green
        else:
            role_label = "Đơn vị sử dụng bổ sung"
            role_font = f_data

        row_fill = fill_zebra if idx % 2 == 0 else fill_white
        vals = [idx, dept, cls_cnt, ca_cnt, f"{pct:.2f}%", role_label]
        for c_idx, val in enumerate(vals, 1):
            cell = ws1.cell(r_curr, c_idx, val)
            cell.fill = row_fill
            cell.border = b_grid
            if c_idx in (1, 3, 4, 5):
                cell.alignment = al_center
                cell.font = f_data_bold if c_idx in (4, 5) else f_data
            elif c_idx == 2:
                cell.alignment = al_left
                cell.font = f_data_bold
            else:
                cell.alignment = al_left
                cell.font = role_font
        ws1.row_dimensions[r_curr].height = 23

    # Dòng Tổng cộng Bảng 7
    r_curr += 1
    tot_b7 = ["", "TỔNG CỘNG TOÀN TRƯỜNG", stats['total_classes'], stats['total_ca_used'], "100.00%", f"Đạt {stats['utilization_rate']:.2f}% tổng công suất 192 ca/tuần"]
    for c_idx, val in enumerate(tot_b7, 1):
        cell = ws1.cell(r_curr, c_idx, val)
        cell.fill = fill_total
        cell.border = b_total
        cell.font = f_total
        cell.alignment = al_center if c_idx in (1, 3, 4, 5) else al_left
    ws1.row_dimensions[r_curr].height = 26

    # --- MỤC IV: NHẬN XÉT & ĐÁNH GIÁ TỔNG HỢP TỰ ĐỘNG ---
    r_curr += 2
    add_section_banner(ws1, r_curr, max_c1, "IV. NHẬN XÉT & PHÂN TÍCH THỰC TRẠNG KHAI THÁC")

    full_rooms = [r for r in ROOMS_LIST if stats['room_ca_used'][r] >= 24]
    low_room = min(ROOMS_LIST, key=lambda r: stats['room_ca_used'][r])
    low_room_free = MAX_CA_PER_ROOM_WEEK - stats['room_ca_used'][low_room]
    top_2_depts = stats['sorted_departments'][:2]

    r_curr += 1
    ws1.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=max_c1)
    comment_1 = (
        f"1. Về hiệu suất khai thác: Trong {semester_short}, tổng số ca khai thác đạt {stats['total_ca_used']}/{stats['max_system_ca']} ca "
        f"(tỷ lệ lấp đầy đạt {stats['utilization_rate']:.2f}%), số ca còn trống là {stats['total_ca_free']} ca ({stats['free_rate']:.2f}%). "
        + (f"Đặc biệt các phòng {', '.join(full_rooms)} đã khai thác tối đa 100% công suất (24/24 ca/tuần). " if full_rooms else "")
        + f"Phòng có dư địa trống nhiều nhất là phòng {low_room} (trống {low_room_free}/24 ca)."
    )
    ws1.cell(r_curr, 1, comment_1).font = f_note
    ws1.cell(r_curr, 1).alignment = al_left
    ws1.row_dimensions[r_curr].height = 34

    r_curr += 1
    ws1.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=max_c1)
    if len(top_2_depts) >= 2:
        d1, d2 = top_2_depts[0], top_2_depts[1]
        c1_ca, c2_ca = stats['dept_ca_total'][d1], stats['dept_ca_total'][d2]
        c12_pct = round(((c1_ca + c2_ca) / max(1, stats['total_ca_used'])) * 100, 2)
        comment_2 = (
            f"2. Về đơn vị sử dụng chủ lực: Hai đơn vị dẫn đầu tần suất sử dụng là {d1} ({c1_ca} ca) và {d2} ({c2_ca} ca), "
            f"chiếm tổng cộng {c12_pct:.2f}% tổng thời lượng khai thác phòng máy toàn cơ sở."
        )
    else:
        comment_2 = "2. Về đơn vị sử dụng: Lịch học được phân bổ phục vụ các khoa và trung tâm theo đúng kế hoạch đào tạo."
    ws1.cell(r_curr, 1, comment_2).font = f_note
    ws1.cell(r_curr, 1).alignment = al_left
    ws1.row_dimensions[r_curr].height = 26

    # Độ rộng cột Sheet 1
    col_widths_s1 = [8, 34, 20, 20, 18, 42]
    for i, w in enumerate(col_widths_s1, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # =========================================================================
    # SHEET 2: 2. TẦN SUẤT THEO PHÒNG (Bảng 1 & Bảng 2/3 trong Word)
    # =========================================================================
    ws2 = wb.create_sheet(title="2. Tần suất theo Phòng")
    max_c2 = 11  # STT + Đơn vị + 8 Phòng + TỔNG
    add_official_header(
        ws2,
        max_c2,
        f"THỐNG KÊ TẦN SUẤT SỬ DỤNG THEO PHÒNG MÁY & ĐƠN VỊ — {semester_label}",
        f"Định mức chuẩn: 24 ca / phòng / tuần • Tổng 8 phòng máy = 192 ca / tuần • Năm học {academic_year}"
    )

    # --- BẢNG 1: THỐNG KÊ SỬ DỤNG PHÒNG MÁY TÍNH (BẢNG 1 MẪU WORD) ---
    r2 = 8
    add_section_banner(ws2, r2, max_c2, f"I. BẢNG 1: THỐNG KÊ TẦN SUẤT SỬ DỤNG VÀ SỐ CA TRỐNG TẠI MỖI PHÒNG MÁY ({semester_label})")

    r2 += 1
    # Trình bày gói gọn, rõ ràng trên lưới 11 cột
    ws2.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=3)
    ws2.merge_cells(start_row=r2, start_column=4, end_row=r2, end_column=5)
    ws2.merge_cells(start_row=r2, start_column=9, end_row=r2, end_column=11)

    b1_hdr_map = {
        1: "STT",
        2: "Phòng máy tính",
        4: "Tổng số ca / tuần\n(Chuẩn mẫu báo cáo)",
        6: "Số ca sử dụng",
        7: "Số ca trống",
        8: "Tỷ lệ khai thác (%)",
        9: "Đánh giá tình trạng khai thác & Dư địa"
    }
    for c_i in range(1, max_c2 + 1):
        cell = ws2.cell(r2, c_i)
        if c_i in b1_hdr_map:
            cell.value = b1_hdr_map[c_i]
        cell.font = f_tbl_hdr
        cell.fill = fill_navy
        cell.alignment = al_center
        cell.border = b_header
    ws2.row_dimensions[r2].height = 28

    for idx, rm in enumerate(ROOMS_LIST, 1):
        r2 += 1
        used = stats['room_ca_used'][rm]
        free = max(0, MAX_CA_PER_ROOM_WEEK - used)
        pct = round((used / MAX_CA_PER_ROOM_WEEK) * 100, 2)
        ratio_str = f"{used}/{MAX_CA_PER_ROOM_WEEK}"

        if used >= 24:
            status_txt = "Khai thác tối đa 100% công suất (Kín lịch cả tuần)"
            st_font = f_data_red
            badge_fill = fill_soft_red
        elif used >= 20:
            status_txt = f"Cường độ khai thác rất cao (Chỉ còn trống {free} ca)"
            st_font = f_data_amber
            badge_fill = fill_soft_amber
        elif used >= 15:
            status_txt = f"Vận hành ổn định, cân đối (Còn trống {free} ca)"
            st_font = f_data_green
            badge_fill = fill_soft_green
        else:
            status_txt = f"Còn dư địa cao ({free}/24 ca trống - Ưu tiên xếp lịch mượn phòng)"
            st_font = f_data_blue
            badge_fill = fill_blue_sub

        ws2.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=3)
        ws2.merge_cells(start_row=r2, start_column=4, end_row=r2, end_column=5)
        ws2.merge_cells(start_row=r2, start_column=9, end_row=r2, end_column=11)

        row_fill = fill_zebra if idx % 2 == 0 else fill_white
        for c_i in range(1, max_c2 + 1):
            c = ws2.cell(r2, c_i)
            c.fill = row_fill
            c.border = b_grid

        ws2.cell(r2, 1, idx).alignment = al_center
        ws2.cell(r2, 1).font = f_data

        ws2.cell(r2, 2, f"Phòng máy {rm}").alignment = al_center
        ws2.cell(r2, 2).font = f_data_bold

        ws2.cell(r2, 4, ratio_str).alignment = al_center
        ws2.cell(r2, 4).font = f_data_blue
        ws2.cell(r2, 4).fill = badge_fill
        ws2.cell(r2, 5).fill = badge_fill

        ws2.cell(r2, 6, used).alignment = al_center
        ws2.cell(r2, 6).font = f_data_bold

        ws2.cell(r2, 7, free).alignment = al_center
        ws2.cell(r2, 7).font = f_data_bold if free > 0 else f_data

        ws2.cell(r2, 8, f"{pct:.2f}%").alignment = al_center
        ws2.cell(r2, 8).font = st_font

        ws2.cell(r2, 9, status_txt).alignment = al_left
        ws2.cell(r2, 9).font = st_font
        ws2.row_dimensions[r2].height = 23

    # Dòng Tổng Bảng 1
    r2 += 1
    ws2.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=3)
    ws2.merge_cells(start_row=r2, start_column=4, end_row=r2, end_column=5)
    ws2.merge_cells(start_row=r2, start_column=9, end_row=r2, end_column=11)
    for c_i in range(1, max_c2 + 1):
        c = ws2.cell(r2, c_i)
        c.fill = fill_total
        c.border = b_total
        c.font = f_total
        c.alignment = al_center

    ws2.cell(r2, 2, "TỔNG CỘNG (8 PHÒNG)")
    ws2.cell(r2, 4, f"{stats['total_ca_used']}/{stats['max_system_ca']}")
    ws2.cell(r2, 6, stats['total_ca_used'])
    ws2.cell(r2, 7, stats['total_ca_free'])
    ws2.cell(r2, 8, f"{stats['utilization_rate']:.2f}%")
    ws2.cell(r2, 9, f"Tổng sử dụng {stats['total_ca_used']} ca • Còn trống {stats['total_ca_free']} ca ({stats['free_rate']:.2f}%)").alignment = al_left
    ws2.row_dimensions[r2].height = 26

    # --- BẢNG 2: MA TRẬN SỐ CA PHÂN BỔ THEO PHÒNG CỦA MỖI ĐƠN VỊ (BẢNG 2 MẪU WORD) ---
    r2 += 2
    add_section_banner(ws2, r2, max_c2, f"II. BẢNG 2: SỐ CA PHÂN BỔ THEO CÁC PHÒNG MÁY TÍNH CỦA MỖI ĐƠN VỊ ({semester_label})")

    r2 += 1
    b2_headers = ["STT", "Đơn vị"] + [f"P.{rm}" for rm in ROOMS_LIST] + ["TỔNG"]
    for idx, h_text in enumerate(b2_headers, 1):
        c = ws2.cell(r2, idx, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_navy
        c.alignment = al_center
        c.border = b_header
    ws2.row_dimensions[r2].height = 26

    for idx, dept in enumerate(stats['sorted_departments'], 1):
        r2 += 1
        row_fill = fill_zebra if idx % 2 == 0 else fill_white
        ws2.cell(r2, 1, idx).alignment = al_center
        ws2.cell(r2, 1).font = f_data
        ws2.cell(r2, 1).fill = row_fill
        ws2.cell(r2, 1).border = b_grid

        ws2.cell(r2, 2, dept).alignment = al_left
        ws2.cell(r2, 2).font = f_data_bold
        ws2.cell(r2, 2).fill = row_fill
        ws2.cell(r2, 2).border = b_grid

        for r_i, rm in enumerate(ROOMS_LIST, 3):
            val = stats['dept_room_matrix'][dept].get(rm, 0)
            # Để trống ô có giá trị 0 giống bảng mẫu Word để dễ quan sát
            c = ws2.cell(r2, r_i, val if val > 0 else "")
            c.alignment = al_center
            c.font = f_data_blue if val >= 10 else f_data
            c.fill = fill_blue_sub if val >= 18 else row_fill
            c.border = b_grid

        tot_dept = stats['dept_ca_total'][dept]
        c_tot = ws2.cell(r2, max_c2, tot_dept)
        c_tot.alignment = al_center
        c_tot.font = f_data_bold
        c_tot.fill = fill_total
        c_tot.border = b_grid
        ws2.row_dimensions[r2].height = 23

    # Dòng TỔNG CỘNG Bảng 2
    r2 += 1
    ws2.cell(r2, 1, "").fill = fill_total
    ws2.cell(r2, 1).border = b_total
    ws2.cell(r2, 2, "TỔNG SỐ CA SỬ DỤNG").alignment = al_left
    ws2.cell(r2, 2).font = f_total
    ws2.cell(r2, 2).fill = fill_total
    ws2.cell(r2, 2).border = b_total

    for r_i, rm in enumerate(ROOMS_LIST, 3):
        c = ws2.cell(r2, r_i, stats['room_ca_used'][rm])
        c.alignment = al_center
        c.font = f_total
        c.fill = fill_total
        c.border = b_total

    c_grand = ws2.cell(r2, max_c2, stats['total_ca_used'])
    c_grand.alignment = al_center
    c_grand.font = f_total
    c_grand.fill = fill_total
    c_grand.border = b_total
    ws2.row_dimensions[r2].height = 25

    # Dòng SỐ CA TRỐNG (Dư địa từng phòng)
    r2 += 1
    ws2.cell(r2, 1, "").fill = fill_soft_amber
    ws2.cell(r2, 1).border = b_grid
    ws2.cell(r2, 2, "SỐ CA TRỐNG (Trên 24 ca/phòng)").alignment = al_left
    ws2.cell(r2, 2).font = f_data_amber
    ws2.cell(r2, 2).fill = fill_soft_amber
    ws2.cell(r2, 2).border = b_grid

    for r_i, rm in enumerate(ROOMS_LIST, 3):
        free_rm = max(0, MAX_CA_PER_ROOM_WEEK - stats['room_ca_used'][rm])
        c = ws2.cell(r2, r_i, free_rm)
        c.alignment = al_center
        c.font = f_data_amber if free_rm > 0 else f_data_green
        c.fill = fill_soft_amber if free_rm > 0 else fill_soft_green
        c.border = b_grid

    c_free_tot = ws2.cell(r2, max_c2, stats['total_ca_free'])
    c_free_tot.alignment = al_center
    c_free_tot.font = f_data_amber
    c_free_tot.fill = fill_soft_amber
    c_free_tot.border = b_grid
    ws2.row_dimensions[r2].height = 23

    col_widths_s2 = [6, 28, 10, 10, 10, 10, 10, 10, 10, 10, 14]
    for i, w in enumerate(col_widths_s2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # =========================================================================
    # SHEET 3: 3. PHÂN BỔ BUỔI (SÁNG-CHIỀU) (Bảng 4 & Bảng 5 trong Word)
    # =========================================================================
    ws3 = wb.create_sheet(title="3. Phân bổ Buổi (Sáng-Chiều)")
    max_c3 = 15  # STT + Đơn vị + 12 cột (Thứ 2 S/C -> Thứ 7 S/C) + TỔNG
    add_official_header(
        ws3,
        max_c3,
        f"MA TRẬN SỐ CA PHÂN BỔ THEO BUỔI (SÁNG / CHIỀU) TRONG TUẦN — {semester_label}",
        f"Ký hiệu: S = Buổi Sáng (Ca 1, Ca 2) • C = Buổi Chiều (Ca 3, Ca 4) • Công suất tối đa mỗi buổi (8 phòng × 2 ca) = 16 ca"
    )

    # --- BẢNG 1: SỐ CA PHÂN BỔ THEO BUỔI CỦA MỖI ĐƠN VỊ (BẢNG 4 MẪU WORD) ---
    r3 = 8
    add_section_banner(ws3, r3, max_c3, f"I. BẢNG 4: THỐNG KÊ SỐ CA PHÂN BỔ THEO BUỔI (SÁNG / CHIỀU) CỦA MỖI ĐƠN VỊ ({semester_label})")

    r3_hdr1 = r3 + 1
    r3_hdr2 = r3 + 2

    ws3.merge_cells(start_row=r3_hdr1, start_column=1, end_row=r3_hdr2, end_column=1)
    ws3.cell(r3_hdr1, 1, "STT")

    ws3.merge_cells(start_row=r3_hdr1, start_column=2, end_row=r3_hdr2, end_column=2)
    ws3.cell(r3_hdr1, 2, "Đơn vị")

    for d_idx, dow in enumerate(range(2, 8)):
        col_s = 3 + d_idx * 2
        col_c = col_s + 1
        ws3.merge_cells(start_row=r3_hdr1, start_column=col_s, end_row=r3_hdr1, end_column=col_c)
        ws3.cell(r3_hdr1, col_s, f"Thứ {dow}")
        ws3.cell(r3_hdr2, col_s, "S")
        ws3.cell(r3_hdr2, col_c, "C")

    ws3.merge_cells(start_row=r3_hdr1, start_column=max_c3, end_row=r3_hdr2, end_column=max_c3)
    ws3.cell(r3_hdr1, max_c3, "TỔNG")

    for r_h in (r3_hdr1, r3_hdr2):
        for c_i in range(1, max_c3 + 1):
            cell = ws3.cell(r_h, c_i)
            cell.font = f_tbl_hdr if r_h == r3_hdr1 or c_i in (1, 2, max_c3) else f_tbl_subhdr
            cell.fill = fill_navy if r_h == r3_hdr1 or c_i in (1, 2, max_c3) else fill_blue_sub
            cell.alignment = al_center
            cell.border = b_header
        ws3.row_dimensions[r_h].height = 22

    r3 = r3_hdr2
    for idx, dept in enumerate(stats['sorted_departments'], 1):
        r3 += 1
        row_fill = fill_zebra if idx % 2 == 0 else fill_white

        c_stt = ws3.cell(r3, 1, idx)
        c_stt.alignment = al_center
        c_stt.font = f_data
        c_stt.fill = row_fill
        c_stt.border = b_grid

        c_dept = ws3.cell(r3, 2, dept)
        c_dept.alignment = al_left
        c_dept.font = f_data_bold
        c_dept.fill = row_fill
        c_dept.border = b_grid

        for d_idx, dow in enumerate(range(2, 8)):
            col_s = 3 + d_idx * 2
            col_c = col_s + 1
            val_s = stats['dept_day_session_matrix'][dept].get((dow, 'S'), 0)
            val_c = stats['dept_day_session_matrix'][dept].get((dow, 'C'), 0)

            cs = ws3.cell(r3, col_s, val_s if val_s > 0 else "")
            cs.alignment = al_center
            cs.font = f_data_blue if val_s >= 6 else f_data
            cs.fill = row_fill
            cs.border = b_grid

            cc = ws3.cell(r3, col_c, val_c if val_c > 0 else "")
            cc.alignment = al_center
            cc.font = f_data_blue if val_c >= 6 else f_data
            cc.fill = row_fill
            cc.border = b_grid

        c_tot = ws3.cell(r3, max_c3, stats['dept_ca_total'][dept])
        c_tot.alignment = al_center
        c_tot.font = f_data_bold
        c_tot.fill = fill_total
        c_tot.border = b_grid
        ws3.row_dimensions[r3].height = 23

    # Dòng TỔNG SỐ CA SỬ DỤNG THEO BUỔI
    r3 += 1
    ws3.cell(r3, 1, "").fill = fill_total
    ws3.cell(r3, 1).border = b_total
    ws3.cell(r3, 2, "TỔNG SỐ CA SỬ DỤNG").alignment = al_left
    ws3.cell(r3, 2).font = f_total
    ws3.cell(r3, 2).fill = fill_total
    ws3.cell(r3, 2).border = b_total

    for d_idx, dow in enumerate(range(2, 8)):
        col_s = 3 + d_idx * 2
        col_c = col_s + 1
        tot_s = stats['day_session_totals'][(dow, 'S')]
        tot_c = stats['day_session_totals'][(dow, 'C')]

        cs = ws3.cell(r3, col_s, tot_s)
        cs.alignment = al_center
        cs.font = f_total
        cs.fill = fill_total
        cs.border = b_total

        cc = ws3.cell(r3, col_c, tot_c)
        cc.alignment = al_center
        cc.font = f_total
        cc.fill = fill_total
        cc.border = b_total

    c_grand3 = ws3.cell(r3, max_c3, stats['total_ca_used'])
    c_grand3.alignment = al_center
    c_grand3.font = f_total
    c_grand3.fill = fill_total
    c_grand3.border = b_total
    ws3.row_dimensions[r3].height = 25

    # Dòng SỐ CA TRỐNG THEO BUỔI (trên 16 ca/buổi)
    r3 += 1
    ws3.cell(r3, 1, "").fill = fill_soft_amber
    ws3.cell(r3, 1).border = b_grid
    ws3.cell(r3, 2, "SỐ CA TRỐNG (Tối đa 16 ca/buổi)").alignment = al_left
    ws3.cell(r3, 2).font = f_data_amber
    ws3.cell(r3, 2).fill = fill_soft_amber
    ws3.cell(r3, 2).border = b_grid

    for d_idx, dow in enumerate(range(2, 8)):
        col_s = 3 + d_idx * 2
        col_c = col_s + 1
        free_s = max(0, MAX_CA_SYSTEM_SESSION - stats['day_session_totals'][(dow, 'S')])
        free_c = max(0, MAX_CA_SYSTEM_SESSION - stats['day_session_totals'][(dow, 'C')])

        cs = ws3.cell(r3, col_s, free_s)
        cs.alignment = al_center
        cs.font = f_data_amber if free_s > 0 else f_data_green
        cs.fill = fill_soft_amber if free_s > 0 else fill_soft_green
        cs.border = b_grid

        cc = ws3.cell(r3, col_c, free_c)
        cc.alignment = al_center
        cc.font = f_data_amber if free_c > 0 else f_data_green
        cc.fill = fill_soft_amber if free_c > 0 else fill_soft_green
        cc.border = b_grid

    c_free_all = ws3.cell(r3, max_c3, stats['total_ca_free'])
    c_free_all.alignment = al_center
    c_free_all.font = f_data_amber
    c_free_all.fill = fill_soft_amber
    c_free_all.border = b_grid
    ws3.row_dimensions[r3].height = 23

    # --- BẢNG 2: MA TRẬN LẤP ĐẦY PHÒNG MÁY × BUỔI TRONG TUẦN (Tối đa 2 ca / phòng / buổi) ---
    r3 += 2
    add_section_banner(ws3, r3, max_c3, "II. MA TRẬN CHI TIẾT SỐ CA HOẠT ĐỘNG CỦA TỪNG PHÒNG THEO BUỔI (Tối đa 2 ca / phòng / buổi)")

    r3_rm_h1 = r3 + 1
    r3_rm_h2 = r3 + 2
    ws3.merge_cells(start_row=r3_rm_h1, start_column=1, end_row=r3_rm_h2, end_column=1)
    ws3.cell(r3_rm_h1, 1, "STT")
    ws3.merge_cells(start_row=r3_rm_h1, start_column=2, end_row=r3_rm_h2, end_column=2)
    ws3.cell(r3_rm_h1, 2, "Phòng máy")

    for d_idx, dow in enumerate(range(2, 8)):
        col_s = 3 + d_idx * 2
        col_c = col_s + 1
        ws3.merge_cells(start_row=r3_rm_h1, start_column=col_s, end_row=r3_rm_h1, end_column=col_c)
        ws3.cell(r3_rm_h1, col_s, f"Thứ {dow}")
        ws3.cell(r3_rm_h2, col_s, "S")
        ws3.cell(r3_rm_h2, col_c, "C")

    ws3.merge_cells(start_row=r3_rm_h1, start_column=max_c3, end_row=r3_rm_h2, end_column=max_c3)
    ws3.cell(r3_rm_h1, max_c3, "Tổng ca/tuần")

    for r_h in (r3_rm_h1, r3_rm_h2):
        for c_i in range(1, max_c3 + 1):
            cell = ws3.cell(r_h, c_i)
            cell.font = f_tbl_hdr if r_h == r3_rm_h1 or c_i in (1, 2, max_c3) else f_tbl_subhdr
            cell.fill = fill_navy if r_h == r3_rm_h1 or c_i in (1, 2, max_c3) else fill_blue_sub
            cell.alignment = al_center
            cell.border = b_header
        ws3.row_dimensions[r_h].height = 22

    r3 = r3_rm_h2
    for idx, rm in enumerate(ROOMS_LIST, 1):
        r3 += 1
        row_fill = fill_zebra if idx % 2 == 0 else fill_white
        ws3.cell(r3, 1, idx).alignment = al_center
        ws3.cell(r3, 1).font = f_data
        ws3.cell(r3, 1).fill = row_fill
        ws3.cell(r3, 1).border = b_grid

        ws3.cell(r3, 2, f"Phòng {rm}").alignment = al_center
        ws3.cell(r3, 2).font = f_data_bold
        ws3.cell(r3, 2).fill = row_fill
        ws3.cell(r3, 2).border = b_grid

        for d_idx, dow in enumerate(range(2, 8)):
            col_s = 3 + d_idx * 2
            col_c = col_s + 1
            vs = stats['room_day_session'][rm].get((dow, 'S'), 0)
            vc = stats['room_day_session'][rm].get((dow, 'C'), 0)

            cs = ws3.cell(r3, col_s, vs if vs > 0 else "")
            cs.alignment = al_center
            cs.font = f_data_green if vs == 2 else (f_data_amber if vs == 1 else f_data)
            cs.fill = fill_soft_green if vs == 2 else row_fill
            cs.border = b_grid

            cc = ws3.cell(r3, col_c, vc if vc > 0 else "")
            cc.alignment = al_center
            cc.font = f_data_green if vc == 2 else (f_data_amber if vc == 1 else f_data)
            cc.fill = fill_soft_green if vc == 2 else row_fill
            cc.border = b_grid

        c_rm_tot = ws3.cell(r3, max_c3, f"{stats['room_ca_used'][rm]}/24")
        c_rm_tot.alignment = al_center
        c_rm_tot.font = f_data_blue
        c_rm_tot.fill = fill_total
        c_rm_tot.border = b_grid
        ws3.row_dimensions[r3].height = 22

    col_widths_s3 = [6, 26] + [6.5] * 12 + [14]
    for i, w in enumerate(col_widths_s3, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # =========================================================================
    # SHEET 4: 4. VẬN HÀNH & KỸ THUẬT (Hạ tầng, Sự cố & Mượn phòng/Phần mềm)
    # =========================================================================
    ws4 = wb.create_sheet(title="4. Vận hành & Kỹ thuật")
    max_c4 = 8
    add_official_header(
        ws4,
        max_c4,
        f"BÁO CÁO CÔNG TÁC VẬN HÀNH, HẠ TẦNG KỸ THUẬT & SỰ CỐ — {semester_label}",
        f"Tổng hợp tình trạng thiết bị, mạng LAN/Internet (B42Lab NetAgent), nhật ký sự cố và cài đặt phần mềm ({academic_year})"
    )

    # --- BẢNG 1: HẠ TẦNG PHÒNG MÁY & TRẠNG THÁI VẬN HÀNH ---
    r4 = 8
    add_section_banner(ws4, r4, max_c4, "I. THỐNG KÊ HẠ TẦNG THIẾT BỊ, ĐƯỜNG TRUYỀN MẠNG & ĐỘ ỔN ĐỊNH 08 PHÒNG MÁY")

    r4 += 1
    infra_headers = [
        "STT", "Tên phòng máy", "Quy mô máy tính", "Cấu hình / Model chuẩn",
        "IP Router / Gateway", "Trạng thái Mạng", "Tần suất khai thác", "Đánh giá độ ổn định vận hành"
    ]
    for idx, h_text in enumerate(infra_headers, 1):
        c = ws4.cell(r4, idx, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_navy
        c.alignment = al_center
        c.border = b_header
    ws4.row_dimensions[r4].height = 26

    net_map = {str(n['room']): dict(n) for n in net_status_list} if net_status_list else {}
    open_issues_by_room = {r: 0 for r in ROOMS_LIST}
    total_issues_by_room = {r: 0 for r in ROOMS_LIST}
    for iss in (issues_list or []):
        rm_code = str(iss['room'] or '').strip()
        if rm_code in total_issues_by_room:
            total_issues_by_room[rm_code] += 1
            if (iss['status'] or 'Pending') != 'Resolved':
                open_issues_by_room[rm_code] += 1

    for idx, rm in enumerate(ROOMS_LIST, 1):
        r4 += 1
        row_fill = fill_zebra if idx % 2 == 0 else fill_white
        net_info = net_map.get(rm, {})
        ip_str = net_info.get('target_ip') or '10.29.37.xx'
        net_st = net_info.get('status') or 'online'
        ping_ms = net_info.get('ping_ms') or 12
        net_label = f"Online ({ping_ms}ms)" if net_st == 'online' else "Offline / Mất kết nối"

        open_cnt = open_issues_by_room.get(rm, 0)
        tot_cnt = total_issues_by_room.get(rm, 0)
        if open_cnt == 0:
            stability = "Máy vận hành ổn định, mạng LAN/Wi-Fi tốt"
            st_font = f_data_green
        else:
            stability = f"Đang xử lý {open_cnt} báo cáo sự cố thiết bị (Tổng {tot_cnt} phiếu)"
            st_font = f_data_amber

        row_vals = [
            idx,
            f"Phòng máy tính {rm}",
            "40 + 01 máy GV",
            "Đồng bộ PC nâng cấp 2025 (Windows 11)",
            ip_str,
            net_label,
            f"{stats['room_ca_used'][rm]}/24 ca/tuần",
            stability
        ]
        for c_i, val in enumerate(row_vals, 1):
            c = ws4.cell(r4, c_i, val)
            c.fill = row_fill
            c.border = b_grid
            if c_i in (1, 3, 5, 6, 7):
                c.alignment = al_center
                c.font = f_data_green if (c_i == 6 and net_st == 'online') else (f_data_red if c_i == 6 else f_data)
            elif c_i == 2:
                c.alignment = al_center
                c.font = f_data_bold
            elif c_i == 8:
                c.alignment = al_left
                c.font = st_font
            else:
                c.alignment = al_left
                c.font = f_data
        ws4.row_dimensions[r4].height = 23

    # --- BẢNG 2: NHẬT KÝ & TIẾN ĐỘ XỬ LÝ SỰ CỐ KỸ THUẬT ---
    r4 += 2
    add_section_banner(ws4, r4, max_c4, "II. THỐNG KÊ NHẬT KÝ BÁO CÁO SỰ CỐ KỸ THUẬT & TIẾN ĐỘ KHẮC PHỤC")

    r4 += 1
    iss_headers = ["STT", "Phòng", "Vị trí máy", "Hạng mục sự cố", "Mô tả chi tiết tình trạng", "Người báo / Đơn vị", "Thời gian ghi nhận", "Trạng thái xử lý"]
    for idx, h_text in enumerate(iss_headers, 1):
        c = ws4.cell(r4, idx, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_navy
        c.alignment = al_center
        c.border = b_header
    ws4.row_dimensions[r4].height = 26

    if issues_list:
        for idx, iss in enumerate(issues_list, 1):
            r4 += 1
            row_fill = fill_zebra if idx % 2 == 0 else fill_white
            raw_st = iss['status'] or 'Pending'
            if raw_st in ('Resolved', 'Approved', 'Đã khắc phục', 'Đã duyệt'):
                st_label = "Đã khắc phục hoàn tất"
                st_font = f_data_green
                st_fill = fill_soft_green
            elif raw_st in ('Processing', 'Đã tiếp nhận'):
                st_label = "Đang tiếp nhận xử lý"
                st_font = f_data_blue
                st_fill = fill_blue_sub
            else:
                st_label = "Chờ kiểm tra kỹ thuật"
                st_font = f_data_amber
                st_fill = fill_soft_amber

            rep_unit = f"{iss['reporter_name'] or ''} ({iss['department'] or ''})".strip()
            vals = [
                idx,
                f"P.{iss['room']}",
                iss['machine_number'] or 'Chung',
                iss['issue_category'] or 'Thiết bị',
                iss['description'] or '',
                rep_unit,
                iss['created_at'] or '',
                st_label
            ]
            for c_i, val in enumerate(vals, 1):
                c = ws4.cell(r4, c_i, val)
                c.fill = st_fill if c_i == 8 else row_fill
                c.border = b_grid
                if c_i in (1, 2, 3, 7, 8):
                    c.alignment = al_center
                    c.font = st_font if c_i == 8 else (f_data_bold if c_i == 2 else f_data)
                else:
                    c.alignment = al_left
                    c.font = f_data
            ws4.row_dimensions[r4].height = 24
    else:
        r4 += 1
        ws4.merge_cells(start_row=r4, start_column=1, end_row=r4, end_column=max_c4)
        style_range_border(ws4, r4, r4, 1, max_c4, b_grid)
        c_empty = ws4.cell(r4, 1, "✅ Hệ thống vận hành ổn định — Hiện không có sự cố kỹ thuật nào tồn đọng trong kỳ.")
        c_empty.font = f_data_green
        c_empty.alignment = al_center
        c_empty.fill = fill_soft_green
        ws4.row_dimensions[r4].height = 26

    # --- BẢNG 3: TỔNG HỢP ĐĂNG KÝ MƯỢN PHÒNG / DẠY BÙ & YÊU CẦU PHẦN MỀM ---
    r4 += 2
    add_section_banner(ws4, r4, max_c4, "III. TỔNG HỢP YÊU CẦU MƯỢN PHÒNG / DẠY BÙ & CÀI ĐẶT PHẦN MỀM CHUYÊN NGÀNH")

    r4 += 1
    bk_headers = ["STT", "Phòng", "Thứ / Tiết", "Môn học / Mục đích", "Giảng viên / Đơn vị", "Phần mềm yêu cầu cài đặt", "Ngày bắt đầu", "Trạng thái duyệt"]
    for idx, h_text in enumerate(bk_headers, 1):
        c = ws4.cell(r4, idx, h_text)
        c.font = f_tbl_hdr
        c.fill = fill_navy
        c.alignment = al_center
        c.border = b_header
    ws4.row_dimensions[r4].height = 26

    if bookings_list:
        for idx, bk in enumerate(bookings_list, 1):
            r4 += 1
            row_fill = fill_zebra if idx % 2 == 0 else fill_white
            raw_st = bk['status'] or 'Pending'
            if raw_st == 'Approved':
                st_txt = "Đã duyệt xếp phòng"
                st_f = f_data_green
            elif raw_st == 'Processing':
                st_txt = "Đã tiếp nhận"
                st_f = f_data_blue
            elif raw_st == 'Rejected':
                st_txt = "Từ chối"
                st_f = f_data_red
            else:
                st_txt = "Chờ duyệt"
                st_f = f_data_amber

            subj_txt = bk['subject_name'] or bk['purpose'] or 'Mượn phòng thực hành'
            if bk['class_code']:
                subj_txt += f" ({bk['class_code']})"
            req_who = f"{bk['requester_name'] or ''} ({bk['department'] or ''})".strip()
            sw_txt = bk['software_required'] or 'Phần mềm tiêu chuẩn phòng máy'

            vals = [
                idx,
                f"P.{bk['room']}",
                f"Thứ {bk['day_of_week']} (T.{bk['start_period']:g}-{bk['end_period']:g})",
                subj_txt,
                req_who,
                sw_txt,
                bk['study_period'] or '',
                st_txt
            ]
            for c_i, val in enumerate(vals, 1):
                c = ws4.cell(r4, c_i, val)
                c.fill = row_fill
                c.border = b_grid
                if c_i in (1, 2, 3, 7, 8):
                    c.alignment = al_center
                    c.font = st_f if c_i == 8 else (f_data_bold if c_i == 2 else f_data)
                else:
                    c.alignment = al_left
                    c.font = f_data_blue if (c_i == 6 and bk['software_required']) else f_data
            ws4.row_dimensions[r4].height = 24
    else:
        r4 += 1
        ws4.merge_cells(start_row=r4, start_column=1, end_row=r4, end_column=max_c4)
        style_range_border(ws4, r4, r4, 1, max_c4, b_grid)
        c_empty_bk = ws4.cell(r4, 1, "Chưa có yêu cầu đăng ký mượn phòng phát sinh ngoài Thời khóa biểu chính khóa.")
        c_empty_bk.font = f_data
        c_empty_bk.alignment = al_center
        c_empty_bk.fill = fill_zebra
        ws4.row_dimensions[r4].height = 25

    col_widths_s4 = [6, 18, 18, 30, 26, 26, 18, 34]
    for i, w in enumerate(col_widths_s4, 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

    # Thiết lập in ấn khổ ngang (Landscape) tự co vừa 1 trang ngang cho cả 4 Sheet
    for sheet in (ws1, ws2, ws3, ws4):
        sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0

    wb.remove(default_ws)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    if semester_mode == 'custom':
        safe_sem = "Giai_Doan"
    else:
        safe_sem = "HK2" if (semester_mode == 'hk2' or "II" in semester_label) else "HK1"
    safe_yr = academic_year.replace(" ", "")
    filename = f"Bao_Cao_Khai_Thac_Phong_May_{safe_sem}_{safe_yr}.xlsx"
    return output, filename
