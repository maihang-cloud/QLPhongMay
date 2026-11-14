"""Module xuất Báo cáo Tích hợp Hàng tháng (All-in-One Monthly Report) cho Hệ thống Quản lý Phòng máy.
Hỗ trợ 2 định dạng:
1. File Excel (.xlsx) 4 Sheet chi tiết: Dashboard KPI, Phân bổ TKB 8 phòng, Xử lý yêu cầu User, Nhật ký Admin.
2. File Word (.docx) Báo cáo tóm tắt chuẩn thể thức văn bản hành chính (Tiêu ngữ, Header giống mẫu Phòng QTTB CSLT).
"""

import io
import html
import zipfile
import calendar
import datetime
from timezone_utils import get_vn_today
from report_exporter import (
    ROOMS_LIST,
    parse_report_date as parse_date_safe,
    get_schedule_ca_slots
)
import database

# Openpyxl cho Excel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Python-docx cho Word (Hỗ trợ fallback tự động bằng OpenXML/zipfile nếu server chưa cài python-docx)
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    HAS_PYTHON_DOCX = True
except ImportError:
    HAS_PYTHON_DOCX = False

ROOM_MACHINES = {r: 41 for r in ROOMS_LIST}  # 40 máy SV + 01 máy GV = 41 máy/phòng
TOTAL_MACHINES = sum(ROOM_MACHINES.values())


def get_monthly_report_data(month: int, year: int):
    """Trích xuất và tính toán toàn bộ số liệu thống kê trong tháng được chỉ định."""
    _, last_day_num = calendar.monthrange(year, month)
    month_start = datetime.date(year, month, 1)
    month_end = datetime.date(year, month, last_day_num)

    conn = database.get_db_connection()
    all_schedules = conn.execute("SELECT * FROM schedules ORDER BY room ASC, day_of_week ASC, start_period ASC").fetchall()
    all_bookings = conn.execute("SELECT * FROM booking_requests ORDER BY id DESC").fetchall()
    all_issues = conn.execute("SELECT * FROM issue_reports ORDER BY id DESC").fetchall()
    all_cancellations = conn.execute("SELECT * FROM cancellations ORDER BY id DESC").fetchall()
    all_worklogs = conn.execute("SELECT * FROM work_logs ORDER BY work_date DESC, id DESC").fetchall()
    conn.close()

    # 1. Lọc Lịch học TKB diễn ra trong tháng
    # Một môn học diễn ra trong tháng nếu khoảng ngày học (start_dt -> end_dt) giao thoa với tháng
    # Và tính toán số buổi học thực tế rơi vào các ngày trong tháng
    active_schedules_in_month = []
    room_monthly_stats = {r: {
        'room': r, 'machines': ROOM_MACHINES.get(r, 41),
        'regular_classes': 0, 'regular_ca': 0,
        'booking_classes': 0, 'booking_ca': 0,
        'issues_count': 0, 'issues_resolved': 0,
        'cancels_count': 0
    } for r in ROOMS_LIST}

    dept_stats = {}

    for sch in all_schedules:
        s_date = parse_date_safe(sch['study_period'])
        if not s_date:
            continue
        num_weeks = int(sch['num_weeks'] or 15)
        e_date = s_date + datetime.timedelta(days=num_weeks * 7 - 1)
        is_ended_early = bool(sch['is_ended_early'])

        # Kiểm tra giao thoa với tháng
        if s_date > month_end or e_date < month_start:
            continue

        # Đếm các ngày diễn ra ca học cụ thể trong tháng
        dow_target = int(sch['day_of_week'])  # 2..8
        ca_slots = get_schedule_ca_slots(sch['start_period'], sch['end_period'])
        ca_count_per_session = len(ca_slots)

        sessions_in_month = 0
        curr_dt = month_start
        while curr_dt <= month_end:
            curr_dow = curr_dt.weekday() + 2  # Mon=2, ..., Sun=8
            if curr_dow == dow_target and s_date <= curr_dt <= e_date:
                # Nếu không bị kết thúc sớm trước ngày này
                if not is_ended_early or curr_dt <= get_vn_today():
                    sessions_in_month += 1
            curr_dt += datetime.timedelta(days=1)

        if sessions_in_month > 0:
            total_ca_sch = sessions_in_month * ca_count_per_session
            room_code = str(sch['room']).replace('Phòng ', '').replace('P.', '').strip()
            if room_code in room_monthly_stats:
                room_monthly_stats[room_code]['regular_classes'] += 1
                room_monthly_stats[room_code]['regular_ca'] += total_ca_sch

            # Phân bổ theo Khoa
            dept = sch['department'] or 'Chưa phân khoa'
            dept_stats[dept] = dept_stats.get(dept, 0) + total_ca_sch

            active_schedules_in_month.append({
                'row': sch,
                'sessions_in_month': sessions_in_month,
                'ca_per_session': ca_count_per_session,
                'total_ca_month': total_ca_sch,
                'start_dt': s_date,
                'end_dt': e_date,
                'is_ended_early': is_ended_early
            })

    # 2. Lọc Yêu cầu mượn phòng trong tháng
    bookings_in_month = []
    for b in all_bookings:
        # Xét theo ngày mượn (study_period) hoặc ngày tạo (created_at)
        b_date = parse_date_safe(b['study_period'])
        c_date = parse_date_safe(b['created_at'])
        in_month = False
        if b_date and b_date.year == year and b_date.month == month:
            in_month = True
        elif c_date and c_date.year == year and c_date.month == month:
            in_month = True

        if in_month:
            bookings_in_month.append(b)
            if b['status'] == 'Approved':
                b_room = str(b['room']).replace('Phòng ', '').replace('P.', '').strip()
                b_ca = len(get_schedule_ca_slots(b['start_period'], b['end_period']))
                if b_room in room_monthly_stats:
                    room_monthly_stats[b_room]['booking_classes'] += 1
                    room_monthly_stats[b_room]['booking_ca'] += b_ca

    # 3. Lọc Sự cố kỹ thuật trong tháng
    issues_in_month = []
    issue_categories = {}
    for iss in all_issues:
        iss_dt = parse_date_safe(iss['created_at'])
        if iss_dt and iss_dt.year == year and iss_dt.month == month:
            issues_in_month.append(iss)
            i_room = str(iss['room']).replace('Phòng ', '').replace('P.', '').strip()
            if i_room in room_monthly_stats:
                room_monthly_stats[i_room]['issues_count'] += 1
                if iss['status'] in ('Resolved', 'Approved'):
                    room_monthly_stats[i_room]['issues_resolved'] += 1
            cat = iss['issue_category'] or 'Khác'
            issue_categories[cat] = issue_categories.get(cat, 0) + 1

    # 4. Lọc Báo vắng trong tháng
    cancels_in_month = []
    for c in all_cancellations:
        c_date = parse_date_safe(c['cancel_date'])
        if c_date and c_date.year == year and c_date.month == month:
            cancels_in_month.append(c)
            c_room = str(c['room']).replace('Phòng ', '').replace('P.', '').strip()
            if c_room in room_monthly_stats:
                room_monthly_stats[c_room]['cancels_count'] += 1

    # 5. Lọc Nhật ký Admin (Worklogs) trong tháng
    worklogs_in_month = []
    for wl in all_worklogs:
        w_date = parse_date_safe(wl['work_date'])
        if w_date and w_date.year == year and w_date.month == month:
            worklogs_in_month.append(wl)

    # 6. Tính toán các chỉ số KPI trọng yếu
    total_regular_ca = sum(s['regular_ca'] for s in room_monthly_stats.values())
    total_booking_ca = sum(s['booking_ca'] for s in room_monthly_stats.values())
    total_served_ca = total_regular_ca + total_booking_ca

    # Công suất thiết kế tối đa trong tháng: 8 phòng * 4 ca/ngày * số ngày làm việc (khoảng 26 ngày trừ Chủ nhật)
    workdays_in_month = sum(1 for d in range(1, last_day_num + 1) if datetime.date(year, month, d).weekday() < 6)
    max_capacity_ca = len(ROOMS_LIST) * 4 * workdays_in_month  # 8 * 4 * 26 = 832 ca
    utilization_rate = round((total_served_ca / max_capacity_ca) * 100, 1) if max_capacity_ca > 0 else 0.0

    total_issues = len(issues_in_month)
    resolved_issues = sum(1 for i in issues_in_month if i['status'] in ('Resolved', 'Approved'))
    issue_resolve_rate = round((resolved_issues / total_issues) * 100, 1) if total_issues > 0 else 100.0

    total_bookings = len(bookings_in_month)
    approved_bookings = sum(1 for b in bookings_in_month if b['status'] == 'Approved')
    rejected_bookings = sum(1 for b in bookings_in_month if b['status'] == 'Rejected')

    # Tính tỷ lệ sử dụng riêng từng phòng
    for r in ROOMS_LIST:
        r_total = room_monthly_stats[r]['regular_ca'] + room_monthly_stats[r]['booking_ca']
        r_max = 4 * workdays_in_month
        room_monthly_stats[r]['total_ca'] = r_total
        room_monthly_stats[r]['util_rate'] = round((r_total / r_max) * 100, 1) if r_max > 0 else 0.0

    return {
        'month': month,
        'year': year,
        'month_start': month_start,
        'month_end': month_end,
        'workdays': workdays_in_month,
        'max_capacity_ca': max_capacity_ca,
        'total_served_ca': total_served_ca,
        'total_regular_ca': total_regular_ca,
        'total_booking_ca': total_booking_ca,
        'utilization_rate': utilization_rate,
        'total_issues': total_issues,
        'resolved_issues': resolved_issues,
        'issue_resolve_rate': issue_resolve_rate,
        'total_bookings': total_bookings,
        'approved_bookings': approved_bookings,
        'rejected_bookings': rejected_bookings,
        'total_cancels': len(cancels_in_month),
        'total_worklogs': len(worklogs_in_month),
        'room_stats': room_monthly_stats,
        'dept_stats': dept_stats,
        'issue_categories': issue_categories,
        'active_schedules': active_schedules_in_month,
        'bookings': bookings_in_month,
        'issues': issues_in_month,
        'cancellations': cancels_in_month,
        'worklogs': worklogs_in_month
    }


def build_monthly_excel_report(month: int, year: int) -> io.BytesIO:
    """Tạo Workbook Excel (.xlsx) Báo cáo Tổng hợp Vận hành & Khai thác 8 Phòng máy Tháng MM/YYYY."""
    data = get_monthly_report_data(month, year)
    wb = openpyxl.Workbook()

    # Font và màu sắc chuẩn HCMUS
    f_title = Font(name="Calibri", size=14, bold=True, color="0F2942")
    f_sub = Font(name="Calibri", size=10, italic=True, color="475569")
    f_sec_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    f_tbl_head = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    f_body = Font(name="Calibri", size=10, color="1F2937")
    f_bold = Font(name="Calibri", size=10, bold=True, color="1F2937")
    f_kpi_val = Font(name="Calibri", size=16, bold=True, color="0D9488")
    f_kpi_lbl = Font(name="Calibri", size=9, bold=True, color="475569")

    fill_navy = PatternFill("solid", fgColor="0F2942")
    fill_teal = PatternFill("solid", fgColor="0D9488")
    fill_kpi_bg = PatternFill("solid", fgColor="F0FDFA")
    fill_zebra = PatternFill("solid", fgColor="F8FAFC")
    fill_warn_bg = PatternFill("solid", fgColor="FEF3C7")

    al_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_right = Alignment(horizontal="right", vertical="center")

    thin_side = Side(style='thin', color='CBD5E1')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # ==========================================
    # SHEET 1: TỔNG QUAN KPI & MA TRẬN 8 PHÒNG MÁY
    # ==========================================
    ws1 = wb.active
    ws1.title = "TongQuan_KPI"
    ws1.views.sheetView[0].showGridLines = True

    # Header hành chính
    ws1.merge_cells("A1:C1")
    ws1["A1"] = "TRƯỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN"
    ws1["A1"].font = Font(name="Calibri", size=10, color="334155")
    ws1["A1"].alignment = al_center

    ws1.merge_cells("A2:C2")
    ws1["A2"] = "PHÒNG QUẢN TRỊ THIẾT BỊ (CSLT)"
    ws1["A2"].font = Font(name="Calibri", size=10, bold=True, color="0F2942")
    ws1["A2"].alignment = al_center

    ws1.merge_cells("F1:H1")
    ws1["F1"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    ws1["F1"].font = Font(name="Calibri", size=10, bold=True, color="0F2942")
    ws1["F1"].alignment = al_center

    ws1.merge_cells("F2:H2")
    ws1["F2"] = "Độc lập - Tự do - Hạnh phúc"
    ws1["F2"].font = Font(name="Calibri", size=10, bold=True, color="334155")
    ws1["F2"].alignment = al_center

    ws1.merge_cells("A4:H4")
    ws1["A4"] = f"BÁO CÁO TỔNG QUAN VẬN HÀNH & KHAI THÁC 8 PHÒNG MÁY TÍNH THÁNG {month:02d}/{year}"
    ws1["A4"].font = f_title
    ws1["A4"].alignment = al_center

    ws1.merge_cells("A5:H5")
    ws1["A5"] = f"(Số liệu ghi nhận tự động từ ngày 01/{month:02d}/{year} đến ngày {data['month_end'].day:02d}/{month:02d}/{year} tại Cơ sở 02)"
    ws1["A5"].font = f_sub
    ws1["A5"].alignment = al_center

    # 4 THẺ KPI CARDS (Dòng 7 - 9)
    kpis = [
        ("A", "B", "TỔNG CA PHỤC VỤ", f"{data['total_served_ca']} ca", f"Chính khóa: {data['total_regular_ca']} | Mượn: {data['total_booking_ca']}"),
        ("C", "D", "TẦN SUẤT LẤP ĐẦY", f"{data['utilization_rate']}%", f"Thiết kế: {data['max_capacity_ca']} ca tối đa"),
        ("E", "F", "SỰ CỐ ĐÃ KHẮC PHỤC", f"{data['resolved_issues']}/{data['total_issues']} ({data['issue_resolve_rate']}%)", f"Chờ xử lý: {data['total_issues'] - data['resolved_issues']} sự cố"),
        ("G", "H", "YÊU CẦU ĐÃ DUYỆT", f"{data['approved_bookings']}/{data['total_bookings']}", f"Báo vắng: {data['total_cancels']} ca | Log: {data['total_worklogs']}")
    ]
    for c_start, c_end, title, val, sub in kpis:
        c1 = f"{c_start}7"
        ws1.merge_cells(f"{c_start}7:{c_end}7")
        ws1[c1] = title
        ws1[c1].font = f_kpi_lbl
        ws1[c1].alignment = al_center
        ws1[c1].fill = fill_kpi_bg

        ws1.merge_cells(f"{c_start}8:{c_end}8")
        ws1[f"{c_start}8"] = val
        ws1[f"{c_start}8"].font = f_kpi_val
        ws1[f"{c_start}8"].alignment = al_center
        ws1[f"{c_start}8"].fill = fill_kpi_bg

        ws1.merge_cells(f"{c_start}9:{c_end}9")
        ws1[f"{c_start}9"] = sub
        ws1[f"{c_start}9"].font = f_sub
        ws1[f"{c_start}9"].alignment = al_center
        ws1[f"{c_start}9"].fill = fill_kpi_bg

    # MA TRẬN 8 PHÒNG MÁY (Dòng 11)
    ws1.merge_cells("A11:H11")
    ws1["A11"] = "BẢNG TỔNG HỢP KHAI THÁC & KỸ THUẬT THEO TỪNG PHÒNG MÁY"
    ws1["A11"].font = f_sec_header
    ws1["A11"].fill = fill_navy
    ws1["A11"].alignment = al_left

    headers1 = ["STT", "Phòng máy", "Quy mô", "Ca TKB chính khóa", "Ca mượn bổ sung", "Tổng ca phục vụ", "Tần suất khai thác", "Sự cố kỹ thuật (Đã xong/Tổng)"]
    for col_idx, h_text in enumerate(headers1, 1):
        cell = ws1.cell(12, col_idx, h_text)
        cell.font = f_tbl_head
        cell.fill = fill_teal
        cell.alignment = al_center
        cell.border = border_all

    row_cur = 13
    for idx, r_code in enumerate(ROOMS_LIST, 1):
        st = data['room_stats'][r_code]
        iss_str = f"{st['issues_resolved']}/{st['issues_count']}" if st['issues_count'] > 0 else "0/0 (Tốt)"

        vals = [
            idx,
            f"Phòng {r_code}",
            f"{st['machines']} máy",
            f"{st['regular_ca']} ca",
            f"{st['booking_ca']} ca",
            f"{st['total_ca']} ca",
            f"{st['util_rate']}%",
            iss_str
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws1.cell(row_cur, c_idx, val)
            cell.font = f_body
            cell.border = border_all
            cell.alignment = al_center if c_idx in (1, 2, 3, 7, 8) else al_right
            if row_cur % 2 == 0:
                cell.fill = fill_zebra
        row_cur += 1

    # Dòng Tổng cộng
    total_vals = [
        "—", "TỔNG CỘNG 8 PHÒNG", f"{TOTAL_MACHINES} máy",
        f"{data['total_regular_ca']} ca", f"{data['total_booking_ca']} ca",
        f"{data['total_served_ca']} ca", f"{data['utilization_rate']}%",
        f"{data['resolved_issues']}/{data['total_issues']}"
    ]
    for c_idx, val in enumerate(total_vals, 1):
        cell = ws1.cell(row_cur, c_idx, val)
        cell.font = f_bold
        cell.fill = fill_warn_bg
        cell.border = border_all
        cell.alignment = al_center if c_idx in (1, 2, 3, 7, 8) else al_right
    row_cur += 2

    # Bảng Phân bổ theo Khoa / Bộ môn
    ws1.merge_cells(f"A{row_cur}:D{row_cur}")
    ws1[f"A{row_cur}"] = "CƠ CẤU SỬ DỤNG PHÒNG MÁY THEO KHOA / ĐƠN VỊ"
    ws1[f"A{row_cur}"].font = f_sec_header
    ws1[f"A{row_cur}"].fill = fill_navy
    ws1[f"A{row_cur}"].alignment = al_left
    row_cur += 1

    dept_headers = ["STT", "Khoa / Đơn vị", "Tổng số ca học trong tháng", "Tỷ trọng (%)"]
    for c_idx, h in enumerate(dept_headers, 1):
        cell = ws1.cell(row_cur, c_idx, h)
        cell.font = f_tbl_head
        cell.fill = fill_teal
        cell.alignment = al_center
        cell.border = border_all
    row_cur += 1

    sorted_depts = sorted(data['dept_stats'].items(), key=lambda x: x[1], reverse=True)
    for idx, (dept_name, ca_num) in enumerate(sorted_depts, 1):
        pct = round((ca_num / data['total_regular_ca']) * 100, 1) if data['total_regular_ca'] > 0 else 0
        d_vals = [idx, dept_name, f"{ca_num} ca", f"{pct}%"]
        for c_idx, val in enumerate(d_vals, 1):
            cell = ws1.cell(row_cur, c_idx, val)
            cell.font = f_body
            cell.border = border_all
            cell.alignment = al_center if c_idx in (1, 4) else (al_left if c_idx == 2 else al_right)
            if row_cur % 2 == 0:
                cell.fill = fill_zebra
        row_cur += 1

    # ==========================================
    # SHEET 2: PHÂN BỔ THỜI KHÓA BIỂU 8 PHÒNG MÁY
    # ==========================================
    ws2 = wb.create_sheet("PhanBo_TKB")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:I1")
    ws2["A1"] = f"DANH SÁCH MÔN HỌC & LỊCH PHÂN BỔ 8 PHÒNG MÁY - THÁNG {month:02d}/{year}"
    ws2["A1"].font = f_title
    ws2["A1"].alignment = al_left

    s2_headers = ["STT", "Phòng", "Thứ", "Tiết", "Mã lớp", "Tên môn học", "Khoa / Đơn vị", "Ghi chú", "Số buổi & Ca trong tháng", "Trạng thái"]
    for c_idx, h in enumerate(s2_headers, 1):
        cell = ws2.cell(3, c_idx, h)
        cell.font = f_tbl_head
        cell.fill = fill_navy
        cell.alignment = al_center
        cell.border = border_all

    s2_row = 4
    for idx, sch_item in enumerate(data['active_schedules'], 1):
        sch = sch_item['row']
        room_num = str(sch['room']).replace('Phòng ', '').replace('P.', '').strip()
        status_txt = "Đã kết thúc sớm" if sch_item['is_ended_early'] else "Đang học bình thường"
        keys = sch.keys() if hasattr(sch, 'keys') else []

        c_vals = [
            idx,
            f"P.{room_num}",
            f"Thứ {sch['day_of_week']}",
            f"Tiết {sch['start_period']}-{sch['end_period']}",
            sch['class_code'] or '',
            sch['subject_name'] or '',
            sch['department'] or '',
            (sch['notes'] if 'notes' in keys else '') or '',
            f"{sch_item['sessions_in_month']} buổi ({sch_item['total_ca_month']} ca)",
            status_txt
        ]
        for c_idx, val in enumerate(c_vals, 1):
            cell = ws2.cell(s2_row, c_idx, val)
            cell.font = f_body
            cell.border = border_all
            cell.alignment = al_center if c_idx in (1, 2, 3, 4, 9, 10) else al_left
            if s2_row % 2 == 0:
                cell.fill = fill_zebra
        s2_row += 1

    # ==========================================
    # SHEET 3: VẬN HÀNH KỸ THUẬT & YÊU CẦU USER
    # ==========================================
    ws3 = wb.create_sheet("VanHanh_KyThuat")
    ws3.views.sheetView[0].showGridLines = True

    # 3.1. Bảng Mượn phòng
    ws3.merge_cells("A1:H1")
    ws3["A1"] = f"1. NHẬT KÝ TIẾP NHẬN & XỬ LÝ MƯỢN PHÒNG / DẠY BÙ - THÁNG {month:02d}/{year}"
    ws3["A1"].font = f_sec_header
    ws3["A1"].fill = fill_navy
    ws3["A1"].alignment = al_left

    s3_b_headers = ["STT", "Thời gian gửi", "Người mượn", "Phòng", "Thời gian mượn", "Mục đích & Phần mềm", "Trạng thái", "Phản hồi QTTB"]
    for c_idx, h in enumerate(s3_b_headers, 1):
        cell = ws3.cell(2, c_idx, h)
        cell.font = f_tbl_head
        cell.fill = fill_teal
        cell.alignment = al_center
        cell.border = border_all

    s3_row = 3
    for idx, b in enumerate(data['bookings'], 1):
        st_label = "Đã duyệt" if b['status'] == 'Approved' else ("Từ chối" if b['status'] == 'Rejected' else "Đã tiếp nhận")
        b_vals = [
            idx,
            b['created_at'] or '',
            f"{b['requester_name']} ({b['department']})",
            f"P.{b['room']}",
            f"Thứ {b['day_of_week']} (Tiết {b['start_period']}-{b['end_period']})",
            f"{b['purpose']}" + (f" - PM: {b['software_required']}" if b['software_required'] else ''),
            st_label,
            b['admin_note'] or ''
        ]
        for c_idx, val in enumerate(b_vals, 1):
            cell = ws3.cell(s3_row, c_idx, val)
            cell.font = f_body
            cell.border = border_all
            cell.alignment = al_center if c_idx in (1, 2, 4, 7) else al_left
            if s3_row % 2 == 0:
                cell.fill = fill_zebra
        s3_row += 1
    if not data['bookings']:
        ws3.cell(s3_row, 1, "Không có yêu cầu mượn phòng nào trong tháng.").font = f_sub
        s3_row += 1

    s3_row += 2

    # 3.2. Bảng Sự cố
    ws3.merge_cells(f"A{s3_row}:H{s3_row}")
    ws3[f"A{s3_row}"] = f"2. NHẬT KÝ BÁO HỎNG & KHẮC PHỤC SỰ CỐ KỸ THUẬT - THÁNG {month:02d}/{year}"
    ws3[f"A{s3_row}"].font = f_sec_header
    ws3[f"A{s3_row}"].fill = fill_navy
    ws3[f"A{s3_row}"].alignment = al_left
    s3_row += 1

    s3_i_headers = ["STT", "Thời gian báo", "Người báo", "Vị trí", "Phân loại", "Mô tả sự cố", "Trạng thái", "Ghi chú sửa chữa"]
    for c_idx, h in enumerate(s3_i_headers, 1):
        cell = ws3.cell(s3_row, c_idx, h)
        cell.font = f_tbl_head
        cell.fill = fill_teal
        cell.alignment = al_center
        cell.border = border_all
    s3_row += 1

    for idx, iss in enumerate(data['issues'], 1):
        st_label = "Đã xong" if iss['status'] in ('Resolved', 'Approved') else "Đang xử lý"
        i_vals = [
            idx,
            iss['created_at'] or '',
            f"{iss['reporter_name']} ({iss['department']})",
            f"P.{iss['room']} - {iss['machine_number']}",
            iss['issue_category'],
            iss['description'],
            st_label,
            iss['admin_note'] or ''
        ]
        for c_idx, val in enumerate(i_vals, 1):
            cell = ws3.cell(s3_row, c_idx, val)
            cell.font = f_body
            cell.border = border_all
            cell.alignment = al_center if c_idx in (1, 2, 4, 7) else al_left
            if s3_row % 2 == 0:
                cell.fill = fill_zebra
        s3_row += 1
    if not data['issues']:
        ws3.cell(s3_row, 1, "Không có sự cố kỹ thuật nào phát sinh trong tháng.").font = f_sub
        s3_row += 1

    s3_row += 2

    # 3.3. Bảng Báo vắng
    ws3.merge_cells(f"A{s3_row}:H{s3_row}")
    ws3[f"A{s3_row}"] = f"3. DANH SÁCH GIẢNG VIÊN BÁO VẮNG LỚP HỌC - THÁNG {month:02d}/{year}"
    ws3[f"A{s3_row}"].font = f_sec_header
    ws3[f"A{s3_row}"].fill = fill_navy
    ws3[f"A{s3_row}"].alignment = al_left
    s3_row += 1

    s3_c_headers = ["STT", "Thời gian gửi", "Giảng viên", "Phòng", "Ngày vắng", "Ca / Tiết", "Lý do", "Trạng thái"]
    for c_idx, h in enumerate(s3_c_headers, 1):
        cell = ws3.cell(s3_row, c_idx, h)
        cell.font = f_tbl_head
        cell.fill = fill_teal
        cell.alignment = al_center
        cell.border = border_all
    s3_row += 1

    for idx, c in enumerate(data['cancellations'], 1):
        st_label = "Đã duyệt" if c['status'] == 'Approved' else "Đã tiếp nhận"
        c_vals = [
            idx,
            c['created_at'] or '',
            c['teacher_name'],
            f"P.{c['room']}",
            c['cancel_date'],
            c['period'],
            c['reason'] or '',
            st_label
        ]
        for c_idx, val in enumerate(c_vals, 1):
            cell = ws3.cell(s3_row, c_idx, val)
            cell.font = f_body
            cell.border = border_all
            cell.alignment = al_center if c_idx in (1, 2, 4, 5, 8) else al_left
            if s3_row % 2 == 0:
                cell.fill = fill_zebra
        s3_row += 1
    if not data['cancellations']:
        ws3.cell(s3_row, 1, "Không có lượt báo vắng nào trong tháng.").font = f_sub

    # ==========================================
    # SHEET 4: NHẬT KÝ CA TRỰC ADMIN
    # ==========================================
    ws4 = wb.create_sheet("NhatKy_Admin")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:H1")
    ws4["A1"] = f"SỔ THEO DÕI CA TRỰC & BẢO TRÌ HỆ THỐNG - THÁNG {month:02d}/{year}"
    ws4["A1"].font = f_title
    ws4["A1"].alignment = al_left

    s4_headers = ["STT", "Ngày trực", "Ca trực", "Phạm vi phòng", "Phân loại", "Nội dung thực hiện", "Trạng thái", "Nhân sự & Bàn giao"]
    for c_idx, h in enumerate(s4_headers, 1):
        cell = ws4.cell(3, c_idx, h)
        cell.font = f_tbl_head
        cell.fill = fill_navy
        cell.alignment = al_center
        cell.border = border_all

    s4_row = 4
    for idx, wl in enumerate(data['worklogs'], 1):
        st_label = "Hoàn thành" if wl['status'] == 'completed' else "Đang làm"
        w_vals = [
            idx,
            wl['work_date'],
            wl['shift_name'],
            wl['target_rooms'] or '',
            wl['category'] or '',
            wl['content'] or '',
            st_label,
            f"{wl['created_by']} - {wl['notes'] or ''}".strip(' -')
        ]
        for c_idx, val in enumerate(w_vals, 1):
            cell = ws4.cell(s4_row, c_idx, val)
            cell.font = f_body
            cell.border = border_all
            cell.alignment = al_center if c_idx in (1, 2, 3, 7) else al_left
            if s4_row % 2 == 0:
                cell.fill = fill_zebra
        s4_row += 1
    if not data['worklogs']:
        ws4.cell(s4_row, 1, "Chưa có nhật ký ca trực nào được ghi nhận trong tháng.").font = f_sub

    # Tự động co giãn độ rộng cột cho cả 4 Sheet
    for ws in [ws1, ws2, ws3, ws4]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.coordinate in ws.merged_cells:
                    continue
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def _build_monthly_word_report_openxml(month: int, year: int, data: dict) -> io.BytesIO:
    """Tạo file Word (.docx) thuần Python (zipfile + OpenXML) khi server chưa cài thư viện python-docx."""
    esc = html.escape
    today = get_vn_today()

    def _w_rpr(sz_pt=12, bold=False, italic=False, color="000000"):
        sz_val = int(round(sz_pt * 2))
        b_tag = "<w:b/><w:bCs/>" if bold else ""
        i_tag = "<w:i/><w:iCs/>" if italic else ""
        return (
            f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            f'{b_tag}{i_tag}<w:color w:val="{color}"/><w:sz w:val="{sz_val}"/><w:szCs w:val="{sz_val}"/></w:rPr>'
        )

    def _w_p(text="", align="left", sz_pt=12.5, bold=False, italic=False, bold_prefix=None, before=40, after=60, line=300, color="000000"):
        runs = []
        if bold_prefix:
            runs.append(f'<w:r>{_w_rpr(sz_pt=sz_pt, bold=True, color=color)}<w:t xml:space="preserve">{esc(str(bold_prefix))}</w:t></w:r>')
        lines = str(text).split('\n')
        for idx_l, ln in enumerate(lines):
            br = "<w:br/>" if idx_l > 0 else ""
            runs.append(f'<w:r>{_w_rpr(sz_pt=sz_pt, bold=bold, italic=italic, color=color)}{br}<w:t xml:space="preserve">{esc(ln)}</w:t></w:r>')
        return (
            f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'
            f'<w:jc w:val="{align}"/></w:pPr>{"".join(runs)}</w:p>'
        )

    # 1. Bảng Header Quốc hiệu - Tiêu ngữ (Không viền)
    no_borders = '<w:tcBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'
    hdr_tbl = (
        '<w:tbl><w:tblPr><w:tblW w:w="9500" w:type="dxa"/><w:jc w:val="center"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4200"/><w:gridCol w:w="5300"/></w:tblGrid>'
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="4200" w:type="dxa"/>{no_borders}</w:tcPr>{_w_p("TRƯỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN", align="center", sz_pt=11, before=0, after=40, line=260)}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="5300" w:type="dxa"/>{no_borders}</w:tcPr>{_w_p("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", align="center", sz_pt=11, bold=True, before=0, after=40, line=260)}</w:tc>'
        '</w:tr>'
        '<w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="4200" w:type="dxa"/>{no_borders}</w:tcPr>'
        f'{_w_p("PHÒNG QUẢN TRỊ THIẾT BỊ (CSLT)", align="center", sz_pt=11, bold=True, before=0, after=20, line=260)}'
        f'{_w_p("─────────────", align="center", sz_pt=8, before=0, after=0, line=240, color="323232")}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="5300" w:type="dxa"/>{no_borders}</w:tcPr>'
        f'{_w_p("Độc lập - Tự do - Hạnh phúc", align="center", sz_pt=12, bold=True, before=0, after=20, line=260)}'
        f'{_w_p("──────────────────", align="center", sz_pt=8, before=0, after=0, line=240, color="323232")}</w:tc>'
        '</w:tr></w:tbl>'
    )

    # 2. Bảng thống kê 8 phòng máy
    t_headers = ["STT", "Phòng", "Quy mô", "Ca chính khóa", "Ca mượn bù", "Tần suất (%)"]
    stats_rows_xml = []
    hdr_cells = "".join(
        f'<w:tc><w:tcPr><w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar></w:tcPr>'
        f'{_w_p(h, align="center", sz_pt=11, bold=True, before=40, after=40, line=240)}</w:tc>'
        for h in t_headers
    )
    stats_rows_xml.append(f'<w:tr>{hdr_cells}</w:tr>')

    for idx_r, r_code in enumerate(ROOMS_LIST, 1):
        st = data['room_stats'][r_code]
        vals = [str(idx_r), f"P.{r_code}", f"{st['machines']} máy", f"{st['regular_ca']} ca", f"{st['booking_ca']} ca", f"{st['util_rate']}%"]
        row_cells = "".join(
            f'<w:tc><w:tcPr><w:tcMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar></w:tcPr>'
            f'{_w_p(v, align=("left" if c_i == 1 else "center"), sz_pt=11, before=20, after=20, line=240)}</w:tc>'
            for c_i, v in enumerate(vals)
        )
        stats_rows_xml.append(f'<w:tr>{row_cells}</w:tr>')

    tot_vals = ["—", "TỔNG", f"{TOTAL_MACHINES} máy", f"{data['total_regular_ca']} ca", f"{data['total_booking_ca']} ca", f"{data['utilization_rate']}%"]
    tot_cells = "".join(
        f'<w:tc><w:tcPr><w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar></w:tcPr>'
        f'{_w_p(v, align=("left" if c_i == 1 else "center"), sz_pt=11, bold=True, before=40, after=40, line=240)}</w:tc>'
        for c_i, v in enumerate(tot_vals)
    )
    stats_rows_xml.append(f'<w:tr>{tot_cells}</w:tr>')

    stats_tbl = (
        '<w:tbl><w:tblPr><w:tblW w:w="9300" w:type="dxa"/><w:jc w:val="center"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        '<w:left w:val="none"/><w:right w:val="none"/></w:tblBorders></w:tblPr>'
        f'{"".join(stats_rows_xml)}</w:tbl>'
    )

    # 3. Bảng chữ ký & nơi nhận
    noi_nhan_p = _w_p("- Như kính gửi;\n- Lãnh đạo Phòng QTTB (để b/c);\n- Lưu: VT, Tổ Phòng máy B4.2.", align="left", sz_pt=10.5, before=0, after=20, line=260)
    ky_ten_p = _w_p("(Ký và ghi rõ họ tên)\n\n\n\n", align="center", sz_pt=11, italic=True, before=0, after=40, line=260)
    sig_tbl = (
        '<w:tbl><w:tblPr><w:tblW w:w="9500" w:type="dxa"/><w:jc w:val="center"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4600"/><w:gridCol w:w="4900"/></w:tblGrid><w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="4600" w:type="dxa"/>{no_borders}</w:tcPr>'
        f'{_w_p("Nơi nhận:", align="left", sz_pt=11, bold=True, italic=True, before=0, after=40, line=260)}'
        f'{noi_nhan_p}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="4900" w:type="dxa"/>{no_borders}</w:tcPr>'
        f'{_w_p("NGƯỜI LẬP BÁO CÁO", align="center", sz_pt=12, bold=True, before=0, after=40, line=260)}'
        f'{ky_ten_p}</w:tc>'
        '</w:tr></w:tbl>'
    )

    booking_pct = round(data['approved_bookings'] / data['total_bookings'] * 100, 1) if data['total_bookings'] else 100
    body_parts = [
        hdr_tbl,
        _w_p(f"TP. Hồ Chí Minh, ngày {today.day:02d} tháng {today.month:02d} năm {today.year}", align="right", sz_pt=12, italic=True, before=240, after=360),
        _w_p("BÁO CÁO", align="center", sz_pt=16, bold=True, before=0, after=80),
        _w_p(f"Tần suất khai thác và vận hành 8 phòng máy tính tháng {month:02d} năm {year}", align="center", sz_pt=13.5, bold=True, before=0, after=80),
        _w_p("(tại cơ sở 02)", align="center", sz_pt=12, bold=True, before=0, after=280),
        _w_p("Kính gửi:  Lãnh đạo Phòng Quản trị Thiết bị\n                 Ban Giám hiệu Trường ĐH Khoa học Tự nhiên", align="center", sz_pt=12, italic=True, before=0, after=280),
        _w_p("I. TÌNH HÌNH PHÂN BỔ THỜI KHÓA BIỂU & KHAI THÁC 8 PHÒNG MÁY", align="left", sz_pt=13, bold=True, before=200, after=80),
        _w_p(
            f"Trong tháng {month:02d}/{year} (từ ngày 01/{month:02d}/{year} đến ngày {data['month_end'].day:02d}/{month:02d}/{year}), "
            f"Tổ Phòng máy tính Cơ sở 2 đã vận hành liên tục 8 phòng máy (P.5.1 đến P.6.6) với tổng quy mô {TOTAL_MACHINES} máy trạm. Kết quả khai thác ghi nhận:",
            align="both", bold_prefix="1. Tổng quan số liệu: "
        ),
        _w_p(f"• Tổng số ca máy phục vụ thực tế: {data['total_served_ca']} ca (trong đó: {data['total_regular_ca']} ca chính khóa và {data['total_booking_ca']} ca mượn phòng ngoài giờ/thi/dạy bù).", align="both"),
        _w_p(f"• Tần suất lấp đầy bình quân toàn hệ thống: {data['utilization_rate']}% trên tổng công suất thiết kế ({data['max_capacity_ca']} ca/tháng).", align="both"),
        _w_p(f"• Số môn học có lịch diễn ra trong tháng: {len(data['active_schedules'])} môn thuộc các Khoa: CNTT, Toán - Tin, Hóa học, Sinh học, Vật lý, Điện tử...", align="both"),
        _w_p("2. Bảng tổng hợp công suất khai thác theo từng phòng máy:", align="both"),
        stats_tbl,
        _w_p("II. CÔNG TÁC KỸ THUẬT & PHỤC VỤ NGƯỜI DÙNG (USER REQUESTS)", align="left", sz_pt=13, bold=True, before=200, after=80),
        _w_p(
            f"Hệ thống trực tuyến đã tiếp nhận tổng cộng {data['total_bookings']} yêu cầu mượn phòng học, thi và thực hành ngoài giờ. "
            f"Trong đó: Đã duyệt {data['approved_bookings']} yêu cầu (đáp ứng {booking_pct}% nhu cầu giảng viên); "
            f"Từ chối {data['rejected_bookings']} yêu cầu do trùng phòng đã có lịch chính khóa hoặc phòng không phù hợp cấu hình phần mềm.",
            align="both", bold_prefix="1. Tiếp nhận & điều phối mượn phòng: "
        ),
        _w_p(
            f"Ghi nhận tổng cộng {data['total_issues']} sự cố kỹ thuật từ Giảng viên và Sinh viên. "
            f"Tổ kỹ thuật đã xử lý dứt điểm {data['resolved_issues']}/{data['total_issues']} sự cố (đạt tỷ lệ khắc phục {data['issue_resolve_rate']}%). "
            f"Các dạng sự cố chủ yếu gồm: Lỗi chuột/phím ({data['issue_categories'].get('Chuột / Bàn phím', 0)} trường hợp), "
            f"Lỗi kết nối mạng/máy trạm ({data['issue_categories'].get('Mạng LAN / Internet', 0)} trường hợp) và hỗ trợ phần mềm chuyên ngành.",
            align="both", bold_prefix="2. Khắc phục sự cố kỹ thuật máy trạm: "
        ),
        _w_p(
            f"Đã tiếp nhận {data['total_cancels']} thông báo nghỉ lớp đột xuất từ Giảng viên. Bộ phận QTTB đã kịp thời cập nhật phòng trống trên bản đồ Thời khóa biểu trực tuyến "
            f"để nhân viên quản lý nắm bắt và tạo điều kiện bố trí phòng cho các lớp khác mượn dạy bù nhanh chóng.",
            align="both", bold_prefix="3. Xử lý thông báo nghỉ & Báo vắng: "
        ),
        _w_p("III. NHẬT KÝ CA TRỰC & CÔNG VIỆC QUẢN TRỊ HỆ THỐNG", align="left", sz_pt=13, bold=True, before=200, after=80),
        _w_p(
            f"Trong tháng ghi nhận {data['total_worklogs']} lượt nhật ký ca trực và biên bản kỹ thuật từ nhân sự phòng máy. Các công tác trọng tâm đã thực hiện:",
            align="both", bold_prefix="1. Theo dõi ca trực: "
        ),
        _w_p("• Kiểm tra và duy trì kết nối mạng LAN thông suốt cho 8 IP Router phòng máy (cụm 5.1-5.3 và 6.1-6.6) qua công cụ giám sát B42Lab NetAgent.", align="both"),
        _w_p("• Triển khai cài đặt và cấu hình phần mềm thi, kiểm tra định kỳ theo đúng danh mục phần mềm đã được các Khoa/Bộ môn đăng ký.", align="both"),
        _w_p("• Thực hiện vệ sinh công nghiệp máy trạm, bàn ghế phòng học và kiểm tra hệ thống điện lạnh đảm bảo điều kiện học tập tốt nhất cho sinh viên.", align="both"),
        _w_p("IV. ĐÁNH GIÁ CHUNG, TỒN ĐỌNG & ĐỀ XUẤT KIẾN NGHỊ", align="left", sz_pt=13, bold=True, before=200, after=80),
        _w_p(
            "Toàn bộ 8 phòng máy được vận hành an toàn, đúng tiến độ đào tạo của Nhà trường. "
            "Việc số hóa công tác đăng ký mượn phòng, báo vắng và tiếp nhận sự cố kỹ thuật trực tuyến đã giúp giảm thiểu thời gian xử lý và phục vụ người dùng minh bạch, chính xác.",
            align="both", bold_prefix="1. Đánh giá chung: "
        ),
        _w_p(
            "Một số thiết bị ngoại vi (chuột, bàn phím) tại các phòng máy cụm P.5.x đã hao mòn qua nhiều học kỳ cần được thay thế định kỳ để đảm bảo hiệu suất thực hành môn học.",
            align="both", bold_prefix="2. Khó khăn, tồn đọng: "
        ),
        _w_p(
            "Kính đề xuất Lãnh đạo Phòng Quản trị Thiết bị xem xét phê duyệt mua sắm bổ sung thiết bị dự phòng (chuột quang, bàn phím, dây mạng đúc sẵn) "
            "và bảo dưỡng hệ thống điều hòa nhiệt độ tại 8 phòng máy trước đợt thi học kỳ sắp tới.",
            align="both", bold_prefix="3. Đề xuất kiến nghị: "
        ),
        _w_p("", before=0, after=240),
        sig_tbl,
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    ]

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{"".join(body_parts)}</w:body></w:document>'
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )

    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )

    doc_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    )

    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types_xml)
        zf.writestr('_rels/.rels', rels_xml)
        zf.writestr('word/_rels/document.xml.rels', doc_rels_xml)
        zf.writestr('word/document.xml', document_xml)
    output.seek(0)
    return output


def set_cell_margins_and_border(cell, top=80, bottom=80, left=100, right=100):
    """Thiết lập padding và xóa viền ô trong Word table."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in (('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)):
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def build_monthly_word_report(month: int, year: int) -> io.BytesIO:
    """Tạo Document Word (.docx) Báo cáo Tóm tắt chuẩn thể thức văn bản hành chính (Theo đúng ảnh mẫu)."""
    data = get_monthly_report_data(month, year)
    if not HAS_PYTHON_DOCX:
        return _build_monthly_word_report_openxml(month, year, data)
    doc = Document()

    # Thiết lập lề trang chuẩn Nghị định 30/2020/NĐ-CP (A4: Top 2cm, Bottom 2cm, Left 3cm, Right 1.5cm)
    for section in doc.sections:
        section.top_margin = Inches(0.79)     # 2.0 cm
        section.bottom_margin = Inches(0.79)  # 2.0 cm
        section.left_margin = Inches(1.18)    # 3.0 cm
        section.right_margin = Inches(0.60)   # 1.5 cm

    # ==============================================================
    # HEADER THEO ĐÚNG ẢNH MẪU CỦA USER:
    # TRƯỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN   |  CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM
    #   PHÒNG QUẢN TRỊ THIẾT BỊ (CSLT)    |       Độc lập - Tự do - Hạnh phúc
    #   ───────────────────────────────   |       ───────────────────────────
    # ==============================================================
    table = doc.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Độ rộng 2 cột (Cột 1: 45%, Cột 2: 55%)
    table.columns[0].width = Inches(2.9)
    table.columns[1].width = Inches(3.7)

    # Ô TRÁI TRÊN (Hàng 1, Cột 1): Tên Trường
    c00 = table.cell(0, 0)
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p00.paragraph_format.space_before = Pt(0)
    p00.paragraph_format.space_after = Pt(2)
    p00.paragraph_format.line_spacing = 1.15
    r00 = p00.add_run("TRƯỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN")
    r00.font.name = "Times New Roman"
    r00.font.size = Pt(11)
    r00.font.color.rgb = RGBColor(0, 0, 0)

    # Ô TRÁI DƯỚI (Hàng 2, Cột 1): Tên Đơn Vị + Gạch dưới
    c10 = table.cell(1, 0)
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p10.paragraph_format.space_before = Pt(0)
    p10.paragraph_format.space_after = Pt(1)
    p10.paragraph_format.line_spacing = 1.15
    r10 = p10.add_run("PHÒNG QUẢN TRỊ THIẾT BỊ (CSLT)")
    r10.font.name = "Times New Roman"
    r10.font.size = Pt(11)
    r10.font.bold = True
    r10.font.color.rgb = RGBColor(0, 0, 0)

    # Gạch dưới ngắn ô trái
    p10_line = c10.add_paragraph()
    p10_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p10_line.paragraph_format.space_before = Pt(0)
    p10_line.paragraph_format.space_after = Pt(0)
    r10_line = p10_line.add_run("─────────────")
    r10_line.font.name = "Times New Roman"
    r10_line.font.size = Pt(8)
    r10_line.font.color.rgb = RGBColor(50, 50, 50)

    # Ô PHẢI TRÊN (Hàng 1, Cột 2): Quốc Hiệu
    c01 = table.cell(0, 1)
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p01.paragraph_format.space_before = Pt(0)
    p01.paragraph_format.space_after = Pt(2)
    p01.paragraph_format.line_spacing = 1.15
    r01 = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    r01.font.name = "Times New Roman"
    r01.font.size = Pt(11)
    r01.font.bold = True
    r01.font.color.rgb = RGBColor(0, 0, 0)

    # Ô PHẢI DƯỚI (Hàng 2, Cột 2): Tiêu Ngữ + Gạch dưới
    c11 = table.cell(1, 1)
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p11.paragraph_format.space_before = Pt(0)
    p11.paragraph_format.space_after = Pt(1)
    p11.paragraph_format.line_spacing = 1.15
    r11 = p11.add_run("Độc lập - Tự do - Hạnh phúc")
    r11.font.name = "Times New Roman"
    r11.font.size = Pt(12)
    r11.font.bold = True
    r11.font.color.rgb = RGBColor(0, 0, 0)

    # Gạch dưới tiêu ngữ ô phải
    p11_line = c11.add_paragraph()
    p11_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p11_line.paragraph_format.space_before = Pt(0)
    p11_line.paragraph_format.space_after = Pt(0)
    r11_line = p11_line.add_run("──────────────────")
    r11_line.font.name = "Times New Roman"
    r11_line.font.size = Pt(8)
    r11_line.font.color.rgb = RGBColor(50, 50, 50)

    # Xóa viền bảng Header
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'''
                <w:tcBorders %s>
                    <w:top w:val="none"/>
                    <w:left w:val="none"/>
                    <w:bottom w:val="none"/>
                    <w:right w:val="none"/>
                </w:tcBorders>
            ''' % nsdecls('w'))
            tcPr.append(tcBorders)

    # ==============================================================
    # NGÀY THÁNG NĂM (Căn phải, in nghiêng)
    # TP. Hồ Chí Minh, ngày ... tháng ... năm ...
    # ==============================================================
    today = get_vn_today()
    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.paragraph_format.space_before = Pt(12)
    p_date.paragraph_format.space_after = Pt(18)
    r_date = p_date.add_run(f"TP. Hồ Chí Minh, ngày {today.day:02d} tháng {today.month:02d} năm {today.year}")
    r_date.font.name = "Times New Roman"
    r_date.font.size = Pt(12)
    r_date.font.italic = True

    # ==============================================================
    # TIÊU ĐỀ BÁO CÁO (Căn giữa, In hoa, Đậm, Size 15-16)
    # BÁO CÁO
    # Tần suất khai thác và vận hành 8 phòng máy tính tháng MM năm YYYY
    # (tại cơ sở 02)
    # ==============================================================
    p_title1 = doc.add_paragraph()
    p_title1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title1.paragraph_format.space_before = Pt(0)
    p_title1.paragraph_format.space_after = Pt(4)
    r_t1 = p_title1.add_run("BÁO CÁO")
    r_t1.font.name = "Times New Roman"
    r_t1.font.size = Pt(16)
    r_t1.font.bold = True

    p_title2 = doc.add_paragraph()
    p_title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title2.paragraph_format.space_before = Pt(0)
    p_title2.paragraph_format.space_after = Pt(4)
    r_t2 = p_title2.add_run(f"Tần suất khai thác và vận hành 8 phòng máy tính tháng {month:02d} năm {year}")
    r_t2.font.name = "Times New Roman"
    r_t2.font.size = Pt(13.5)
    r_t2.font.bold = True

    p_title3 = doc.add_paragraph()
    p_title3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title3.paragraph_format.space_before = Pt(0)
    p_title3.paragraph_format.space_after = Pt(14)
    r_t3 = p_title3.add_run("(tại cơ sở 02)")
    r_t3.font.name = "Times New Roman"
    r_t3.font.size = Pt(12)
    r_t3.font.bold = True

    # Kính gửi
    p_kg = doc.add_paragraph()
    p_kg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kg.paragraph_format.space_before = Pt(0)
    p_kg.paragraph_format.space_after = Pt(14)
    r_kg = p_kg.add_run("Kính gửi:  Lãnh đạo Phòng Quản trị Thiết bị\n                 Ban Giám hiệu Trường ĐH Khoa học Tự nhiên")
    r_kg.font.name = "Times New Roman"
    r_kg.font.size = Pt(12)
    r_kg.font.italic = True

    # Hàm tiện ích thêm đề mục
    def add_heading_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.25
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(13)
        r.font.bold = True
        return p

    def add_body_p(text, bold_prefix=None, italic=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = "Times New Roman"
            rb.font.size = Pt(12.5)
            rb.font.bold = True
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12.5)
        r.font.italic = italic
        return p

    # ==============================================================
    # NỘI DUNG 4 PHẦN CHÍNH
    # ==============================================================
    add_heading_1("I. TÌNH HÌNH PHÂN BỔ THỜI KHÓA BIỂU & KHAI THÁC 8 PHÒNG MÁY")

    add_body_p(
        f"Trong tháng {month:02d}/{year} (từ ngày 01/{month:02d}/{year} đến ngày {data['month_end'].day:02d}/{month:02d}/{year}), "
        f"Tổ Phòng máy tính Cơ sở 2 đã vận hành liên tục 8 phòng máy (P.5.1 đến P.6.6) với tổng quy mô {TOTAL_MACHINES} máy trạm. "
        f"Kết quả khai thác ghi nhận:",
        bold_prefix="1. Tổng quan số liệu: "
    )

    add_body_p(f"• Tổng số ca máy phục vụ thực tế: {data['total_served_ca']} ca (trong đó: {data['total_regular_ca']} ca chính khóa và {data['total_booking_ca']} ca mượn phòng ngoài giờ/thi/dạy bù).")
    add_body_p(f"• Tần suất lấp đầy bình quân toàn hệ thống: {data['utilization_rate']}% trên tổng công suất thiết kế ({data['max_capacity_ca']} ca/tháng).")
    add_body_p(f"• Số môn học có lịch diễn ra trong tháng: {len(data['active_schedules'])} môn thuộc các Khoa: CNTT, Toán - Tin, Hóa học, Sinh học, Vật lý, Điện tử...")

    add_body_p("2. Bảng tổng hợp công suất khai thác theo từng phòng máy:")

    # Bảng số liệu Word
    tbl_stats = doc.add_table(rows=len(ROOMS_LIST) + 2, cols=6)
    tbl_stats.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_stats.autofit = False

    t_headers = ["STT", "Phòng", "Quy mô", "Ca chính khóa", "Ca mượn bù", "Tần suất (%)"]
    for c_i, h in enumerate(t_headers):
        cell = tbl_stats.cell(0, c_i)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(h)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.font.bold = True
        set_cell_margins_and_border(cell)

    for i, r_code in enumerate(ROOMS_LIST, 1):
        st = data['room_stats'][r_code]
        vals = [str(i), f"P.{r_code}", f"{st['machines']} máy", f"{st['regular_ca']} ca", f"{st['booking_ca']} ca", f"{st['util_rate']}%"]
        for c_i, v in enumerate(vals):
            cell = tbl_stats.cell(i, c_i)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_i != 1 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(v)
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)
            set_cell_margins_and_border(cell)

    # Dòng Tổng
    tot_row_idx = len(ROOMS_LIST) + 1
    t_vals = ["—", "TỔNG", f"{TOTAL_MACHINES} máy", f"{data['total_regular_ca']} ca", f"{data['total_booking_ca']} ca", f"{data['utilization_rate']}%"]
    for c_i, v in enumerate(t_vals):
        cell = tbl_stats.cell(tot_row_idx, c_i)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_i != 1 else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(v)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.font.bold = True
        set_cell_margins_and_border(cell)

    # Đóng khung bảng Word
    tbl_borders = parse_xml(r'''
        <w:tblBorders %s>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
        </w:tblBorders>
    ''' % nsdecls('w'))
    tbl_stats._tbl.tblPr.append(tbl_borders)

    # ==============================================================
    # PHẦN II: KỸ THUẬT & YÊU CẦU USER
    # ==============================================================
    add_heading_1("II. CÔNG TÁC KỸ THUẬT & PHỤC VỤ NGƯỜI DÙNG (USER REQUESTS)")

    add_body_p(
        f"Hệ thống trực tuyến đã tiếp nhận tổng cộng {data['total_bookings']} yêu cầu mượn phòng học, thi và thực hành ngoài giờ. "
        f"Trong đó: Đã duyệt {data['approved_bookings']} yêu cầu (đáp ứng {round(data['approved_bookings']/data['total_bookings']*100, 1) if data['total_bookings'] else 100}% nhu cầu giảng viên); "
        f"Từ chối {data['rejected_bookings']} yêu cầu do trùng phòng đã có lịch chính khóa hoặc phòng không phù hợp cấu hình phần mềm.",
        bold_prefix="1. Tiếp nhận & điều phối mượn phòng: "
    )

    add_body_p(
        f"Ghi nhận tổng cộng {data['total_issues']} sự cố kỹ thuật từ Giảng viên và Sinh viên. "
        f"Tổ kỹ thuật đã xử lý dứt điểm {data['resolved_issues']}/{data['total_issues']} sự cố (đạt tỷ lệ khắc phục {data['issue_resolve_rate']}%). "
        f"Các dạng sự cố chủ yếu gồm: Lỗi chuột/phím ({data['issue_categories'].get('Chuột / Bàn phím', 0)} trường hợp), "
        f"Lỗi kết nối mạng/máy trạm ({data['issue_categories'].get('Mạng LAN / Internet', 0)} trường hợp) và hỗ trợ phần mềm chuyên ngành.",
        bold_prefix="2. Khắc phục sự cố kỹ thuật máy trạm: "
    )

    add_body_p(
        f"Đã tiếp nhận {data['total_cancels']} thông báo nghỉ lớp đột xuất từ Giảng viên. Bộ phận QTTB đã kịp thời cập nhật phòng trống trên bản đồ Thời khóa biểu trực tuyến "
        f"để nhân viên quản lý nắm bắt và tạo điều kiện bố trí phòng cho các lớp khác mượn dạy bù nhanh chóng.",
        bold_prefix="3. Xử lý thông báo nghỉ & Báo vắng: "
    )

    # ==============================================================
    # PHẦN III: NHẬT KÝ CA TRỰC & BẢO TRÌ
    # ==============================================================
    add_heading_1("III. NHẬT KÝ CA TRỰC & CÔNG VIỆC QUẢN TRỊ HỆ THỐNG")

    add_body_p(
        f"Trong tháng ghi nhận {data['total_worklogs']} lượt nhật ký ca trực và biên bản kỹ thuật từ nhân sự phòng máy. Các công tác trọng tâm đã thực hiện:",
        bold_prefix="1. Theo dõi ca trực: "
    )
    add_body_p("• Kiểm tra và duy trì kết nối mạng LAN thông suốt cho 8 IP Router phòng máy (cụm 5.1-5.3 và 6.1-6.6) qua công cụ giám sát B42Lab NetAgent.")
    add_body_p("• Triển khai cài đặt và cấu hình phần mềm thi, kiểm tra định kỳ theo đúng danh mục phần mềm đã được các Khoa/Bộ môn đăng ký.")
    add_body_p("• Thực hiện vệ sinh công nghiệp máy trạm, bàn ghế phòng học và kiểm tra hệ thống điện lạnh đảm bảo điều kiện học tập tốt nhất cho sinh viên.")

    # ==============================================================
    # PHẦN IV: ĐÁNH GIÁ CHUNG & KIẾN NGHỊ
    # ==============================================================
    add_heading_1("IV. ĐÁNH GIÁ CHUNG, TỒN ĐỌNG & ĐỀ XUẤT KIẾN NGHỊ")

    add_body_p(
        "Toàn bộ 8 phòng máy được vận hành an toàn, đúng tiến độ đào tạo của Nhà trường. "
        "Việc số hóa công tác đăng ký mượn phòng, báo vắng và tiếp nhận sự cố kỹ thuật trực tuyến đã giúp giảm thiểu thời gian xử lý và phục vụ người dùng minh bạch, chính xác.",
        bold_prefix="1. Đánh giá chung: "
    )

    add_body_p(
        "Một số thiết bị ngoại vi (chuột, bàn phím) tại các phòng máy cụm P.5.x đã hao mòn qua nhiều học kỳ cần được thay thế định kỳ để đảm bảo hiệu suất thực hành môn học.",
        bold_prefix="2. Khó khăn, tồn đọng: "
    )

    add_body_p(
        "Kính đề xuất Lãnh đạo Phòng Quản trị Thiết bị xem xét phê duyệt mua sắm bổ sung thiết bị dự phòng (chuột quang, bàn phím, dây mạng đúc sẵn) "
        "và bảo dưỡng hệ thống điều hòa nhiệt độ tại 8 phòng máy trước đợt thi học kỳ sắp tới.",
        bold_prefix="3. Đề xuất kiến nghị: "
    )

    # ==============================================================
    # CHỮ KÝ & NƠI NHẬN
    # ==============================================================
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    tbl_sig = doc.add_table(rows=1, cols=2)
    tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_sig.autofit = False
    tbl_sig.columns[0].width = Inches(3.2)
    tbl_sig.columns[1].width = Inches(3.4)

    # Nơi nhận
    c_rec = tbl_sig.cell(0, 0)
    p_rec = c_rec.paragraphs[0]
    p_rec.paragraph_format.space_before = Pt(0)
    p_rec.paragraph_format.space_after = Pt(2)
    p_rec.paragraph_format.line_spacing = 1.15
    r_rn = p_rec.add_run("Nơi nhận:\n")
    r_rn.font.name = "Times New Roman"
    r_rn.font.size = Pt(11)
    r_rn.font.bold = True
    r_rn.font.italic = True

    r_rn_list = p_rec.add_run("- Như kính gửi;\n- Lãnh đạo Phòng QTTB (để b/c);\n- Lưu: VT, Tổ Phòng máy B4.2.")
    r_rn_list.font.name = "Times New Roman"
    r_rn_list.font.size = Pt(10.5)

    # Người lập báo cáo
    c_sign = tbl_sig.cell(0, 1)
    p_sign = c_sign.paragraphs[0]
    p_sign.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sign.paragraph_format.space_before = Pt(0)
    p_sign.paragraph_format.space_after = Pt(4)
    p_sign.paragraph_format.line_spacing = 1.15
    r_sg_t = p_sign.add_run("NGƯỜI LẬP BÁO CÁO\n")
    r_sg_t.font.name = "Times New Roman"
    r_sg_t.font.size = Pt(12)
    r_sg_t.font.bold = True

    r_sg_sub = p_sign.add_run("(Ký và ghi rõ họ tên)\n\n\n\n")
    r_sg_sub.font.name = "Times New Roman"
    r_sg_sub.font.size = Pt(11)
    r_sg_sub.font.italic = True

    # Xóa viền bảng chữ ký
    for row in tbl_sig.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(r'''
                <w:tcBorders %s>
                    <w:top w:val="none"/>
                    <w:left w:val="none"/>
                    <w:bottom w:val="none"/>
                    <w:right w:val="none"/>
                </w:tcBorders>
            ''' % nsdecls('w'))
            tcPr.append(tcBorders)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output
