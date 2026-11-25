import os
import unittest

import app as app_module
from database import (
    get_db_connection,
    encrypt_router_secret,
    decrypt_router_secret,
    get_or_create_env_secret,
)
from grandstream_router_monitor import (
    DEFAULT_ROOM_MODELS,
    DEFAULT_LAN_SUBNETS,
    classify_lab_pc_device,
)


class TestCoreSystem(unittest.TestCase):
    def setUp(self):
        self.client = app_module.app.test_client()
        with self.client.session_transaction() as sess:
            sess['user_email'] = 'mlhang@hcmus.edu.vn'
            sess['user_name'] = 'Admin QTTB'
            sess['is_admin'] = True
            sess['is_verified'] = True

    def test_sqlite_wal_mode(self):
        conn = get_db_connection()
        mode = conn.execute("PRAGMA journal_mode;").fetchone()[0]
        conn.close()
        self.assertEqual(mode.lower(), 'wal')

    def test_secret_encryption_and_tamper_resistance(self):
        sample = "TestSecret@2026!"
        enc = encrypt_router_secret(sample)
        self.assertTrue(enc.startswith("enc:v1:"))
        self.assertEqual(decrypt_router_secret(enc), sample)
        # Tamper with ciphertext/tag -> must fail safely to ''
        tampered = enc[:-4] + ("AAAA" if enc[-4:] != "AAAA" else "BBBB")
        self.assertEqual(decrypt_router_secret(tampered), "")

    def test_flask_secret_key_not_hardcoded_default(self):
        sk = get_or_create_env_secret('SECRET_KEY')
        self.assertTrue(len(sk) >= 24)
        self.assertNotEqual(app_module.app.secret_key, 'KHTN_TKB_Secret_Key_@2026')

    def test_room_models_and_subnets(self):
        self.assertEqual(DEFAULT_ROOM_MODELS['5.2'], 'GCC6010')
        self.assertEqual(DEFAULT_ROOM_MODELS['6.3'], 'GCC6010')
        for r in ('5.1', '5.3', '6.1', '6.2', '6.5', '6.6'):
            self.assertEqual(DEFAULT_ROOM_MODELS[r], 'AIRPRO')
        self.assertEqual(DEFAULT_LAN_SUBNETS['6.1'], '192.168.60.1')

    def test_admin_panel_and_static_assets(self):
        r_admin = self.client.get('/admin/panel')
        self.assertEqual(r_admin.status_code, 200)
        html = r_admin.get_data(as_text=True)
        self.assertIn('/static/css/admin_panel.css', html)
        self.assertIn('/static/js/admin_panel.js', html)
        self.assertIn('window.ADMIN_PANEL_CFG', html)

        # Ensure DB router passwords are NOT leaked in HTML DOM
        conn = get_db_connection()
        rows = conn.execute(
            "SELECT value FROM system_settings WHERE key IN ('router_admin_pass', 'airpro_admin_pass')"
        ).fetchall()
        conn.close()
        for row in rows:
            plain = decrypt_router_secret(row['value'] or '')
            if plain:
                self.assertNotIn(plain, html)

        r_css = self.client.get('/static/css/admin_panel.css')
        self.assertEqual(r_css.status_code, 200)
        r_js = self.client.get('/static/js/admin_panel.js')
        self.assertEqual(r_js.status_code, 200)
        js_txt = r_js.get_data(as_text=True)
        self.assertNotIn('{{', js_txt)
        self.assertNotIn('{%', js_txt)

    def test_pc_mac_classification(self):
        self.assertEqual(classify_lab_pc_device('4C:D7:17:A0:AB:12', 'PC-62-.11'), (True, 'PC Phòng'))
        self.assertEqual(classify_lab_pc_device('00:68:EB:C5:96:B6', 'P46-WS20'), (True, 'PC Phòng'))
        self.assertEqual(classify_lab_pc_device('EC:FC:C6:CA:44:50', 'PC-52-.75'), (False, 'Hạ tầng mạng'))
        self.assertEqual(classify_lab_pc_device('BE:4C:FC:59:23:31', 'iPhone'), (False, 'Di động / Cá nhân'))
        self.assertEqual(classify_lab_pc_device('EE:7B:6E:5E:41:FA', 'PC-62-.11'), (False, 'Wi-Fi / MAC ẩn'))

    def test_hybrid_heartbeat_push_full_telemetry(self):
        # Invalid key -> 403
        bad = self.client.post('/api/network-heartbeat', json={'agent_key': 'wrong-key', 'rooms': {}})
        self.assertEqual(bad.status_code, 403)

        # Valid key + full telemetry payload -> 200 & returns synced configs
        conn = get_db_connection()
        ak_row = conn.execute("SELECT value FROM system_settings WHERE key = 'netagent_key'").fetchone()
        conn.close()
        key = ak_row['value'] if (ak_row and ak_row['value']) else 'b42lab-secret-2026'

        payload = {
            'agent_key': key,
            'rooms': {
                '6.1': {
                    'status': 'online',
                    'ping_ms': 4,
                    'router_model': 'AIRPRO',
                    'https_port': 80,
                    'lan_ip': '192.168.60.1',
                    'room_state': 'open_active',
                    'room_state_label': 'Đang mở học (40 PC online)',
                    'online_clients': 40,
                    'arp_count': 42,
                    'telemetry': {'connected_api': True, 'ports': [], 'devices': []}
                }
            }
        }
        ok_resp = self.client.post('/api/network-heartbeat', json=payload)
        self.assertEqual(ok_resp.status_code, 200)
        res_json = ok_resp.get_json()
        self.assertTrue(res_json['success'])
        self.assertEqual(res_json['persisted_count'], 1)
        self.assertIn('pc_mac_ouis', res_json)
        self.assertEqual(len(res_json['room_configs']), 8)

    def test_ca_time_config_and_usage_report_excel(self):
        from usage_report_generator import get_ca_time_config, detect_current_ca_slot
        import datetime

        conn = get_db_connection()
        cfg = get_ca_time_config(conn)
        conn.close()
        self.assertEqual(cfg[1]['start'], '07:30')
        self.assertEqual(cfg[4]['end'], '17:00')

        slot, _, _ = detect_current_ca_slot(datetime.datetime(2026, 11, 23, 15, 10), cfg)
        self.assertEqual(slot, 4)

        r_exp = self.client.get('/admin/export-usage-report?mode=week')
        self.assertEqual(r_exp.status_code, 200)
        self.assertIn('spreadsheetml.sheet', r_exp.content_type)
        self.assertGreater(len(r_exp.data), 2000)

    def test_merge_full_session_and_waste_sheet(self):
        from usage_report_generator import merge_full_session_rows, build_usage_excel_report
        from openpyxl import load_workbook
        import io

        conn = get_db_connection()
        sample_raw = [
            {
                'log_date': '2026-11-23', 'day_of_week': 2, 'room': '6.3', 'ca_slot': 1,
                'ca_start_cfg': '07:30', 'ca_end_cfg': '10:00',
                'subject_name': 'Tin học cơ sở', 'class_code': '', 'department': 'Trung tâm tin học',
                'first_online_at': '07:35', 'active_minutes': 140, 'idle_minutes': 0,
                'last_pc_count': 38, 'peak_pc_count': 40, 'avg_pc_count': 36.0, 'sample_count': 10,
            },
            {
                'log_date': '2026-11-23', 'day_of_week': 2, 'room': '6.3', 'ca_slot': 2,
                'ca_start_cfg': '10:05', 'ca_end_cfg': '12:00',
                'subject_name': 'Tin học cơ sở', 'class_code': '', 'department': 'Trung tâm tin học',
                'first_online_at': '10:05', 'active_minutes': 110, 'idle_minutes': 0,
                'last_pc_count': 35, 'peak_pc_count': 39, 'avg_pc_count': 34.0, 'sample_count': 10,
            },
        ]
        merged = merge_full_session_rows(conn, sample_raw)
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0]['ca_label'], 'Ca 1 + Ca 2 (Sáng)')
        self.assertEqual(merged[0]['ca_time'], '07:30 - 12:00')
        self.assertEqual(merged[0]['active_minutes'], 250)
        self.assertEqual(merged[0]['peak_pc_count'], 40)

        buf, _ = build_usage_excel_report(conn, mode='week')
        conn.close()
        wb = load_workbook(io.BytesIO(buf.getvalue()))
        self.assertEqual(wb.sheetnames, ['1. De xuat Chuyen phong', '2. Chi tiet Buoi hoc', '3. Canh bao Lang phi'])
        ws1 = wb['1. De xuat Chuyen phong']
        h1 = [ws1.cell(row=4, column=c).value for c in range(1, 12)]
        self.assertIn('Số PC hoạt động', h1)
        self.assertIn('Tỷ lệ sử dụng (% / 40 PC)', h1)
        ws3 = wb['3. Canh bao Lang phi']
        headers = [ws3.cell(row=4, column=c).value for c in range(1, 10)]
        self.assertEqual(
            headers,
            ['Ngày & Ca học', 'Phòng', 'Mã lớp học', 'Môn học', 'Đơn vị', 'Số mẫu quét', 'Số PC hoạt động', 'Tỷ lệ sử dụng (% / 40 PC)', 'Mức độ cảnh báo']
        )

    def test_cancellation_and_issue_form_enhancements(self):
        self.assertEqual(app_module.get_vn_weekday_label('2026-11-23'), 'Thứ 2')
        self.assertEqual(app_module.get_vn_weekday_label('2026-11-22'), 'Chủ Nhật')

        conn = get_db_connection()
        cols = [r['name'] for r in conn.execute("PRAGMA table_info(cancellations)").fetchall()]
        conn.close()
        self.assertIn('department', cols)

        r_can = self.client.get('/thong-bao-nghi')
        self.assertEqual(r_can.status_code, 200)
        can_html = r_can.get_data(as_text=True)
        self.assertIn('cancel_day_display', can_html)
        self.assertIn('name="department"', can_html)
        self.assertIn('Khoa Công nghệ Thông tin', can_html)

        r_rep = self.client.get('/bao-cao-su-co')
        self.assertEqual(r_rep.status_code, 200)
        rep_html = r_rep.get_data(as_text=True)
        self.assertIn('<select class="form-select bg-light border-0 py-2" name="department" required>', rep_html)
        self.assertIn('id="issue_description"', rep_html)
        self.assertIn('Khoa Công nghệ Thông tin', rep_html)

        # Whitespace-only description must be rejected at backend
        conn = get_db_connection()
        before_cnt = conn.execute("SELECT COUNT(*) FROM issue_reports").fetchone()[0]
        conn.close()
        r_bad = self.client.post('/submit_issue', data={
            'room': '6.1',
            'machine_number': 'Máy Học Viên',
            'issue_category': 'Khác',
            'reporter_name': 'Tester',
            'department': 'Khoa Toán',
            'description': '   ',
        }, follow_redirects=False)
        self.assertEqual(r_bad.status_code, 302)
        conn = get_db_connection()
        after_cnt = conn.execute("SELECT COUNT(*) FROM issue_reports").fetchone()[0]
        conn.close()
        self.assertEqual(before_cnt, after_cnt)

    def test_network_status_and_scan_cooldown(self):
        r_st = self.client.get('/api/network-status')
        self.assertEqual(r_st.status_code, 200)
        data = r_st.get_json()
        self.assertTrue(data['success'])
        self.assertEqual(len(data['room_map']), 8)


if __name__ == '__main__':
    unittest.main()
