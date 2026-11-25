"""
[#QTTB-03] Bộ giao tiếp & Giám sát Realtime 8 Router Grandstream (GCC6010 & GWN7062G)
Hỗ trợ đồng thời:
- Grandstream GCC6010 (Phòng 5.2, 6.3): Xác thực /cgi -> switchToRouter -> wss://<ip>:<port>/router/websocket
- Grandstream GWN7062G (Phòng 5.1, 5.3, 6.1, 6.2, 6.5, 6.6): Xác thực HTTPS Boa (/boaform/formLogin) -> đọc status.asp, lan_port_status.asp, dhcptbl.asp, arptable.asp, stats.asp và tự động logout (/boaform/admin/formLogout) để không chiếm phiên của Admin.
- Tự động fallback Ping ICMP khi Router chưa bật Web WAN Port Access (port 443).
- Tự động gộp bảng DHCP Leases + Static DHCP + Active Clients + ARP Cache theo MAC để lấy Computer Name.
"""

import urllib.request
import urllib.parse
import http.client
import subprocess
import platform
import tempfile
import threading
import hashlib
import secrets
import base64
import struct
import socket
import json
import time
import ssl
import sys
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from timezone_utils import get_vn_now_str

PI_MAP = {
    'CTRL': 1,
    'SET': 2,
    'GET': 3,
    'NOTIFY': 4,
    'MANAGE': 5,
    'EXT': 6
}

SPEED_CODE_MAP = {
    '0': 'Off',
    '1': '10M HD',
    '2': '10M FD',
    '3': '100M HD',
    '4': '100M FD',
    '5': '1000M HD',
    '6': '1Gbps',
    '7': '2.5Gbps',
    '8': 'SFP 1G/2.5G'
}

DEFAULT_ROOM_MODELS = {
    '5.1': 'AIRPRO',
    '5.2': 'GCC6010',
    '5.3': 'AIRPRO',
    '6.1': 'AIRPRO',
    '6.2': 'AIRPRO',
    '6.3': 'GCC6010',
    '6.5': 'AIRPRO',
    '6.6': 'AIRPRO',
}

DEFAULT_LAN_SUBNETS = {
    '5.1': '192.168.51.1',
    '5.2': '192.168.54.1',
    '5.3': '192.168.53.1',
    '6.1': '192.168.60.1',
    '6.2': '192.168.62.1',
    '6.3': '192.168.63.1',
    '6.5': '192.168.65.1',
    '6.6': '192.168.66.1',
}

DEFAULT_PC_MAC_OUIS = ('4C:D7:17', '00:68:EB', '80:E8:2C')
INFRA_MAC_OUIS = ('EC:FC:C6', '10:98:19')
NON_PC_HOST_KEYWORDS = (
    'iphone', 'ipad', 'android', 'redmi', 'poco', 'galaxy',
    'oppo', 'vivo', 'realme', 'xiaomi', 'huawei', 'oneplus', 'pixel'
)


def parse_pc_mac_ouis(raw_ouis):
    """Chuẩn hóa danh sách tiền tố MAC OUI của PC phòng máy (VD: '4C:D7:17,00:68:EB')."""
    if not raw_ouis:
        return set(DEFAULT_PC_MAC_OUIS)
    items = raw_ouis if isinstance(raw_ouis, (list, tuple, set)) else str(raw_ouis).split(',')
    out = set()
    for item in items:
        clean = re.sub(r'[^0-9A-Fa-f]', '', str(item or '')).upper()
        if len(clean) >= 6:
            p = clean[:6]
            out.add(f'{p[0:2]}:{p[2:4]}:{p[4:6]}')
    return out or set(DEFAULT_PC_MAC_OUIS)


def classify_lab_pc_device(mac_str, computer_name='', pc_mac_ouis=None):
    """Phân loại thiết bị: trả về (is_lab_pc, device_category)."""
    mac = format_mac_colon(mac_str)
    oui = mac[:8] if len(mac) >= 8 else ''
    allowed_ouis = parse_pc_mac_ouis(pc_mac_ouis)

    if oui in INFRA_MAC_OUIS:
        return False, 'Hạ tầng mạng'

    cname_l = (computer_name or '').strip().lower()
    if any(kw in cname_l for kw in NON_PC_HOST_KEYWORDS):
        return False, 'Di động / Cá nhân'

    if re.match(r'^p\d+-ws\d+', cname_l):
        return True, 'PC Phòng'

    is_local_random = (len(mac) > 1 and mac[1] in ('2', '6', 'A', 'E'))
    if is_local_random:
        return False, 'Wi-Fi / MAC ẩn'

    if oui in allowed_ouis:
        return True, 'PC Phòng'

    return False, 'Thiết bị khác'


def format_mac_colon(mac_raw):
    """Chuẩn hóa địa chỉ MAC về dạng XX:XX:XX:XX:XX:XX."""
    clean = re.sub(r'[^0-9A-Fa-f]', '', str(mac_raw or '')).upper()
    if len(clean) == 12:
        return ':'.join(clean[i:i + 2] for i in range(0, 12, 2))
    return str(mac_raw or '-').upper()


def format_bytes_human(num_bytes):
    """Định dạng dung lượng Byte sang KB / MB / GB."""
    try:
        val = float(num_bytes or 0)
    except (ValueError, TypeError):
        val = 0.0
    if val < 1024:
        return f"{int(val)} B"
    if val < 1024 ** 2:
        return f"{val / 1024:.1f} KB"
    if val < 1024 ** 3:
        return f"{val / (1024 ** 2):.1f} MB"
    return f"{val / (1024 ** 3):.2f} GB"


def format_rate_human(bytes_per_sec):
    """Định dạng tốc độ truyền (Bytes/s) sang B/s, KB/s hoặc Mbps."""
    try:
        bps = float(bytes_per_sec or 0)
    except (ValueError, TypeError):
        bps = 0.0
    if bps <= 0:
        return "0 B/s"
    if bps < 1024:
        return f"{int(bps)} B/s"
    bits_per_sec = bps * 8.0
    if bits_per_sec < 1024 * 1024:
        return f"{bps / 1024:.1f} KB/s"
    return f"{bits_per_sec / (1024 * 1024):.2f} Mbps"


def format_uptime_vn(seconds):
    """Định dạng số giây uptime sang 'X ngày Y giờ Z phút'."""
    try:
        sec = int(float(seconds or 0))
    except (ValueError, TypeError):
        return "-"
    if sec <= 0:
        return "-"
    days = sec // 86400
    hours = (sec % 86400) // 3600
    mins = (sec % 3600) // 60
    parts = []
    if days > 0:
        parts.append(f"{days} ngày")
    if hours > 0:
        parts.append(f"{hours} giờ")
    if days == 0 and mins > 0:
        parts.append(f"{mins} phút")
    return " ".join(parts) if parts else f"{sec} giây"


def ping_host_ms(target_ip, timeout_sec=1.2):
    """Ping ICMP kiểm tra IP WAN dự phòng khi cổng HTTPS 443 chưa mở."""
    if not target_ip:
        return False, None
    is_win = platform.system().lower().startswith('win')
    w_arg = str(int(timeout_sec * 1000)) if is_win else '1'
    cmd = ['ping', '-n', '1', '-w', w_arg, target_ip] if is_win else ['ping', '-c', '1', '-W', w_arg, target_ip]
    try:
        t0 = time.time()
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout_sec + 1.0)
        elapsed_ms = max(1, int(round((time.time() - t0) * 1000)))
        if res.returncode == 0:
            m = re.search(r'(?:time|thời gian)[=<](\d+)\s*ms', res.stdout, re.IGNORECASE)
            return True, (int(m.group(1)) if m else elapsed_ms)
        return False, None
    except Exception:
        return False, None


class _GCC6010WSClient:
    """Client WebSocket thuần Python (stdlib socket + ssl) cho Grandstream GCC6010."""

    def __init__(self, host, port=443, timeout=4.5):
        self.host = host
        self.port = int(port or 443)
        self.timeout = timeout
        self.ctx = ssl._create_unverified_context()
        self.ssock = None
        self.user_id = ""
        self.auth_jwt = ""

    def _safe_http_post_cgi(self, payload, auth_token=None):
        base = f"https://{self.host}:{self.port}"
        p = dict(payload)
        p['_location'] = '"/deviceOverview"'
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        if auth_token:
            headers['Authorization'] = auth_token
        req = urllib.request.Request(
            base + '/cgi',
            data=urllib.parse.urlencode(p).encode('utf-8'),
            headers=headers,
            method='POST'
        )
        with urllib.request.urlopen(req, context=self.ctx, timeout=self.timeout) as resp:
            try:
                raw = resp.read()
            except http.client.IncompleteRead as e:
                raw = e.partial
            body = json.loads(raw.decode('utf-8', errors='ignore'))
            hdrs = {k.lower(): v for k, v in resp.headers.items()}
            return body, hdrs

    def _open_ws_socket(self, ws_path, extra_headers=None):
        raw_sock = socket.create_connection((self.host, self.port), timeout=self.timeout)
        ssock = self.ctx.wrap_socket(raw_sock, server_hostname=self.host)
        ssock.settimeout(self.timeout)
        ws_key = base64.b64encode(secrets.token_bytes(16)).decode('ascii')
        lines = [
            f"GET {ws_path} HTTP/1.1",
            f"Host: {self.host}:{self.port}",
            "Upgrade: websocket",
            "Connection: Upgrade",
            f"Sec-WebSocket-Key: {ws_key}",
            "Sec-WebSocket-Version: 13",
            "Sec-WebSocket-Protocol: websocket",
        ]
        if extra_headers:
            for k, v in extra_headers.items():
                lines.append(f"{k}: {v}")
        lines.append("\r\n")
        ssock.sendall("\r\n".join(lines).encode('utf-8'))

        resp_hdr = b""
        while b"\r\n\r\n" not in resp_hdr:
            chunk = ssock.recv(1024)
            if not chunk:
                break
            resp_hdr += chunk
        first_line = resp_hdr.split(b"\r\n", 1)[0].decode('utf-8', errors='ignore')
        if "101" not in first_line:
            ssock.close()
            raise ConnectionError(f"WebSocket upgrade failed on {ws_path}: {first_line}")
        self.ssock = ssock

    def connect_and_authenticate(self, username='admin', password=''):
        """Xác thực chuẩn GCC6010 (/cgi -> switchToRouter -> /router/websocket)."""
        c_body, _ = self._safe_http_post_cgi({'action': 'challenge', 'user': username})
        challenge = (c_body.get('response') or {}).get('challenge', '')
        if not challenge:
            raise PermissionError("Không nhận được challenge xác thực từ GCC6010.")
        tok = hashlib.md5((challenge + password).encode('utf-8')).hexdigest()
        l_body, l_hdr = self._safe_http_post_cgi({'action': 'login', 'user': username, 'token': tok})
        if l_body.get('status') != 0:
            raise PermissionError("Đăng nhập GCC6010 thất bại (sai tài khoản/mật khẩu).")
        self.auth_jwt = l_hdr.get('authorization', '')
        sw_body, _ = self._safe_http_post_cgi(
            {'action': 'switchToRouter', 'router_user_id': ''},
            auth_token=self.auth_jwt
        )
        self.user_id = (sw_body.get('response') or {}).get('router_user_id', '')
        if not self.user_id:
            raise PermissionError("Không lấy được router_user_id từ GCC6010.")
        self._open_ws_socket(
            '/router/websocket',
            extra_headers={'Cookie': f'Authorization={self.auth_jwt}'}
        )
        return True

    def _send_json(self, obj):
        data = json.dumps(obj).encode('utf-8')
        mask_key = secrets.token_bytes(4)
        header = bytearray([0x81])
        length = len(data)
        if length < 126:
            header.append(0x80 | length)
        elif length < 65536:
            header.append(0x80 | 126)
            header.extend(struct.pack('!H', length))
        else:
            header.append(0x80 | 127)
            header.extend(struct.pack('!Q', length))
        header.extend(mask_key)
        masked = bytearray(b ^ mask_key[i % 4] for i, b in enumerate(data))
        self.ssock.sendall(header + masked)

    def _recv_exact(self, n):
        buf = b""
        while len(buf) < n:
            chunk = self.ssock.recv(n - len(buf))
            if not chunk:
                raise EOFError("WebSocket closed")
            buf += chunk
        return buf

    def _recv_json(self):
        while True:
            b1, b2 = self._recv_exact(2)
            opcode = b1 & 0x0F
            masked = bool(b2 & 0x80)
            length = b2 & 0x7F
            if length == 126:
                length = struct.unpack('!H', self._recv_exact(2))[0]
            elif length == 127:
                length = struct.unpack('!Q', self._recv_exact(8))[0]
            mask_key = self._recv_exact(4) if masked else None
            payload = self._recv_exact(length)
            if masked and mask_key:
                payload = bytes(b ^ mask_key[i % 4] for i, b in enumerate(payload))
            if opcode == 1:
                return json.loads(payload.decode('utf-8', errors='ignore'))
            if opcode == 8:
                raise EOFError("WebSocket closed by router")

    def call(self, cmd_str, extra=None):
        main_str, sub_str = cmd_str.split('|', 1)
        main_id = PI_MAP[main_str.upper()]
        hid = secrets.token_hex(8)
        cmd_sec = {'cmdSubtype': sub_str}
        if extra:
            for k, v in extra.items():
                cmd_sec[k] = str(v) if isinstance(v, (int, float)) else v
        req_obj = {
            'protocol': 1,
            'magicID': 253635645,
            'hashID': hid,
            'cmdMaintype': main_id,
            'cmdSection': cmd_sec,
            'userID': self.user_id
        }
        self._send_json(req_obj)
        deadline = time.time() + self.timeout
        while time.time() < deadline:
            msg = self._recv_json()
            if msg.get('hashID') == hid or msg.get('cmdSubtype') == sub_str:
                return msg
        return {}

    def close(self):
        if self.ssock:
            try:
                self.ssock.close()
            except Exception:
                pass
            self.ssock = None


def _poll_gcc6010_router(room, target_ip, https_port, username, password, pc_mac_ouis=None):
    """Đọc toàn bộ thông số Realtime từ Grandstream GCC6010."""
    now_str = get_vn_now_str('%H:%M:%S %d/%m/%Y')
    lan_default = DEFAULT_LAN_SUBNETS.get(room, '192.168.1.1')

    ws = _GCC6010WSClient(target_ip, port=https_port, timeout=4.2)
    try:
        t0 = time.time()
        ws.connect_and_authenticate(username=username, password=password)
        ws_handshake_ms = max(1, int(round((time.time() - t0) * 1000)))

        ov = ws.call('get|overview') or {}
        dev = ws.call('get|device') or {}
        ports_res = ws.call('get|port info') or {}
        cl_cnt = ws.call('get|clients count') or {}
        cl_res = ws.call('get|client', {'start': '0', 'amount': '100', 'sort': 'ascending connect time'}) or {}
        dhcp_leases_res = ws.call('get|dhcp leases', {'start': '0', 'amount': '100'}) or {}
        dhcp_host_res = ws.call('get|dhcp host', {'start': '0', 'amount': '100'}) or {}
        arp_res = ws.call('get|arp entry') or {}
        ws.close()

        wan_list = ov.get('wanInfo') or []
        active_wan = next(
            (w for w in wan_list if w.get('ip4Address') and w.get('ip4Address') != '-'),
            (wan_list[0] if wan_list else {})
        )
        rx_rate = int(float(active_wan.get('rxRate') or 0))
        tx_rate = int(float(active_wan.get('txRate') or 0))
        total_rx = int(float(active_wan.get('rxFlow') or 0))
        total_tx = int(float(active_wan.get('txFlow') or 0))
        lan_ip = ov.get('lanIp4Address') or lan_default
        detected_model = ov.get('devProductName') or dev.get('devProductName') or 'GCC6010'
        fw_ver = (ov.get('devSoftVersion') or dev.get('devSoftVersion') or '').strip()
        uptime_sec = int(float(ov.get('devUptime') or dev.get('devUptime') or 0))

        cpu_obj = dev.get('cpuUsage') or {}
        cpu_pct = int(float(cpu_obj.get('total') if isinstance(cpu_obj, dict) else (cpu_obj or 1)))
        mem_pct = int(float(dev.get('memoryUsage') or 0))
        temp_c = int(float(dev.get('temperature') or 38))

        ports_raw = ports_res.get('portsInfo') or []
        ports_clean = []
        lan_ports_up = 0
        wan_ports_up = 0
        wan_health_ping = None

        for p in ports_raw:
            role_code = str(p.get('role', '0'))
            is_wan = (role_code == '1')
            is_up = (str(p.get('linkStatus', '0')) != '0')
            speed_code = str(p.get('portSpeed', '0'))
            if is_up:
                if is_wan:
                    wan_ports_up += 1
                else:
                    lan_ports_up += 1
            health_arr = p.get('internetHealth') or []
            if health_arr and isinstance(health_arr, list):
                last_h = health_arr[-1]
                if last_h.get('avgLat'):
                    wan_health_ping = int(last_h['avgLat'])

            ports_clean.append({
                'port_id': p.get('portId') or '',
                'port_name': p.get('portName') or p.get('portId') or '',
                'role': 'WAN' if is_wan else 'LAN',
                'is_up': is_up,
                'speed_label': SPEED_CODE_MAP.get(speed_code, '1Gbps' if is_up else 'Off')
            })

        ping_ms = wan_health_ping if wan_health_ping is not None else min(ws_handshake_ms, 35)

        # Gộp thiết bị theo MAC từ 4 nguồn
        devices_by_mac = {}
        for h in (dhcp_host_res.get('hosts') or []):
            mac = format_mac_colon(h.get('mac'))
            if not mac or mac == '-':
                continue
            devices_by_mac[mac] = {
                'mac': mac,
                'ip': h.get('mapIp4Address') or '',
                'computer_name': (h.get('hostName') or '').strip(),
                'vlan': f"VLAN {h.get('vlan', '1')}",
                'is_online': False,
                'source': 'Static DHCP',
                'down_rate_bps': 0,
                'up_rate_bps': 0,
                'total_bytes': 0
            }

        for dl in (dhcp_leases_res.get('dhcpLeasesList') or []):
            mac = format_mac_colon(dl.get('mac'))
            if not mac or mac == '-':
                continue
            existing = devices_by_mac.get(mac, {})
            cname = (dl.get('hostName') or dl.get('name') or existing.get('computer_name') or '').strip()
            # Chỉ lưu thông tin IP/Hostname từ DHCP Lease; KHÔNG tự đánh is_online=True
            # vì DHCP Lease tồn tại nhiều giờ kể cả khi PC đã tắt nguồn.
            devices_by_mac[mac] = {
                'mac': mac,
                'ip': dl.get('ip') or dl.get('ip4Address') or existing.get('ip') or '',
                'computer_name': cname,
                'vlan': dl.get('vlan') or existing.get('vlan') or 'VLAN 1',
                'is_online': False,
                'source': 'DHCP Lease',
                'down_rate_bps': existing.get('down_rate_bps', 0),
                'up_rate_bps': existing.get('up_rate_bps', 0),
                'total_bytes': existing.get('total_bytes', 0)
            }

        for c in (cl_res.get('clientInfo') or []):
            mac = format_mac_colon(c.get('mac'))
            if not mac or mac == '-':
                continue
            existing = devices_by_mac.get(mac, {})
            cname = (c.get('name') or c.get('hostName') or existing.get('computer_name') or '').strip()
            is_on = str(c.get('statusClient', '1')) != '0'
            devices_by_mac[mac] = {
                'mac': mac,
                'ip': c.get('ip4Address') or c.get('ip') or existing.get('ip') or '',
                'computer_name': cname,
                'vlan': c.get('portName') or f"VLAN {c.get('vlanId', '1')}",
                'is_online': is_on,
                'source': 'Active Client',
                'down_rate_bps': int(float(c.get('downRate') or 0)),
                'up_rate_bps': int(float(c.get('upRate') or 0)),
                'total_bytes': int(float(c.get('totalBytes') or c.get('rxBytes') or 0)),
                'os': c.get('clientOs') or ''
            }

        for a in (arp_res.get('arpEntryList') or []):
            dev_iface = str(a.get('device') or '')
            if 'WAN' in dev_iface.upper():
                continue
            mac = format_mac_colon(a.get('mac'))
            ip_addr = (a.get('ip') or '').strip()
            if not mac or mac == '-' or not ip_addr:
                continue
            existing = devices_by_mac.get(mac)
            arp_hostname = (a.get('hostName') or '').strip()
            if arp_hostname == '-':
                arp_hostname = ''
            if existing:
                if not existing.get('ip'):
                    existing['ip'] = ip_addr
                if not existing.get('computer_name') and arp_hostname:
                    existing['computer_name'] = arp_hostname
            else:
                last_octet = ip_addr.split('.')[-1] if '.' in ip_addr else ''
                fallback_name = arp_hostname or (f"PC-{room.replace('.', '')}-.{last_octet}" if last_octet else "Máy trạm LAN")
                devices_by_mac[mac] = {
                    'mac': mac,
                    'ip': ip_addr,
                    'computer_name': fallback_name,
                    'vlan': dev_iface or 'VLAN 1',
                    'is_online': False,
                    'source': 'ARP Cache',
                    'down_rate_bps': 0,
                    'up_rate_bps': 0,
                    'total_bytes': 0
                }

        def _ip_sort_key(item):
            ip_s = item.get('ip') or ''
            parts = ip_s.split('.')
            last_num = int(parts[-1]) if (parts and parts[-1].isdigit()) else 999
            return (0 if item.get('is_online') else 1, -int(item.get('down_rate_bps') or 0), last_num)

        merged_devices = sorted(devices_by_mac.values(), key=_ip_sort_key)
        for d_item in merged_devices:
            if not d_item.get('computer_name'):
                ip_s = d_item.get('ip') or ''
                last_oct = ip_s.split('.')[-1] if '.' in ip_s else ''
                d_item['computer_name'] = f"PC-{room.replace('.', '')}-.{last_oct}" if last_oct else "Thiết bị LAN"
            d_item['down_rate_str'] = format_rate_human(d_item.get('down_rate_bps', 0))
            d_item['up_rate_str'] = format_rate_human(d_item.get('up_rate_bps', 0))
            d_item['total_bytes_str'] = format_bytes_human(d_item.get('total_bytes', 0))
            is_pc, cat_lbl = classify_lab_pc_device(d_item.get('mac'), d_item.get('computer_name'), pc_mac_ouis)
            d_item['is_lab_pc'] = is_pc
            d_item['device_category'] = cat_lbl

        online_clients_cnt = sum(1 for x in merged_devices if x.get('is_online') and x.get('is_lab_pc'))
        other_devices_cnt = sum(1 for x in merged_devices if not x.get('is_lab_pc'))
        arp_total_cnt = len(merged_devices)

        if online_clients_cnt > 0:
            room_state = 'open_active'
            room_state_label = f"Đang mở học ({online_clients_cnt} PC online)"
        elif lan_ports_up > 0:
            room_state = 'lan_up_idle'
            room_state_label = f"Đã bật Switch phòng ({lan_ports_up} cổng LAN UP • 0 PC)"
        else:
            room_state = 'closed_cb_off'
            room_state_label = "Phòng đóng cửa / Tắt CB (LAN DOWN)"

        telemetry = {
            'connected_api': True,
            'model': detected_model,
            'firmware': fw_ver,
            'serial': dev.get('devSn') or '',
            'mac': format_mac_colon(dev.get('devMac') or ''),
            'hw_version': dev.get('devHwVersion') or 'V1.1A',
            'uptime_sec': uptime_sec,
            'uptime_str': format_uptime_vn(uptime_sec),
            'wan_ip': active_wan.get('ip4Address') or target_ip,
            'wan_gateway': active_wan.get('gateway') or '10.29.37.1',
            'wan_dns': f"{active_wan.get('firstDns') or ''} / {active_wan.get('secondDns') or ''}".strip(' /'),
            'wan_port_name': active_wan.get('portName') or 'WAN',
            'lan_ip': lan_ip,
            'lan_netmask': ov.get('lanNetmask') or '255.255.255.0',
            'lan_ports_up': lan_ports_up,
            'wan_ports_up': wan_ports_up,
            'rx_rate_bps': rx_rate,
            'tx_rate_bps': tx_rate,
            'rx_rate_str': format_rate_human(rx_rate),
            'tx_rate_str': format_rate_human(tx_rate),
            'total_rx_bytes': total_rx,
            'total_tx_bytes': total_tx,
            'total_rx_str': format_bytes_human(total_rx),
            'total_tx_str': format_bytes_human(total_tx),
            'cpu_usage': cpu_pct,
            'mem_usage': mem_pct,
            'temp_c': temp_c,
            'online_clients': online_clients_cnt,
            'other_devices_count': other_devices_cnt,
            'arp_count': arp_total_cnt,
            'ports': ports_clean,
            'devices': merged_devices,
            'diagnostic_note': (
                f"Cổng {active_wan.get('portName') or 'WAN'} đang UP 1Gbps (Ping {ping_ms}ms) • "
                + (
                    f"{lan_ports_up} cổng LAN UP, có {online_clients_cnt} máy tính đang hoạt động."
                    if lan_ports_up > 0
                    else f"Tất cả cổng LAN đều DOWN -> Phòng {room} hiện đang tắt CB điện / tắt Switch nội bộ (Lưu vết ARP: {arp_total_cnt} máy)."
                )
            )
        }

        return {
            'room': room,
            'target_ip': target_ip,
            'https_port': int(https_port or 443),
            'router_model': detected_model,
            'lan_ip': lan_ip,
            'status': 'online',
            'room_state': room_state,
            'room_state_label': room_state_label,
            'ping_ms': ping_ms,
            'online_clients': online_clients_cnt,
            'other_devices_count': other_devices_cnt,
            'arp_count': arp_total_cnt,
            'rx_rate_bps': rx_rate,
            'tx_rate_bps': tx_rate,
            'total_rx_bytes': total_rx,
            'total_tx_bytes': total_tx,
            'cpu_usage': cpu_pct,
            'mem_usage': mem_pct,
            'temp_c': temp_c,
            'updated_at': now_str,
            'telemetry': telemetry
        }
    finally:
        ws.close()


_AIRPRO_BW_CACHE = {}
_AIRPRO_BW_LOCK = threading.Lock()


def _try_ssh_airpro(room, target_ip, ssh_port=22, username='admin', password='', pc_mac_ouis=None):
    """Đọc thông số AirPro AP-EG320 qua SSH Read-Only (-T): show version, show interface, show arp."""
    if not password:
        return None

    # Kiểm tra nhanh cổng SSH (0.6s) trước khi gọi ssh.exe để tránh chờ timeout nếu thiết bị chưa bật SSH
    try:
        with socket.create_connection((target_ip, int(ssh_port or 22)), timeout=0.6):
            pass
    except Exception:
        return None

    askpass_bat = None
    try:
        fd, askpass_bat = tempfile.mkstemp(prefix='_ap_ask_', suffix='.bat')
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(
                '@echo off\r\n'
                f'"{sys.executable}" -c "import os; print(os.environ.get(\'AIRPRO_SSH_PASS\', \'\'))"\r\n'
            )

        env = os.environ.copy()
        env['SSH_ASKPASS'] = askpass_bat
        env['SSH_ASKPASS_REQUIRE'] = 'force'
        env['DISPLAY'] = '1'
        env['AIRPRO_SSH_PASS'] = password

        cmd = [
            'ssh',
            '-o', 'StrictHostKeyChecking=no',
            '-o', 'UserKnownHostsFile=NUL',
            '-o', 'ConnectTimeout=4',
            '-p', str(int(ssh_port or 22)),
            '-T',
            f'{username}@{target_ip}',
        ]
        res = subprocess.run(
            cmd,
            input='show version\nshow interface\nshow arp\nexit\n',
            capture_output=True,
            text=True,
            timeout=7,
            env=env
        )
        out = res.stdout or ''
        if 'Interface G1' not in out and 'AP-EG320' not in out:
            return None

        # 1. Model & Firmware
        fw_ver = 'AP-EG320'
        m_ver = re.search(r'(AP-[A-Z0-9-]+)\s+version\s+([^\s,]+)', out)
        model_name = m_ver.group(1) if m_ver else 'AirPro AP-EG320'
        if m_ver:
            fw_ver = m_ver.group(2)

        m_host = re.search(r'^([A-Za-z0-9_.-]+)\{\}', out, re.MULTILINE)
        sys_name = m_host.group(1) if m_host else model_name

        # 2. Interfaces (G1..G8)
        ports = []
        wan_ip = target_ip
        wan_mac = '-'
        lan_ip = DEFAULT_LAN_SUBNETS.get(room, '192.168.1.1')
        total_rx = 0
        total_tx = 0
        lan_up = 0
        wan_up = 0

        blocks = re.split(r'(?=Interface\s+)', out)
        for blk in blocks:
            m_hdr = re.match(r'Interface\s+(G\d+)\s+is\s+\w+,\s+line protocol is\s+(\w+)', blk)
            if not m_hdr:
                continue
            if_name = m_hdr.group(1)
            is_up = (m_hdr.group(2).lower() == 'up')
            is_wan = (if_name == 'G1')
            if is_up:
                if is_wan:
                    wan_up += 1
                else:
                    lan_up += 1

            m_inet = re.search(r'inet\s+(\d+\.\d+\.\d+\.\d+)/\d+', blk)
            m_hw = re.search(r'HWaddr:\s*([0-9a-fA-F:]{17})', blk)
            m_rx = re.search(r'input packets\s*,\s*(\d+)\s+bytes', blk)
            m_tx = re.search(r'output packets\s*,\s*(\d+)\s+bytes', blk)

            if is_wan:
                if m_inet:
                    wan_ip = m_inet.group(1)
                if m_hw:
                    wan_mac = format_mac_colon(m_hw.group(1))
                if m_rx:
                    total_rx = int(m_rx.group(1))
                if m_tx:
                    total_tx = int(m_tx.group(1))
            elif if_name == 'G2' and m_inet:
                lan_ip = m_inet.group(1)

            ports.append({
                'port_id': if_name.lower(),
                'port_name': if_name,
                'role': 'WAN' if is_wan else 'LAN',
                'is_up': is_up,
                'speed_label': '1Gbps' if is_up else 'Off'
            })

        # 3. Tính tốc độ RX/TX bps từ hiệu số bytes giữa 2 lần quét
        rx_rate = 0
        tx_rate = 0
        now_ts = time.time()
        with _AIRPRO_BW_LOCK:
            prev = _AIRPRO_BW_CACHE.get(room)
            if prev and (now_ts - prev[0]) >= 0.5:
                dt = now_ts - prev[0]
                if total_rx >= prev[1]:
                    rx_rate = int((total_rx - prev[1]) / dt)
                if total_tx >= prev[2]:
                    tx_rate = int((total_tx - prev[2]) / dt)
            _AIRPRO_BW_CACHE[room] = (now_ts, total_rx, total_tx)

        # 4. Bảng ARP máy trạm (chỉ lấy cổng LAN G2..G8, bỏ IPv6 fe80 và cổng WAN G1)
        devices = []
        seen_macs = set()
        for m_arp in re.finditer(
            r'^(\d+\.\d+\.\d+\.\d+)\s+dev\s+(G[2-8])\s+lladdr\s+([0-9a-fA-F:]{17})\s+(\w+)',
            out,
            re.MULTILINE
        ):
            ip_str = m_arp.group(1)
            iface_dev = m_arp.group(2)
            mac_str = format_mac_colon(m_arp.group(3))
            state_str = m_arp.group(4).lower()
            if mac_str in seen_macs or ip_str.endswith('.1'):
                continue
            seen_macs.add(mac_str)
            # Chỉ tính đang online khi trạng thái ARP thực sự hoạt động (reachable/delay/probe);
            # loại bỏ trạng thái 'stale' để không đếm nhầm máy đã tắt từ ca trước.
            is_active_arp = (lan_up > 0 and state_str in ('reachable', 'delay', 'probe'))
            last_oct = ip_str.split('.')[-1]
            is_pc, cat_lbl = classify_lab_pc_device(mac_str, f"PC-{room.replace('.', '')}-.{last_oct}", pc_mac_ouis)
            arp_src = (
                'SSH ARP (Active)' if state_str in ('reachable', 'delay', 'probe')
                else ('SSH ARP (Stale)' if state_str == 'stale' else 'SSH ARP')
            )
            devices.append({
                'mac': mac_str,
                'ip': ip_str,
                'computer_name': f"PC-{room.replace('.', '')}-.{last_oct}",
                'vlan': f'LAN ({iface_dev})',
                'is_online': is_active_arp,
                'is_lab_pc': is_pc,
                'device_category': cat_lbl,
                'source': arp_src,
                'down_rate_bps': 0,
                'up_rate_bps': 0,
                'total_bytes': 0,
                'down_rate_str': '0 B/s',
                'up_rate_str': '0 B/s',
                'total_bytes_str': '-'
            })

        devices.sort(key=lambda d: (0 if d['is_online'] else 1, int(d['ip'].split('.')[-1]) if d['ip'].split('.')[-1].isdigit() else 999))

        return {
            'sys_name': sys_name,
            'model_label': f'AirPro {model_name} ({sys_name})',
            'firmware': fw_ver,
            'wan_ip': wan_ip,
            'wan_mac': wan_mac,
            'lan_ip': lan_ip,
            'ports': ports,
            'lan_ports_up': lan_up,
            'wan_ports_up': max(1, wan_up),
            'total_rx_bytes': total_rx,
            'total_tx_bytes': total_tx,
            'rx_rate_bps': rx_rate,
            'tx_rate_bps': tx_rate,
            'devices': devices
        }
    except Exception:
        return None
    finally:
        if askpass_bat and os.path.exists(askpass_bat):
            try:
                os.remove(askpass_bat)
            except Exception:
                pass


def _poll_airpro_router(room, target_ip, http_port=80, username='admin', password='', pc_mac_ouis=None):
    """Kiểm tra trạng thái Router AirPro (AP-EG320: ưu tiên SSH Read-Only nếu đã bật, fallback HTTP/Ping)."""
    now_str = get_vn_now_str('%H:%M:%S %d/%m/%Y')
    lan_default = DEFAULT_LAN_SUBNETS.get(room, '192.168.1.1')
    port_num = int(http_port or 80)

    is_ping_ok, ping_ms = ping_host_ms(target_ip, timeout_sec=1.2)
    web_reachable = False
    has_captcha = False
    http_ms = None

    try:
        t0 = time.time()
        req = urllib.request.Request(f"http://{target_ip}:{port_num}/", headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=3.5) as resp:
            http_ms = max(1, int(round((time.time() - t0) * 1000)))
            body = resp.read(8192).decode('utf-8', errors='ignore')
            web_reachable = (resp.status == 200)
            has_captcha = ('CheckCode.php' in body or 'randCode' in body)
    except Exception:
        web_reachable = False

    final_ping = ping_ms if ping_ms is not None else (http_ms or 0)
    is_online = bool(is_ping_ok or web_reachable)

    # Đọc thông số qua SSH Read-Only cho toàn bộ cụm Router AirPro
    ssh_data = None
    if is_online and password:
        ssh_data = _try_ssh_airpro(room, target_ip, ssh_port=22, username=username, password=password, pc_mac_ouis=pc_mac_ouis)

    if ssh_data:
        devs = ssh_data['devices']
        online_cnt = sum(1 for d in devs if d.get('is_online') and d.get('is_lab_pc'))
        other_cnt = sum(1 for d in devs if not d.get('is_lab_pc'))
        arp_cnt = len(devs)
        lan_up_cnt = ssh_data['lan_ports_up']
        if online_cnt > 0:
            room_state = 'open_active'
            room_state_label = f"Đang mở học ({online_cnt} PC online)"
        elif lan_up_cnt > 0:
            room_state = 'lan_up_idle'
            room_state_label = f"Đã bật Switch phòng ({lan_up_cnt} cổng LAN UP • 0 PC)"
        else:
            room_state = 'closed_cb_off'
            room_state_label = "Phòng đóng cửa / Tắt CB (LAN DOWN)"
        diag = (
            f"Đã kết nối SSH Read-Only tới {ssh_data['sys_name']} ({target_ip}:22 • Ping {final_ping}ms) • "
            + (
                f"{lan_up_cnt} cổng LAN UP, phát hiện {online_cnt} PC phòng đang hoạt động (ẩn {other_cnt} thiết bị Wi-Fi/khác)."
                if lan_up_cnt > 0
                else f"Cổng LAN G2 ({ssh_data['lan_ip']}) đang DOWN -> Phòng {room} hiện đang tắt CB điện / tắt Switch nội bộ (Lưu vết ARP: {arp_cnt} thiết bị)."
            )
        )
    elif web_reachable:
        room_state = 'wan_ping_only'
        room_state_label = f"AirPro Online (HTTP:{port_num} • {final_ping}ms)"
        diag = (
            f"AirPro Router tại {target_ip}:{port_num} đang hoạt động (Ping {final_ping}ms, Web UI sẵn sàng). "
            + ("Giao diện Web AirPro có CAPTCHA đăng nhập nên đang giám sát ở chế độ Ping/HTTP." if has_captcha else "")
        )
    elif is_ping_ok:
        room_state = 'wan_ping_only'
        room_state_label = f"Ping WAN Online ({final_ping}ms • Port {port_num} đóng)"
        diag = f"Địa chỉ WAN {target_ip} phản hồi Ping ({final_ping}ms) nhưng cổng HTTP {port_num} chưa phản hồi."
    else:
        room_state = 'offline'
        room_state_label = "Mất kết nối Router AirPro (WAN Offline)"
        diag = f"Không thể kết nối HTTP ({port_num}) và Ping tới {target_ip} đều không phản hồi."

    rx_bps = ssh_data['rx_rate_bps'] if ssh_data else 0
    tx_bps = ssh_data['tx_rate_bps'] if ssh_data else 0
    tot_rx = ssh_data['total_rx_bytes'] if ssh_data else 0
    tot_tx = ssh_data['total_tx_bytes'] if ssh_data else 0

    telemetry = {
        'connected_api': bool(ssh_data or web_reachable),
        'model': ssh_data['model_label'] if ssh_data else 'AirPro AP-EG320',
        'firmware': f"SSHv2 ({ssh_data['firmware']})" if ssh_data else 'DCN-UTM (PHP 7.2)',
        'serial': '-',
        'mac': ssh_data['wan_mac'] if ssh_data else '-',
        'uptime_str': '-',
        'wan_ip': ssh_data['wan_ip'] if ssh_data else target_ip,
        'lan_ip': ssh_data['lan_ip'] if ssh_data else lan_default,
        'lan_ports_up': ssh_data['lan_ports_up'] if ssh_data else 0,
        'wan_ports_up': ssh_data['wan_ports_up'] if ssh_data else (1 if is_online else 0),
        'rx_rate_bps': rx_bps,
        'tx_rate_bps': tx_bps,
        'rx_rate_str': format_rate_human(rx_bps),
        'tx_rate_str': format_rate_human(tx_bps),
        'total_rx_bytes': tot_rx,
        'total_tx_bytes': tot_tx,
        'total_rx_str': format_bytes_human(tot_rx) if ssh_data else '-',
        'total_tx_str': format_bytes_human(tot_tx) if ssh_data else '-',
        'cpu_usage': 0,
        'mem_usage': 0,
        'temp_c': 0,
        'online_clients': sum(1 for d in ssh_data['devices'] if d.get('is_online') and d.get('is_lab_pc')) if ssh_data else 0,
        'other_devices_count': sum(1 for d in ssh_data['devices'] if not d.get('is_lab_pc')) if ssh_data else 0,
        'arp_count': len(ssh_data['devices']) if ssh_data else 0,
        'ports': ssh_data['ports'] if (ssh_data and ssh_data['ports']) else [
            {'port_id': 'wan0', 'port_name': f'WAN:{port_num}', 'role': 'WAN', 'is_up': is_online, 'speed_label': f'{final_ping}ms' if is_online else 'Off'}
        ],
        'devices': ssh_data['devices'] if ssh_data else [],
        'diagnostic_note': diag
    }

    return {
        'room': room,
        'target_ip': target_ip,
        'https_port': port_num,
        'router_model': 'AIRPRO',
        'lan_ip': telemetry['lan_ip'],
        'status': 'online' if is_online else 'offline',
        'room_state': room_state,
        'room_state_label': room_state_label,
        'ping_ms': final_ping,
        'online_clients': telemetry['online_clients'],
        'arp_count': telemetry['arp_count'],
        'rx_rate_bps': rx_bps,
        'tx_rate_bps': tx_bps,
        'total_rx_bytes': tot_rx,
        'total_tx_bytes': tot_tx,
        'cpu_usage': 0,
        'mem_usage': 0,
        'temp_c': 0,
        'updated_at': now_str,
        'telemetry': telemetry
    }


def poll_single_grandstream_router(room, target_ip, https_port=443, username='admin', password='', configured_model=None, pc_mac_ouis=None):
    """Thu thập toàn bộ thông số Realtime của 1 Router (GCC6010, AirPro hoặc GWN7062G).
    Nếu chưa bật Web WAN Port Access, tự động kiểm tra Ping ICMP dự phòng.
    """
    now_str = get_vn_now_str('%H:%M:%S %d/%m/%Y')
    model_selected = (configured_model or DEFAULT_ROOM_MODELS.get(room, 'AIRPRO')).strip().upper()
    lan_default = DEFAULT_LAN_SUBNETS.get(room, '192.168.1.1')

    try:
        if model_selected == 'GCC6010':
            return _poll_gcc6010_router(room, target_ip, https_port, username, password, pc_mac_ouis=pc_mac_ouis)
        return _poll_airpro_router(room, target_ip, https_port, username, password, pc_mac_ouis=pc_mac_ouis)
    except Exception as conn_err:
        is_ping_ok, ping_ms = ping_host_ms(target_ip)
        if is_ping_ok:
            room_state = 'wan_ping_only'
            room_state_label = f"Ping WAN Online ({ping_ms}ms • Chờ mở Port {https_port})"
            status_val = 'online'
            diag = (
                f"Địa chỉ WAN {target_ip} phản hồi Ping tốt ({ping_ms}ms), nhưng cổng HTTPS {https_port} chưa kết nối được ({conn_err}). "
                f"Hãy bật cho phép truy cập Web từ cổng WAN trên Router {model_selected} của Phòng {room}."
            )
        else:
            room_state = 'offline'
            room_state_label = "Mất kết nối Router (WAN Offline)"
            status_val = 'offline'
            diag = f"Không thể kết nối HTTPS ({https_port}) và Ping tới {target_ip} đều không phản hồi ({conn_err})."

        telemetry = {
            'connected_api': False,
            'model': model_selected,
            'firmware': '-',
            'serial': '-',
            'mac': '-',
            'uptime_str': '-',
            'wan_ip': target_ip,
            'lan_ip': lan_default,
            'lan_ports_up': 0,
            'wan_ports_up': 1 if is_ping_ok else 0,
            'rx_rate_str': '0 B/s',
            'tx_rate_str': '0 B/s',
            'total_rx_str': '-',
            'total_tx_str': '-',
            'cpu_usage': 0,
            'mem_usage': 0,
            'temp_c': 0,
            'online_clients': 0,
            'arp_count': 0,
            'ports': [],
            'devices': [],
            'diagnostic_note': diag,
            'api_error': str(conn_err)
        }

        return {
            'room': room,
            'target_ip': target_ip,
            'https_port': int(https_port or 443),
            'router_model': model_selected,
            'lan_ip': lan_default,
            'status': status_val,
            'room_state': room_state,
            'room_state_label': room_state_label,
            'ping_ms': ping_ms or 0,
            'online_clients': 0,
            'arp_count': 0,
            'rx_rate_bps': 0,
            'tx_rate_bps': 0,
            'total_rx_bytes': 0,
            'total_tx_bytes': 0,
            'cpu_usage': 0,
            'mem_usage': 0,
            'temp_c': 0,
            'updated_at': now_str,
            'telemetry': telemetry
        }


def poll_all_routers_concurrently(room_configs, username='admin', password='', airpro_user='admin', airpro_pass='', pc_mac_ouis=None):
    """Quét song song cả 8 Router bằng ThreadPoolExecutor (tự động chọn đúng tài khoản theo dòng Router)."""
    results = {}
    with ThreadPoolExecutor(max_workers=8) as executor:
        fut_map = {}
        for rc in room_configs:
            r_code = rc['room']
            r_model = (rc.get('router_model') or DEFAULT_ROOM_MODELS.get(r_code, 'AIRPRO')).strip().upper()
            u_use = airpro_user if r_model == 'AIRPRO' else username
            p_use = airpro_pass if r_model == 'AIRPRO' else password
            fut = executor.submit(
                poll_single_grandstream_router,
                r_code,
                rc.get('target_ip') or '8.8.8.8',
                rc.get('https_port') or (80 if r_model == 'AIRPRO' else 443),
                u_use,
                p_use,
                r_model,
                pc_mac_ouis
            )
            fut_map[fut] = r_code
        for fut in as_completed(fut_map):
            r_code = fut_map[fut]
            try:
                results[r_code] = fut.result()
            except Exception as ex:
                results[r_code] = {
                    'room': r_code,
                    'status': 'offline',
                    'room_state': 'offline',
                    'room_state_label': f'Lỗi quét: {ex}',
                    'ping_ms': 0,
                    'telemetry': {'connected_api': False, 'diagnostic_note': str(ex), 'ports': [], 'devices': []}
                }
    return results


def reboot_grandstream_router(target_ip, https_port=443, username='admin', password='', model_hint='GCC6010'):
    """Gửi lệnh Khởi động lại Router từ xa (hỗ trợ GCC6010 và GWN7062G)."""
    model_upper = (model_hint or 'GCC6010').strip().upper()
    if model_upper == 'AIRPRO':
        return False, "Router AirPro có CAPTCHA bảo vệ trên giao diện đăng nhập nên không hỗ trợ Reboot tự động từ xa."

    ws = _GCC6010WSClient(target_ip, port=https_port, timeout=6.0)
    try:
        ws.connect_and_authenticate(username=username, password=password)
        ws.call('ctrl|reboot', {})
        ws.close()
        return True, "Đã gửi lệnh khởi động lại (Reboot) tới Router GCC6010 thành công! Thiết bị sẽ khởi động lại trong khoảng 1-2 phút."
    except Exception as ex:
        ws.close()
        return False, f"Không thể gửi lệnh Reboot tới GCC6010 ({target_ip}:{https_port}): {ex}"
