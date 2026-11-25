#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
B42Lab NetAgent (Hybrid Push Collector) — Thu thập 8 Router tại Trường & Đẩy lên PythonAnywhere / Localhost
----------------------------------------------------------------------------------------------------------
Chạy trên 1 máy kỹ thuật tại Trường (VD: 10.29.37.45):
- Tự động đọc cấu hình & mật khẩu đã mã hóa từ CSDL cục bộ (không đẩy mật khẩu lên GitHub/Cloud).
- Quét song song cả 8 Router (2x GCC6010 WebSocket + 6x AirPro SSH Read-Only), lọc PC Phòng theo MAC OUI.
- Đẩy toàn bộ dữ liệu telemetry lên Server đích qua HTTPS POST /api/network-heartbeat.
- Tự đồng bộ lại danh sách IP/Port/Model và pc_mac_ouis nếu Admin thay đổi trên Web Cloud.

Cách sử dụng:
  python b42_network_agent.py --server https://b42lab.pythonanywhere.com --key b42lab-secret-2026
  python b42_network_agent.py --server http://127.0.0.1:5000 --once
"""

import argparse
import json
import os
import platform
import re
import subprocess
import time
import urllib.request

DEFAULT_SERVER_URL = "http://127.0.0.1:5000"
DEFAULT_AGENT_KEY = "b42lab-secret-2026"

ROOM_TARGETS = {
    "5.1": "10.29.37.53",
    "5.2": "10.29.37.54",
    "5.3": "10.29.37.52",
    "6.1": "10.29.37.60",
    "6.2": "10.29.37.62",
    "6.3": "10.29.37.64",
    "6.5": "10.29.37.63",
    "6.6": "10.29.37.61",
}

try:
    from database import get_db_connection, decrypt_router_secret
    from grandstream_router_monitor import poll_all_routers_concurrently, DEFAULT_ROOM_MODELS
    HAS_FULL_MONITOR = True
except Exception:
    HAS_FULL_MONITOR = False


def ping_host(target_ip):
    """Đo ping (ms) tới địa chỉ IP trên Windows / Linux / macOS (dự phòng khi chạy độc lập)."""
    target_ip = (target_ip or "8.8.8.8").strip()
    is_win = platform.system().lower().startswith("win")
    cmd = ["ping", "-n", "1", "-w", "1000", target_ip] if is_win else ["ping", "-c", "1", "-W", "1", target_ip]
    try:
        t0 = time.time()
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=2.5)
        elapsed_ms = max(1, int(round((time.time() - t0) * 1000)))
        if res.returncode == 0:
            m = re.search(r"(?:time|thời gian)[=<](\d+)\s*ms", res.stdout, re.IGNORECASE)
            ping_ms = int(m.group(1)) if m else elapsed_ms
            return "online", ping_ms
        return "offline", None
    except Exception:
        return "offline", None


def load_local_collector_config(synced_configs=None, synced_ouis=None):
    """Đọc cấu hình router & mật khẩu giải mã từ DB cục bộ (hoặc biến môi trường)."""
    cfg = {
        "r_user": os.environ.get("GCC_ADMIN_USER", "admin"),
        "r_pass": os.environ.get("GCC_ADMIN_PASS", ""),
        "a_user": os.environ.get("AIRPRO_ADMIN_USER", "admin"),
        "a_pass": os.environ.get("AIRPRO_ADMIN_PASS", ""),
        "pc_mac_ouis": synced_ouis or "4C:D7:17,00:68:EB,80:E8:2C",
        "room_configs": synced_configs or [],
    }
    if not HAS_FULL_MONITOR:
        return cfg

    try:
        conn = get_db_connection()
        rows = conn.execute(
            "SELECT key, value FROM system_settings WHERE key IN ('router_admin_user', 'router_admin_pass', 'airpro_admin_user', 'airpro_admin_pass', 'pc_mac_ouis')"
        ).fetchall()
        smap = {r["key"]: (r["value"] or "") for r in rows}
        cfg["r_user"] = smap.get("router_admin_user") or cfg["r_user"]
        if not cfg["r_pass"]:
            cfg["r_pass"] = decrypt_router_secret(smap.get("router_admin_pass") or "")
        cfg["a_user"] = smap.get("airpro_admin_user") or cfg["a_user"]
        if not cfg["a_pass"]:
            cfg["a_pass"] = decrypt_router_secret(smap.get("airpro_admin_pass") or "")
        if not synced_ouis and smap.get("pc_mac_ouis"):
            cfg["pc_mac_ouis"] = smap["pc_mac_ouis"]

        if not cfg["room_configs"]:
            rc_rows = conn.execute("SELECT room, target_ip, https_port, router_model FROM room_network_status ORDER BY room").fetchall()
            cfg["room_configs"] = [dict(r) for r in rc_rows]
        conn.close()
    except Exception as ex:
        print(f"[Cảnh báo] Không đọc được cấu hình DB cục bộ: {ex}")

    if not cfg["room_configs"]:
        cfg["room_configs"] = [
            {
                "room": r_code,
                "target_ip": ip,
                "router_model": DEFAULT_ROOM_MODELS.get(r_code, "AIRPRO") if HAS_FULL_MONITOR else "AIRPRO",
                "https_port": 443 if (HAS_FULL_MONITOR and DEFAULT_ROOM_MODELS.get(r_code) == "GCC6010") else 80,
            }
            for r_code, ip in ROOM_TARGETS.items()
        ]
    return cfg


def push_heartbeat(server_url, agent_key, rooms_payload):
    """Đẩy dữ liệu trạng thái mạng & telemetry lên B42Lab Server và nhận cấu hình đồng bộ."""
    url = f"{server_url.rstrip('/')}/api/network-heartbeat"
    body = json.dumps({
        "agent_key": agent_key,
        "rooms": rooms_payload
    }, ensure_ascii=False).encode("utf-8")

    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "User-Agent": "B42Lab-NetAgent/2.0-Hybrid"
        },
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode("utf-8"))


def main():
    parser = argparse.ArgumentParser(description="B42Lab NetAgent 2.0 - Hybrid Push Collector")
    parser.add_argument("--server", default=DEFAULT_SERVER_URL, help="URL của web B42Lab (VD: https://b42lab.pythonanywhere.com)")
    parser.add_argument("--key", default=DEFAULT_AGENT_KEY, help="Mã khóa bảo mật NetAgent")
    parser.add_argument("--room", default="", help="Chỉ định 1 phòng nếu chạy lẻ (VD: 6.1)")
    parser.add_argument("--interval", type=int, default=30, help="Chu kỳ quét & đẩy tính bằng giây (mặc định: 30s)")
    parser.add_argument("--once", action="store_true", help="Chỉ quét và đẩy 1 lần rồi thoát")
    args = parser.parse_args()

    mode_label = "Full Hybrid Collector (GCC6010 WS + AirPro SSH)" if HAS_FULL_MONITOR else "Ping-Only Fallback"
    print("=== B42Lab NetAgent 2.0 (Hybrid Push) ===")
    print(f"-> Server đích : {args.server}")
    print(f"-> Chế độ      : {mode_label}")
    print(f"-> Chu kỳ      : {args.interval} giây" + (" (Chạy 1 lần)" if args.once else ""))

    synced_configs = None
    synced_ouis = None

    while True:
        try:
            if HAS_FULL_MONITOR:
                cfg = load_local_collector_config(synced_configs=synced_configs, synced_ouis=synced_ouis)
                target_rooms = [rc for rc in cfg["room_configs"] if not args.room or rc["room"] == args.room]
                rooms_payload = poll_all_routers_concurrently(
                    target_rooms,
                    username=cfg["r_user"],
                    password=cfg["r_pass"],
                    airpro_user=cfg["a_user"],
                    airpro_pass=cfg["a_pass"],
                    pc_mac_ouis=cfg["pc_mac_ouis"],
                )
            else:
                rooms_payload = {}
                if args.room:
                    status, ping_ms = ping_host(ROOM_TARGETS.get(args.room, "8.8.8.8"))
                    rooms_payload[args.room] = {"status": status, "ping_ms": ping_ms}
                else:
                    for room_code, target_ip in ROOM_TARGETS.items():
                        status, ping_ms = ping_host(target_ip)
                        rooms_payload[room_code] = {"status": status, "ping_ms": ping_ms}

            resp_data = push_heartbeat(args.server, args.key, rooms_payload)
            if isinstance(resp_data.get("room_configs"), list) and resp_data["room_configs"]:
                synced_configs = resp_data["room_configs"]
            if resp_data.get("pc_mac_ouis"):
                synced_ouis = resp_data["pc_mac_ouis"]
            if isinstance(resp_data.get("targets"), dict):
                for r_code, new_ip in resp_data["targets"].items():
                    if r_code in ROOM_TARGETS and new_ip:
                        ROOM_TARGETS[r_code] = new_ip

            online_cnt = sum(1 for v in rooms_payload.values() if v.get("status") == "online")
            total_pc = sum(int(v.get("online_clients") or 0) for v in rooms_payload.values())
            print(f"[{resp_data.get('updated_at', '')}] Đã đẩy {len(rooms_payload)} phòng ({online_cnt} Online • {total_pc} PC phòng đang mở).")
        except Exception as err:
            print(f"[Cảnh báo] Lỗi khi quét/đẩy tới {args.server}: {err}")

        if args.once:
            break
        time.sleep(max(10, args.interval))


if __name__ == "__main__":
    main()
