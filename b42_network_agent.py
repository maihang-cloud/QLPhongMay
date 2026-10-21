#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
B42Lab NetAgent — Công cụ tự động kiểm tra Internet / Mạng 8 Phòng máy (Heartbeat Push)
---------------------------------------------------------------------------------------
Chạy ngầm nhẹ nhàng (không cần cài thêm thư viện ngoài), tự động đo Ping (ms) của 8 phòng
và đẩy kết quả lên Dashboard Quản trị (PythonAnywhere hoặc Localhost) mỗi 30 giây.

Cách sử dụng:
  1. Chạy kiểm tra cả 8 phòng từ 1 máy trung tâm:
     python b42_network_agent.py --server https://b42lab.pythonanywhere.com

  2. Hoặc chạy riêng tại 1 phòng máy (VD: đặt ở máy giáo viên Phòng 5.1):
     python b42_network_agent.py --server https://b42lab.pythonanywhere.com --room 5.1
"""

import argparse
import json
import platform
import re
import subprocess
import time
import urllib.request

DEFAULT_SERVER_URL = "http://127.0.0.1:5000"
DEFAULT_AGENT_KEY = "b42lab-secret-2026"

# Danh sách IP Router của 8 phòng (đồng bộ trực tiếp với cấu hình Admin trên Web)
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


def ping_host(target_ip):
    """Đo ping (ms) tới địa chỉ IP trên Windows / Linux / macOS."""
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


def push_heartbeat(server_url, agent_key, rooms_payload):
    """Đẩy dữ liệu trạng thái mạng lên B42Lab Server và nhận về danh sách IP mới nhất."""
    url = f"{server_url.rstrip('/')}/api/network-heartbeat"
    body = json.dumps({
        "agent_key": agent_key,
        "rooms": rooms_payload
    }).encode("utf-8")

    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type": "application/json",
            "User-Agent": "B42Lab-NetAgent/1.0"
        },
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=8) as resp:
        return json.loads(resp.read().decode("utf-8"))


def main():
    parser = argparse.ArgumentParser(description="B42Lab NetAgent - Heartbeat Push Monitor")
    parser.add_argument("--server", default=DEFAULT_SERVER_URL, help="URL của web B42Lab (VD: https://b42lab.pythonanywhere.com)")
    parser.add_argument("--key", default=DEFAULT_AGENT_KEY, help="Mã khóa bảo mật NetAgent")
    parser.add_argument("--room", default="", help="Chỉ định 1 phòng nếu chạy lẻ từng phòng (VD: 5.1)")
    parser.add_argument("--interval", type=int, default=30, help="Chu kỳ kiểm tra tính bằng giây (mặc định: 30s)")
    args = parser.parse_args()

    print(f"=== B42Lab NetAgent đang chạy ===")
    print(f"-> Server đích : {args.server}")
    print(f"-> Chu kỳ đo   : {args.interval} giây")
    if args.room:
        print(f"-> Chế độ      : Đo riêng Phòng {args.room}")
    else:
        print(f"-> Chế độ      : Đo tập trung cả 8 Phòng máy (5.1 -> 6.6)")

    while True:
        try:
            rooms_payload = {}
            if args.room:
                status, ping_ms = ping_host("8.8.8.8")
                rooms_payload[args.room] = {"status": status, "ping_ms": ping_ms}
            else:
                for room_code, target_ip in ROOM_TARGETS.items():
                    status, ping_ms = ping_host(target_ip)
                    rooms_payload[room_code] = {"status": status, "ping_ms": ping_ms}

            resp_data = push_heartbeat(args.server, args.key, rooms_payload)
            # Cập nhật lại danh sách IP mục tiêu nếu Admin vừa đổi trên giao diện Web
            if isinstance(resp_data.get("targets"), dict):
                for r_code, new_ip in resp_data["targets"].items():
                    if r_code in ROOM_TARGETS and new_ip:
                        ROOM_TARGETS[r_code] = new_ip

            online_cnt = sum(1 for v in rooms_payload.values() if v["status"] == "online")
            print(f"[{resp_data.get('updated_at', '')}] Đã đẩy tín hiệu thành công: {online_cnt}/{len(rooms_payload)} phòng Online.")
        except Exception as err:
            print(f"[Cảnh báo] Không thể gửi tín hiệu tới {args.server}: {err}")

        time.sleep(max(5, args.interval))


if __name__ == "__main__":
    main()
