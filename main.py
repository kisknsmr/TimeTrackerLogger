#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Activity Logger (Tray Version v2.3)
-----------------------------------
・タスクトレイ常駐型（CUI非表示）
・アクティブウィンドウを Logs/YYYY-MM-DD_HHh.csv に6時間ごと記録
・session + id 付きでViewer側の差分更新に対応
・右クリックメニューから記録間隔(1/3/5/10秒)を可変設定可能
・設定は settings.json に永続化
・BOM付きUTF-8出力（Excel互換）
"""

import time
import csv
import threading
import os
import sys
import uuid
import json
import psutil
import win32gui
import win32process
from PIL import Image, ImageDraw
import pystray

# ===============================
# 多重起動防止
# ===============================
def check_already_running():
    current_name = os.path.basename(sys.argv[0])
    current_pid = os.getpid()
    for proc in psutil.process_iter(['pid', 'cmdline']):
        try:
            if proc.info['pid'] != current_pid and proc.info['cmdline']:
                if current_name in " ".join(proc.info['cmdline']):
                    print("既に起動中です。多重起動を防止します。")
                    sys.exit(0)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

check_already_running()

# ===============================
# 定数・設定
# ===============================
SETTINGS_FILE = "settings.json"
DEFAULT_INTERVAL = 3
MAX_REPEAT_SEC = 300  # 同一ウィンドウ再記録の最長間隔
stop_flag = False
SESSION_ID = str(uuid.uuid4())[:8]

# ===============================
# 設定ロード・保存
# ===============================
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {"interval": DEFAULT_INTERVAL}

def save_settings(settings):
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)

settings = load_settings()
interval_value = settings.get("interval", DEFAULT_INTERVAL)
interval_lock = threading.Lock()

# ===============================
# ユーティリティ
# ===============================
def get_active_window_info():
    try:
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            title = win32gui.GetWindowText(hwnd)
            exe = process.name()
            return exe, title
    except Exception:
        pass
    return None, None

def safe_open_log():
    for _ in range(3):
        try:
            return open(current_log_path(), "a", newline="", encoding="utf-8-sig")
        except PermissionError:
            time.sleep(1)
    raise PermissionError("ログファイルにアクセスできませんでした。")

def current_log_path():
    now = time.localtime()
    base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(base_dir, "Logs")
    os.makedirs(log_dir, exist_ok=True)
    hour_block = (now.tm_hour // 6) * 6
    return os.path.join(log_dir, f"activity_{time.strftime('%Y-%m-%d')}_{hour_block:02d}h.csv")

# ===============================
# ロガースレッド
# ===============================
def logger_thread():
    global stop_flag
    last_exe, last_title = None, None
    last_write_time = 0
    current_file = None
    writer = None
    seq_id = 0
    current_path = None

    while not stop_flag:
        try:
            log_path = current_log_path()
            if log_path != current_path:
                if current_file:
                    current_file.close()
                current_file = safe_open_log()
                writer = csv.writer(current_file)
                if os.stat(log_path).st_size == 0:
                    writer.writerow(["session", "id", "timestamp", "exe", "title"])
                seq_id = 0
                current_path = log_path
                last_exe, last_title, last_write_time = None, None, 0

            exe, title = get_active_window_info()
            now = time.time()
            with interval_lock:
                sleep_int = interval_value
            if exe != last_exe or title != last_title or (now - last_write_time > MAX_REPEAT_SEC):
                seq_id += 1
                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                writer.writerow([SESSION_ID, seq_id, timestamp, exe or "", title or ""])
                current_file.flush()
                last_exe, last_title, last_write_time = exe, title, now

            time.sleep(sleep_int)

        except Exception as e:
            print(f"[WARN] ログ記録中に例外: {e}")
            time.sleep(1)

    if current_file:
        current_file.close()

# ===============================
# トレイメニュー構成
# ===============================
def create_icon():
    img = Image.new("RGB", (64, 64), color=(30, 144, 255))
    d = ImageDraw.Draw(img)
    d.ellipse((8, 8, 56, 56), fill=(255, 255, 255))
    return img

def set_interval(value):
    global interval_value
    with interval_lock:
        interval_value = value
        settings["interval"] = value
        save_settings(settings)
    print(f"[INFO] 記録間隔を {value} 秒に変更しました。")

def build_menu():
    def interval_item(sec):
        return pystray.MenuItem(
            f"{sec} sec", lambda: set_interval(sec), checked=lambda item: interval_value == sec
        )

    return pystray.Menu(
        pystray.MenuItem("Open Log Folder", lambda: os.startfile(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logs"))),
        pystray.MenuItem("Record Interval", pystray.Menu(interval_item(1), interval_item(3), interval_item(5), interval_item(10))),
        pystray.MenuItem("Exit", on_exit)
    )

def on_exit(icon, item):
    global stop_flag
    stop_flag = True
    icon.stop()

def main():
    t = threading.Thread(target=logger_thread, daemon=True)
    t.start()

    icon = pystray.Icon(
        "ActivityLogger",
        create_icon(),
        f"Activity Logger (Session {SESSION_ID})",
        menu=build_menu(),
    )
    icon.run()

if __name__ == "__main__":
    main()