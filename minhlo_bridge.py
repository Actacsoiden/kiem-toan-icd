import ctypes
import sys
import os

# 1. BẮT BUỘC CHẠY QUYỀN ADMIN (Lách luật chặn phím của Windows)
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # Tự động gọi bảng cấp quyền Admin nếu bạn quên không click chuột phải
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

import keyboard
import pyperclip
import webbrowser
import time
import re
import pygetwindow as gw
import pyautogui

print("="*60)
print("🚀 ROBOT F9 LOCAL - BẢN ĐẶC QUYỀN (AUTO-ADMIN & KHÓA ĐÍCH)")
print("="*60)

# 2. KHÓA CỨNG ĐƯỜNG DẪN GHI FILE (Khớp 100% với file Web)
SHARED_TXT_PATH = r"D:\KiemToanICD10\shared_code.txt"
LOCAL_URL = "http://localhost:8501"

last_run_time = 0
cooldown_duration = 2.0

# 3. HÀM ÉP TIÊU ĐIỂM XUYÊN QUA CÁC PHẦN MỀM KHÁC
def force_focus(window_title):
    try:
        for win in gw.getAllWindows():
            if window_title.lower() in win.title.lower():
                hwnd = win._hWnd
                # Nếu trang web đang bị thu nhỏ dưới thanh Taskbar -> Kéo lên
                if ctypes.windll.user32.IsIconic(hwnd):
                    ctypes.windll.user32.ShowWindow(hwnd, 9) 
                
                # Bơm phím Alt ngầm để lách luật khóa màn hình của Windows 10/11
                pyautogui.press('alt')
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                return True
    except Exception:
        pass
    return False

def on_f9_pressed():
    global last_run_time
    current_time = time.time()
    
    if current_time - last_run_time < cooldown_duration:
        return
    last_run_time = current_time
    
    print("\n[Tín hiệu] ⚡ Nhận lệnh F9...")
    pyperclip.copy("") 
    keyboard.send('ctrl+c')
    time.sleep(0.4)
    
    raw_text = pyperclip.paste().strip().upper()
    copied_text = re.sub(r'[^A-Z0-9.]', '', raw_text)
    
    if not copied_text or len(copied_text) > 10:
        print("⚠️ Hủy bỏ: Không bôi đen trúng mã bệnh.")
        return
        
    # Ghi mã trực tiếp vào ổ D
    try:
        with open(SHARED_TXT_PATH, "w", encoding="utf-8") as f:
            f.write(copied_text)
        print(f"✅ Ghi mã '{copied_text}' thành công vào hệ thống.")
    except Exception as e:
        print(f"❌ Lỗi ghi file: {e}")
        return

    # Lôi trình duyệt lên mặt trước và ấn F5
    if force_focus("Kiểm Toán") or force_focus("Streamlit") or force_focus("localhost"):
        time.sleep(0.2)
        pyautogui.press('f5')
        print("🔄 Đã làm mới trang tra cứu!")
    else:
        webbrowser.open(LOCAL_URL)
        print("🌐 Mở tab tra cứu mới!")

keyboard.add_hotkey('f9', on_f9_pressed)
keyboard.wait()