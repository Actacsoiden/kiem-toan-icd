import keyboard
import pyperclip
import webbrowser
import time
import re
import pygetwindow as gw
import pyautogui

print("="*60)
print("🚀 ROBOT F9 CLOUD INTERNET - BẢN PHÁT HÀNH TOÀN VIỆN")
print("="*60)

# ⚠️ ANH HÃY THAY ĐƯỜNG LINK TRANG WEB STREAMLIT THỰC TẾ ĐÃ HOẠT ĐỘNG CỦA ANH VÀO ĐÂY:
CLOUD_URL = "https://icd-vinhbao.streamlit.app" 

last_run_time = 0
cooldown_duration = 2.0  # Khóa phím 2 giây để chống lỗi lặp lệnh đẻ nhiều tab rác

def on_f9_pressed():
    global last_run_time
    current_time = time.time()
    
    if current_time - last_run_time < cooldown_duration:
        return
    last_run_time = current_time
    
    print("\n[Tín hiệu] ⚡ Đã nhận lệnh bắt mã F9...");
    pyperclip.copy("") 
    
    # Kích hoạt tổ hợp phím copy chuỗi ký tự đang bôi đen trên Minh Lộ
    keyboard.send('ctrl+c')
    time.sleep(0.4) # Chờ RAM máy trạm xử lý nạp dữ liệu
    
    raw_text = pyperclip.paste().strip().upper()
    copied_text = re.sub(r'[^A-Z0-9.]', '', raw_text)
    
    if not copied_text:
        print("⚠️ Hủy lệnh: Chuỗi rỗng (Bác sĩ chưa bôi đen trúng mã).")
        return
    if len(copied_text) > 10:
        print("⚠️ Hủy lệnh: Chuỗi quá dài, không phải mã bệnh.")
        return
        
    print(f"✅ Bóc tách thành công mã: {copied_text}")
    full_url = f"{CLOUD_URL}/?code={copied_text}"

    # Tìm kiếm tab trình duyệt Chrome/Edge đang mở sẵn trang web của anh
    target_win = None
    for win in gw.getAllWindows():
        if "Cổng Kiểm Toán" in win.title or "Streamlit" in win.title:
            target_win = win
            break
            
    if target_win:
        try:
            if target_win.isMinimized: 
                target_win.restore()
            pyautogui.press('alt') 
            target_win.activate()
            time.sleep(0.3)
            
            # Cơ chế liên thông URL thông minh dán đè vào thanh địa chỉ (Chống đẻ tab mới)
            pyautogui.hotkey('ctrl', 'l') 
            time.sleep(0.1)
            pyperclip.copy(full_url)
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.1)
            pyautogui.press('enter')
            print("🔄 Đã đẩy dữ liệu trực tuyến lên Tab cũ thành công!")
        except Exception:
            webbrowser.open(full_url)
    else:
        print("🌐 Khởi tạo tab tra cứu trực tuyến mới...")
        webbrowser.open(full_url)

keyboard.add_hotkey('f9', on_f9_pressed)
keyboard.wait()