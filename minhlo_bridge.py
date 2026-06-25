import keyboard
import pyperclip
import webbrowser
import time
import os
import pygetwindow as gw
import pyautogui

print("="*60)
print("🚀 HỆ THỐNG CẦU NỐI F9 - BẢN VÁ ĐƯỜNG DẪN TUYỆT ĐỐI")
print("="*60)

# Lấy chính xác thư mục đang chứa file code này
current_dir = os.path.dirname(os.path.abspath(__file__))
txt_path = os.path.join(current_dir, "shared_code.txt")

def on_f9_pressed():
    print("\n[Bước 1] ⚡ Đã nhận phím F9!");
    pyperclip.copy("") 
    keyboard.send('ctrl+c')
    time.sleep(0.3) # Đợi Windows xử lý copy
    
    copied_text = pyperclip.paste().strip().upper()
    print(f"[Bước 2] 📋 Dữ liệu lấy được: '{copied_text}'")
    
    if not copied_text or len(copied_text) > 10:
        print("❌ LỖI: Dữ liệu không hợp lệ hoặc chưa bôi đen.")
        return

    try:
        # Ghi đè vào tệp trung gian bằng đường dẫn tuyệt đối
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(copied_text)
        print(f"✅ Đã lưu mã thành công vào: {txt_path}")
    except Exception as e:
        print(f"❌ LỖI Hệ thống: {e}")
        return
        
    print("[Bước 3] 🔍 Tìm Web và F5...")
    target_win = None
    for win in gw.getAllWindows():
        if "ICD-10" in win.title or "Streamlit" in win.title:
            target_win = win
            break
            
    if target_win:
        try:
            if target_win.isMinimized: 
                target_win.restore()
            pyautogui.press('alt') 
            target_win.activate()
            time.sleep(0.2)
            pyautogui.press('f5')
            print("🔄 Đã F5 trình duyệt thành công!")
        except Exception as e:
            print(f"⚠️ Không thể tự động F5: {e}")
    else:
        print("🔗 Mở tab mới...")
        webbrowser.open("http://localhost:8501")

keyboard.add_hotkey('f9', on_f9_pressed)
keyboard.wait()