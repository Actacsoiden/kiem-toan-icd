' --- ĐOẠN MÃ TỰ ĐỘNG NÂNG QUYỀN ADMINISTRATOR CHẠY NGẦM ---
If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
    WScript.Quit
End If

' --- ĐỊNH VỊ THƯ MỤC GỐC CHỨA FILE CODE CHỐNG LỖI SYSTEM32 ---
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
CurrentDir = FSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = CurrentDir

' --- KHỞI CHẠY SERVER WEB (ẨN HOÀN TOÀN BẢNG ĐEN) ---
WshShell.Run "cmd /c streamlit run app_icd_pro.py", 0, False

' --- KHỞI CHẠY ROBOT F9 MINH LỘ (ẨN HOÀN TOÀN BẢNG ĐEN) ---
WshShell.Run "cmd /c python minhlo_bridge.py", 0, False