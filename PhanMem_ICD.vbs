Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c streamlit run app_icd_pro.py", 0, False