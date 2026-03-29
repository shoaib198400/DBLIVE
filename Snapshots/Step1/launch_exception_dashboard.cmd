@echo off
setlocal
cd /d "%~dp0"
start "Exception Dashboard Server" /min "C:\Users\31931190\AppData\Local\Microsoft\WindowsApps\python3.13.exe" -m streamlit run app.py --server.port 8502
timeout /t 4 /nobreak >nul
start "" http://localhost:8502
endlocal