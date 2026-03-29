@echo off
setlocal
set BASE=%~dp0
copy /Y "%BASE%app.py" "%BASE%..\..\app.py" >nul
if exist "%BASE%launch_exception_dashboard.cmd" copy /Y "%BASE%launch_exception_dashboard.cmd" "%BASE%..\..\launch_exception_dashboard.cmd" >nul
echo Step1 restore complete.
