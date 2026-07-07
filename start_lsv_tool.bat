@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 scripts\lsv_tool_app.py
  goto :end
)

where python >nul 2>nul
if %errorlevel%==0 (
  python scripts\lsv_tool_app.py
  goto :end
)

echo Python was not found.
echo Please install Python 3.10 or newer and tick "Add python.exe to PATH".
echo Then run:
echo pip install -r requirements.txt

:end
pause
