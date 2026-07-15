@echo off
py -3 -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
  echo Setup failed. Install Python 3 from python.org, then run this file again.
  pause
  exit /b 1
)
echo Setup complete. Double-click Start_Maverick_Roof_Tool.bat.
pause
