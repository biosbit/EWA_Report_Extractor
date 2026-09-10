@echo off
setlocal
cd /d %~dp0
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name EWA_Action_Extractor ewa_action_extractor.py
if errorlevel 1 (
  echo Build failed.
  pause
  exit /b 1
)
echo.
echo Build completed: dist\EWA_Action_Extractor.exe
echo.
pause
