@echo off
setlocal

REM Builds a single portable EXE using PyInstaller.
REM Output will be in the "dist" folder as "CSV Scrambler.exe".

cd /d "%~dp0"

python --version
if errorlevel 1 (
  echo Python is not installed or not on PATH.
  echo Install Python 3, then try again.
  pause
  exit /b 1
)

python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
python -m pip install -r requirements-dev.txt

REM --collect-all msoffcrypto: encrypted .xlsx decryption pulls submodules PyInstaller can miss.
REM openpyxl is covered by PyInstaller's standard hook (avoid --collect-all openpyxl: it drags pandas/numpy).
python -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --name "CSV Scrambler" ^
  --collect-all msoffcrypto ^
  csv_scrambler_app.py

echo.
echo Build complete.
echo EXE: %cd%\dist\CSV Scrambler.exe
pause

