@echo off
chcp 65001 >nul
setlocal
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
echo ============================================
echo   Check DEM - หาวัน DEM. และ REC. จาก Freedays
echo ============================================
echo.

python --version >nul 2>nul
if errorlevel 1 (
    echo [X] ไม่พบ Python กรุณาติดตั้งจาก https://www.python.org/downloads/
    echo     อย่าลืมติ๊ก "Add python.exe to PATH" ตอนติดตั้ง
    echo.
    pause
    exit /b 1
)

echo กำลังตรวจสอบไลบรารีที่ต้องใช้ ...
python -m pip install -q -r requirements.txt
if errorlevel 1 (
    echo [X] ติดตั้งไลบรารีไม่สำเร็จ ตรวจสอบอินเทอร์เน็ตแล้วลองใหม่
    echo.
    pause
    exit /b 1
)

echo.
set /p ATA=ATA (dd/mm/yyyy):
for %%F in (input\*.xls input\*.xlsx) do (
  if /i not "%%~nF"=="CHECK DEM" python check_dem.py "%%F" %ATA%
)

echo.
pause
