@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
set /p ATA=ATA (dd/mm/yyyy):
for %%F in (input\*.xls input\*.xlsx) do (
  if /i not "%%~nF"=="Check DEM" python check_dem.py "%%F" %ATA%
)
pause
