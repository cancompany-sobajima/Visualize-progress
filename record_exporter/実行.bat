@echo off
setlocal

chcp 65001 > nul

for /f "tokens=*" %%i in ('powershell -Command "(Get-Date).AddDays(-1).ToString('yyyy-MM-dd')"') do set "yesterday=%%i"
for /f "tokens=*" %%i in ('powershell -Command "(Get-Date).ToString('yyyy-MM-dd')"') do set "today=%%i"

echo Processing data from %yesterday% to %today%...

python export_production_records.py %yesterday% %today%

if not %errorlevel% == 0 (
    echo.
    echo An error occurred. Please check the messages above.
) else (
    echo.
    echo Successfully exported to production_records.xlsx
)

echo.
pause
