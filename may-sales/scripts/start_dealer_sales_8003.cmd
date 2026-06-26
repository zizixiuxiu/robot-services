@echo off
chcp 65001 >nul
title dealer-sales-8003

set "PYTHON=C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe"
set "SCRIPT=D:\Services\robot-services\may-sales\src\dealer_sales_http.py"
set "PORT=8003"

echo [INFO] Stopping existing dealer-sales service processes...
for /f "skip=1 tokens=2 delims=," %%P in ('wmic process where "name='python.exe' and commandline like '%%dealer_sales_http.py%%'" get ProcessId /format:csv 2^>nul') do (
    if not "%%P"=="" taskkill /F /PID %%P >nul 2>nul
)

timeout /t 2 /nobreak >nul

echo [INFO] Starting dealer-sales HTTP service on port %PORT%...
start /B "" "%PYTHON%" "%SCRIPT%" --port %PORT%

timeout /t 3 /nobreak >nul

echo [INFO] Checking port %PORT%...
netstat -an | findstr ":%PORT%" | findstr "LISTENING"
if %errorlevel%==0 (
    echo [OK] port %PORT% is listening
) else (
    echo [FAIL] port %PORT% is not listening
)

pause
