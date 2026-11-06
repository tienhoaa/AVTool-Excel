@echo off
chcp 65001 >nul
title Tu Dong Cap Nhat AVTool Pro

echo ================================================================
echo        DANG TU DONG CAP NHAT AVTOOL PRO TU GITHUB...
echo ================================================================
echo.

set "TARGET_DIR=%APPDATA%\Microsoft\AddIns"
set "TARGET_FILE=%TARGET_DIR%\TN_FontViet_Pro_2026.xlam"
set "TEMP_FILE=%TEMP%\TN_FontViet_New.xlam"
set "DOWNLOAD_URL=https://raw.githubusercontent.com/tienhoaa/AVTool-Excel/main/TN_FontViet_Pro_2026.xlam"

if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

echo [1/2] Dang tai ban moi nhat tu GitHub...
curl.exe -L -s -o "%TEMP_FILE%" "%DOWNLOAD_URL%"

if not exist "%TEMP_FILE%" (
    echo.
    echo [!] Khong the tai ban cap nhat. Vui long kiem tra ket noi Internet!
    pause
    exit /b
)

echo [2/2] Dang cap nhat file Add-in vao he thong...
set /a retry=0

:WAIT_COPY
ping -n 2 127.0.0.1 >nul
copy /y "%TEMP_FILE%" "%TARGET_FILE%" >nul 2>&1
if errorlevel 1 (
    set /a retry+=1
    if %retry% leq 8 goto WAIT_COPY
    echo.
    echo [!] File Add-in dang duoc mo boi Excel.
    echo Vui long dong hoan toan Excel roi bam phim bat ky de tiep tuc!
    pause >nul
    set /a retry=0
    goto WAIT_COPY
)

del /f /q "%TEMP_FILE%" >nul 2>&1

echo.
echo ================================================================
echo        [OK] DA CAP NHAT AVTOOL PRO THANH CONG!
echo ================================================================
echo Dang khoi dong lai Excel...
ping -n 2 127.0.0.1 >nul
start excel.exe
