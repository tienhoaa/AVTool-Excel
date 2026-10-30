@echo off
chcp 65001 >nul
title Tu Dong Cap Nhat AVTool Pro

echo ================================================================
echo        DANG TU DONG CAP NHAT AVTOOL PRO TU GITHUB...
echo ================================================================
echo.

:: Cho 1 giay de Excel dong hoan toan
timeout /t 1 /nobreak >nul

set "TARGET_DIR=%APPDATA%\Microsoft\AddIns"
set "TARGET_FILE=%TARGET_DIR%\TN_FontViet_Pro_2026.xlam"
set "DOWNLOAD_URL=https://raw.githubusercontent.com/tienhoaa/AVTool-Excel/main/TN_FontViet_Pro_2026.xlam"

if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

echo Dang tai ban moi nhat ve may...
curl.exe -L -s -o "%TARGET_FILE%" "%DOWNLOAD_URL%"

if exist "%TARGET_FILE%" (
    echo.
    echo ================================================================
    echo        [OK] DA CAP NHAT AVTOOL PRO THANH CONG!
    echo ================================================================
    echo Dang khoi dong lai Excel...
    timeout /t 1 /nobreak >nul
    start excel.exe
) else (
    echo.
    echo [!] Khong the tai ban cap nhat. Vui long kiem tra ket noi mang!
    pause
)
