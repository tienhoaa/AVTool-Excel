@echo off
chcp 65001 >nul
title Cap Nhat AVTool Pro Online - GitHub

echo ================================================================
echo          DANG KIEM TRA VA TAI BAN CAP NHAT MOI NHAT...
echo ================================================================
echo.

:: 1. Tat Excel neu dang mo de tranh bi khoa file
echo Dang dong Excel de cap nhat file...
taskkill /f /im excel.exe >nul 2>&1
timeout /t 1 /nobreak >nul

:: 2. Duong dan dich
set "TARGET_DIR=%APPDATA%\Microsoft\AddIns"
set "TARGET_FILE=%TARGET_DIR%\TN_FontViet_Pro_2026.xlam"
set "DOWNLOAD_URL=https://raw.githubusercontent.com/tienhoaa/AVTool-Excel/main/TN_FontViet_Pro_2026.xlam"

if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: 3. Tai truc tiep tu GitHub
echo Dang tai ban moi nhat tu GitHub tienhoaa/AVTool-Excel...
where curl >nul 2>&1
if %errorlevel% equ 0 (
    curl -L -s -o "%TARGET_FILE%" "%DOWNLOAD_URL%"
) else (
    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%DOWNLOAD_URL%', '%TARGET_FILE%')"
)
powershell -Command "Unblock-File -Path '%TARGET_FILE%' -ErrorAction SilentlyContinue"

if exist "%TARGET_FILE%" (
    echo.
    echo ================================================================
    echo        [OK] DA CAP NHAT AVTOOL PRO THANH CONG!
    echo ================================================================
    echo Dang khoi dong lai Excel...
    start excel.exe
) else (
    echo.
    echo [!] Khong the tai ban cap nhat. Vui long kiem tra ket noi mang!
    pause
)
