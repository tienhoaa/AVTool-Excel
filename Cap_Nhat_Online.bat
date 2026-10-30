@echo off
chcp 65001 >nul
title Cap Nhat AVTool Pro Online - GitHub

echo ================================================================
echo          DANG KIEM TRA VA TAI BAN CAP NHAT MOI NHAT...
echo ================================================================
echo.

:: 1. Tat Excel neu dang mo
taskkill /f /im excel.exe >nul 2>&1

:: 2. Duong dan dich
set "TARGET_DIR=%APPDATA%\Microsoft\AddIns"
set "TARGET_FILE=%TARGET_DIR%\TN_FontViet_Pro_2026.xlam"
set "DOWNLOAD_URL=https://raw.githubusercontent.com/tienhoaa/AVTool-Excel/main/TN_FontViet_Pro_2026.xlam"

if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: 3. Tai truc tiep tu GitHub
echo Dang tai ban moi nhat tu GitHub tienhoaa/AVTool-Excel...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%DOWNLOAD_URL%', '%TARGET_FILE%'); Unblock-File -Path '%TARGET_FILE%'"

if exist "%TARGET_FILE%" (
    echo.
    echo ================================================================
    echo        [OK] DA CAP NHAT AVTOOL PRO THANH CONG!
    echo ================================================================
    echo Ban co the mo Excel len de su dung ngay phien ban moi nhat.
) else (
    echo.
    echo [!] Khong the tai ban cap nhat. Vui long kiem tra ket noi mang!
)

echo.
pause
