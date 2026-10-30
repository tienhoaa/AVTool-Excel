@echo off
chcp 65001 >nul
title Cai Dat AVTool Pro - Microsoft Excel

echo ================================================================
echo        DANG CAI DAT AVTOOL PRO (DOI FONT & TINH DVALUE)...
echo ================================================================
echo.

:: Kiem tra xem nguoi dung da giai nen ra chua
if not exist "%~dp0TN_FontViet_Pro_2026.xlam" (
    echo [!] CHU Y: BAN CHUA GIAI NEN FILE ZIP!
    echo.
    echo Khong duoc chay truc tiep trong file nen ZIP.
    echo.
    echo HUONG DAN:
    echo 1. Nhap chuot phai vao file AVTool_Pro_2026.zip
    echo 2. Chon "Extract All..." (hoac "Extract to...")
    echo 3. Mo thu muc vua giai nen ra, roi nhap dup vao file Cai_Dat_AVTool.bat
    echo.
    echo ================================================================
    pause
    exit /b
)

:: 1. Tat Excel neu dang mo de tranh khoa file
taskkill /f /im excel.exe >nul 2>&1

:: 2. Xac dinh thu muc AddIns chuan cua Excel
set "TARGET_DIR=%APPDATA%\Microsoft\AddIns"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: 3. Xoa file cu trong XLSTART tranh bi trung 2 Tab
del /f /q "%APPDATA%\Microsoft\Excel\XLSTART\AVTool_FontViet.xlam" >nul 2>&1
del /f /q "%APPDATA%\Microsoft\Excel\XLSTART\TN_FontViet_Pro_2026.xlam" >nul 2>&1

:: 4. Copy file add-in vao thu muc AddIns
copy /y "%~dp0TN_FontViet_Pro_2026.xlam" "%TARGET_DIR%\TN_FontViet_Pro_2026.xlam" >nul

:: 5. Bo chan bao mat cua Windows (Unblock)
powershell -Command "Unblock-File -Path '%TARGET_DIR%\TN_FontViet_Pro_2026.xlam' -ErrorAction SilentlyContinue"

:: 6. Tu dong cap quyen VBA Project de chuc nang Chen DValue chay muot ma
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\15.0\Excel\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\14.0\Excel\Security" /v AccessVBOM /t REG_DWORD /d 1 /f >nul 2>&1

echo ================================================================
echo             [OK] DA CAI DAT THANH CONG 100%!
echo ================================================================
echo.
echo Tu bay gio:
echo  - Moi khi ban mo Excel, tren thanh menu se co ngay tab [AVTool].
echo  - Nut [Doi Font Tieng Viet]: Chuyen font TCVN3, VNI -> Unicode (Khong mat cong thuc).
echo  - Nut [Chen DValue Vao File]: Tinh khoi luong tu chu va tu hien Save As luu file .xls/.xlsm.
echo  - Nut [Cap Nhat Ban Moi]: Kiem tra va cap nhat phien ban tu GitHub.
echo.
echo ================================================================
pause
