@echo off
chcp 65001 >nul
title Go Cai Dat AVTool Pro

taskkill /f /im excel.exe >nul 2>&1
if exist "%APPDATA%\Microsoft\Excel\XLSTART\AVTool_FontViet.xlam" (
    del /f /q "%APPDATA%\Microsoft\Excel\XLSTART\AVTool_FontViet.xlam"
)
echo ================================================================
echo       [OK] Da go cai dat AVTool thanh cong khoi Excel!
echo ================================================================
pause
