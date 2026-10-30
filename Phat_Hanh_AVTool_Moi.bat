@echo off
chcp 65001 >nul
title Phat Hanh Phien Ban Moi - AVTool Pro

cd /d "%~dp0"
python release.py

if %errorlevel% neq 0 (
    echo.
    echo [!] Co loi xay ra trong qua trinh phat hanh!
    pause
)
