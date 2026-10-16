@echo off
chcp 65001 >nul
title CJ Foods - He Thong DMS Route Map
echo =======================================================
echo    CJ FOODS - HE THONG QUAN LY TUYEN BAN HANG ^& MAP DMS
echo =======================================================
echo.
echo Dang kiem tra va khoi dong may chu...
echo Ung dung se tu dong mo trinh duyet tai: http://127.0.0.1:5000
echo.
start "" http://127.0.0.1:5000
python app.py
pause
