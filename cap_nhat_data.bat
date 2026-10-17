@echo off
chcp 65001 >nul
title CJ Foods - Cập Nhật Dữ Liệu Tuyến Master DMS
color 0A
cd /d "%~dp0"

echo ===================================================================
echo     CJ FOODS - CÔNG CỤ CẬP NHẬT DỮ LIỆU TUYẾN MASTER DMS
echo ===================================================================
echo.
echo [1/3] Đang kiểm tra file nguồn trong thư mục Data_Raw:
echo   - Thư mục 'Data_Raw\Route' (Kế hoạch tuyến MCP)
echo   - Thư mục 'Data_Raw\Shop_Potential' (Điểm bán tiềm năng)
echo.
echo [2/3] Đang tiến hành đọc và đồng bộ hóa vào Master Database...
echo       (Quá trình xử lý ~270.000 điểm bán mất khoảng 30 - 50 giây)
echo.

REM Đóng các tiến trình Excel chạy ngầm nếu có để tránh khóa file
powershell -NoProfile -Command "Get-Process excel -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue"

REM Chạy script ETL Python
python etl_master.py

if %ERRORLEVEL% EQU 0 goto :SUCCESS
goto :FAILED

:SUCCESS
echo.
echo ===================================================================
echo    ĐÃ CẬP NHẬT THÀNH CÔNG DỮ LIỆU MASTER DATABASE!
echo    File dms_data.js đã được tạo mới hoàn tất.
echo ===================================================================
echo.
echo Bạn có thể mở trực tiếp 'index.html' để kiểm tra bản đồ trên máy tính.
echo.
set /p PUSH_CHOICE=">> Ban co muon day du lieu moi len GitHub Pages (Web) luon khong? (Y/N, mac dinh Y): "
if "%PUSH_CHOICE%"=="" set PUSH_CHOICE=Y
if /i "%PUSH_CHOICE%"=="Y" (
    echo.
    call day_len_github.bat
)
goto :END

:FAILED
echo.
echo [LỖI] Có lỗi xảy ra trong quá trình cập nhật. Vui lòng kiểm tra lại file Excel!
goto :END

:END
echo.
pause
