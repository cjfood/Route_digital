@echo off
chcp 65001 >nul
title CJ Foods - Cập Nhật Dữ Liệu Tuyến Master DMS
color 0A

echo ===================================================================
echo     CJ FOODS - CÔNG CỤ CẬP NHẬT DỮ LIỆU TUYẾN MASTER DMS
echo ===================================================================
echo.
echo [1/3] Đang kiểm tra file nguồn trong các thư mục:
echo   - Thư mục 'Shop CJ'
echo   - Thư mục 'Shop Potential'
echo.
echo [2/3] Đang tiến hành đọc và đồng bộ hóa vào Master Database...
echo       (Quá trình xử lý ~270.000 điểm bán mất khoảng 30 - 50 giây)
echo.

REM Đóng các tiến trình Excel chạy ngầm nếu có để tránh khóa file
powershell -NoProfile -Command "Get-Process excel -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue"

REM Chạy script ETL Python
python etl_master.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ===================================================================
    echo    ĐÃ CẬP NHẬT THÀNH CÔNG DỮ LIỆU MASTER DATABASE!
    echo    Bạn chỉ cần tải lại trang web hoặc mở lại 'run_app.bat' để xem.
    echo ===================================================================
) else (
    echo.
    echo [LỖI] Có lỗi xảy ra trong quá trình cập nhật. Vui lòng kiểm tra lại file Excel!
)

echo.
pause
