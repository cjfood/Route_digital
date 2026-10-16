@echo off
chcp 65001 >nul
title ĐẨY BẢN ĐỒ LÊN GITHUB - CJ FOODS ROUTE DIGITAL

echo =======================================================================
echo          TỰ ĐỘNG ĐẨY BẢN ĐỒ LÊN GITHUB (PUBLISH WEB)
echo              Kho: https://github.com/cjfood/Route_digital
echo =======================================================================
echo.

set "GIT_CMD=C:\Users\MinhHien-Sales\AppData\Local\Programs\MinGit\cmd\git.exe"

if not exist "%GIT_CMD%" (
    set "GIT_CMD=git"
)

cd /d "%~dp0"

echo [1/4] Đang kiểm tra các file thay đổi...
"%GIT_CMD%" add .

echo [2/4] Đang đóng gói dữ liệu bản đồ...
set datetime=%date% %time%
"%GIT_CMD%" commit -m "Cap nhat ban do CJ Golden Route Map - %datetime%"

echo [3/4] Đang thiết lập nhánh chính (main)...
"%GIT_CMD%" branch -M main

echo [4/4] Đang tải lên GitHub (cjfood/Route_digital)...
echo.
echo * Luu y: Neu trinh duyet bat len cua so dang nhap GitHub, ban chi can
echo   bam nut "Authorize" hoac "Sign in with your browser" de xac thuc.
echo.

"%GIT_CMD%" push -u origin main

if %errorlevel% neq 0 (
    echo.
    echo =======================================================================
    echo [LOI] Khong the day len GitHub!
    echo Vui long kiem tra lai ket noi Internet hoac quyen truy cap tai khoan.
    echo =======================================================================
) else (
    echo.
    echo =======================================================================
    echo [THANH CONG] Toan bo ban do da duoc day len GitHub thanh cong!
    echo.
    echo Link trang web cua ban (sau 1-2 phut GitHub khoi tao xong):
    echo    👉 https://cjfood.github.io/Route_digital/
    echo =======================================================================
)

echo.
pause
