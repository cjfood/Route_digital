@echo off
title DAY LEN GITHUB - CJ FOODS ROUTE DIGITAL
cd /d "%~dp0"

set "GIT=C:\Users\MinhHien-Sales\AppData\Local\Programs\MinGit\cmd\git.exe"

echo =======================================================================
echo          DANG DONG GOI VA DAY LEN GITHUB
echo          Kho: https://github.com/cjfood/Route_digital
echo =======================================================================
echo.

echo [1/4] Dong goi du lieu...
"%GIT%" add .

echo [2/4] Tao ban ghi commit...
"%GIT%" commit -m "Update-CJ-Golden-Route-Map"

echo [3/4] Chon nhanh main...
"%GIT%" branch -M main

echo [4/4] Dang day du lieu len GitHub...
echo.
echo * Luu y: Neu trinh duyet bat len cua so dang nhap, hay bam Authorize hoac Sign in.
echo.

"%GIT%" push -u origin main

echo.
echo =======================================================================
echo   Hoan tat! Neu bao xong hoac Everything up-to-date,
echo   ban hay vao link duoi day tren trinh duyet (doi 1-2 phut):
echo   Link web: https://cjfood.github.io/Route_digital/
echo =======================================================================
echo.
pause
