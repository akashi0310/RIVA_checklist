@echo off
chcp 65001 > nul
title Cap Nhat Du Lieu Dashboard An Pham Truyenh Thong - Tu&#7845;n
echo.
echo ===================================================
echo   CAP NHAT DU LIEU CHECKLIST AN PHAM TRUYEN THONG
echo ===================================================
echo.
python update_dashboard_tuan.py
echo.
echo Nhap phim bat ki de thoat...
pause > nul
