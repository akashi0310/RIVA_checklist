@echo off
chcp 65001 > nul
title Cap Nhat Du Lieu Dashboard RIVA Checklist
echo.
echo ===================================================
echo     DANG CAP NHAT DU LIEU DASHBOARD RIVA CHECKLIST
echo ===================================================
echo.
python update_dashboard.py
echo.
echo Nhap pham bat ky de thoat...
pause > nul
