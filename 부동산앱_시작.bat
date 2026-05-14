@echo off
chcp 65001 >nul
title 부동산 실거래가 앱
echo.
echo ====================================================
echo   부동산 실거래가 앱을 시작합니다...
echo   잠시 후 브라우저가 자동으로 열립니다.
echo   종료하려면 이 창에서 Ctrl+C 또는 창 닫기.
echo ====================================================
echo.
cd /d "%~dp0"
python -m streamlit run app.py
pause
