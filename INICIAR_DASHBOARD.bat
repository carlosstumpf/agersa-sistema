@echo off
title AGERSA - Dashboard
echo ================================================
echo   AGERSA - Dashboard Orcamentario
echo ================================================
echo.
echo Iniciando dashboard...
echo Abrira automaticamente no navegador.
echo Para parar: feche esta janela ou Ctrl+C
echo.
cd /d "%~dp0"
"C:\Users\Carlos Stumpf\AppData\Local\Python\pythoncore-3.14-64\python.exe" -m streamlit run agersa_dashboard.py --server.port 8501 --browser.gatherUsageStats false
pause
