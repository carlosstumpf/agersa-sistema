@echo off
title AGERSA - Agendador Automatico
echo ================================================
echo   AGERSA - Agendador Automatico
echo ================================================
echo.
echo Iniciando agendador...
echo Para parar: feche esta janela ou Ctrl+C
echo.
cd /d "%~dp0"
"C:\Users\Carlos Stumpf\AppData\Local\Python\pythoncore-3.14-64\python.exe" agersa_agendador.py
pause
