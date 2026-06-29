@echo off
title AGERSA - Instalacao
echo ================================================
echo   AGERSA - Instalacao de Dependencias
echo ================================================
echo.
echo Instalando pacotes necessarios...
echo.
"C:\Users\Carlos Stumpf\AppData\Local\Python\pythoncore-3.14-64\python.exe" -m pip install playwright openpyxl pandas schedule streamlit plotly anthropic
echo.
echo Instalando browser do Playwright...
"C:\Users\Carlos Stumpf\AppData\Local\Python\pythoncore-3.14-64\python.exe" -m playwright install chromium
echo.
echo ================================================
echo   Instalacao concluida!
echo   Agora execute: INICIAR_AGENDADOR.bat
echo ================================================
pause
