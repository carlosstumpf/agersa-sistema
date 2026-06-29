@echo off
title AGERSA - Subir para GitHub
echo ================================================
echo   AGERSA - Configurando GitHub
echo ================================================
echo.

cd /d "C:\Users\Carlos Stumpf\Desktop\AGERSA"

echo Inicializando Git...
git init
git branch -M main

echo.
echo Adicionando arquivos...
git add .
git commit -m "Sistema AGERSA - versao inicial"

echo.
echo Conectando ao GitHub...
git remote add origin https://github.com/carlosstumpf/agersa-sistema.git
git push -u origin main

echo.
echo ================================================
echo   Codigo enviado para o GitHub!
echo   Acesse: https://github.com/carlosstumpf/agersa-sistema
echo ================================================
pause
