@echo off
title Servidor de Cartera - ProFrank
cd /d "%~dp0"
echo ========================================================
echo   Iniciando Servidor Web de Reestructuracion de Cartera
echo ========================================================
echo.
echo Abriendo aplicacion en http://127.0.0.1:8000
echo Para detener el servidor, cierra esta ventana.
echo.
python app.py
pause
