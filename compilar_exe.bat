@echo off
title Gerando GeradorLaudos.exe
echo ============================================
echo  Gerando o aplicativo (GeradorLaudos.exe)
echo ============================================
echo.
echo [1/2] Instalando dependencias...
python -m pip install --upgrade python-docx pyinstaller
if errorlevel 1 goto erro
echo.
echo [2/2] Criando o executavel (pode levar alguns minutos)...
python -m PyInstaller --noconfirm --onefile --windowed --collect-data docx --name GeradorLaudos laudos_app.py
if errorlevel 1 goto erro
echo.
echo Pronto! O aplicativo esta em:  dist\GeradorLaudos.exe
echo Copie esse arquivo para onde quiser e de dois cliques para abrir.
pause
exit /b 0
:erro
echo.
echo Ocorreu um erro. Confira se o Python esta instalado (marque 'Add Python to PATH').
pause
exit /b 1
