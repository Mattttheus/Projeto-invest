@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Verificando dependencias...
python -m pip install -q -r requirements.txt
echo.
python gestao.py
echo.
pause
