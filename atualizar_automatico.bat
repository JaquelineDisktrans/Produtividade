@echo off
setlocal
cd /d "%~dp0"
py -3.12 atualizar_automatico.py
exit /b %errorlevel%
