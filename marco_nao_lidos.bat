@echo off
setlocal
cd /d "%~dp0"
set MARCO=%1
if "%MARCO%"=="" set MARCO=agora

rem 1) captura o numero de nao lidos no marco informado (precisa do Outlook aberto)
py -3.12 nao_lidos.py --marco %MARCO%

rem 2) regenera o JSON do painel (embute o nao_lidos) e 3) publica no GitHub
py -3.12 gerar_dashboard_data.py
call "publicar_dados.bat"
exit /b 0
