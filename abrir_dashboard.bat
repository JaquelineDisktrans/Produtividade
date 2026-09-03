@echo off
setlocal
cd /d "%~dp0"
if not exist "saida\dashboard_data.json" (
  echo Preparando os dados do dashboard...
  py -3.12 gerar_dashboard_data.py
  if errorlevel 1 (
    echo Nao foi possivel preparar os dados. Execute primeiro a analise do Outlook.
    pause
    exit /b 1
  )
)
start "" "http://localhost:8765/index.html"
echo Dashboard aberto em http://localhost:8765/index.html
echo Feche esta janela para encerrar o servidor local.
py -3.12 -m http.server 8765
