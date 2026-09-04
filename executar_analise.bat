@echo off
setlocal
cd /d "%~dp0"
echo.
echo Analise de atendimento do Outlook
echo Esta etapa apenas le os e-mails. Nada sera alterado no Outlook.
echo Caixa: trocas@disktrans.com.br. Periodo: todo o historico acessivel.
echo O texto sera salvo localmente em casos_para_analise.jsonl para a analise completa.
echo.
py -3.12 -m pip show pywin32 >nul 2>&1
if errorlevel 1 (
  echo Instalando a dependencia inicial. Isto pode levar alguns minutos...
  py -3.12 -m pip install --user pywin32
  if errorlevel 1 (
    echo.
    echo Nao foi possivel instalar pywin32. Verifique se o Python esta instalado.
    pause
    exit /b 1
  )
)
echo.
py -3.12 outlook_atendimento.py --caixa trocas@disktrans.com.br --limpar-base
if errorlevel 1 (
  echo.
  echo A coleta nao terminou. Nenhum relatorio foi gerado.
  pause
  exit /b 1
)
py -3.12 dividir_para_claude.py
py -3.12 gerar_relatorio.py
py -3.12 gerar_dashboard_data.py
echo.
pause
