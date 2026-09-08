@echo off
setlocal
cd /d "%~dp0"
echo ================================================================
echo  Configuracao unica: publicar o painel online automaticamente
echo ================================================================
echo.

rem Remove travas antigas do Git, se existirem
if exist ".git\index.lock" del /f /q ".git\index.lock"
if exist ".git\maintenance.lock" del /f /q ".git\maintenance.lock"
if exist ".git\HEAD.lock" del /f /q ".git\HEAD.lock"

where git >nul 2>&1
if errorlevel 1 (
  echo [ERRO] Git nao encontrado no PATH.
  echo Instale o Git para Windows ^(https://git-scm.com^) ou rode pelo GitHub Desktop.
  pause
  exit /b 1
)

echo Alinhando a pasta local com o GitHub...
git fetch origin
git reset --soft origin/main
git add -A

git diff --cached --quiet
if not errorlevel 1 (
  echo Nada novo para enviar. A pasta ja esta alinhada.
  goto fim
)

git -c user.name="Jaqueline Disktrans" -c user.email="jaqueline@disktrans.com.br" commit -m "Front novo do painel + automacao de publicacao"

echo Enviando para o GitHub ^(pode pedir login uma vez^)...
git push origin main
if errorlevel 1 (
  echo.
  echo [ATENCAO] O envio falhou. Provavelmente falta autenticar no GitHub.
  echo Faca login quando o Git/GitHub pedir e rode este arquivo novamente.
  pause
  exit /b 0
)

:fim
echo.
echo Pronto! A partir de agora a atualizacao automatica tambem publica o painel online.
pause
exit /b 0
