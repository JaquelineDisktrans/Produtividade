@echo off
setlocal
cd /d "%~dp0"

where git >nul 2>&1
if errorlevel 1 (
  echo [publicar] Git nao encontrado no PATH. Instale o Git ^(ou rode pelo GitHub Desktop^) para publicar automaticamente.
  exit /b 0
)

git add saida\dashboard_data.json

git diff --cached --quiet
if not errorlevel 1 (
  echo [publicar] Sem mudanca no JSON; nada a publicar.
  exit /b 0
)

git -c user.name="Jaqueline Disktrans" -c user.email="jaqueline@disktrans.com.br" commit -m "Atualiza dados do painel (auto)"

git pull --rebase --autostash origin main

git push origin main
if errorlevel 1 (
  echo [publicar] Falha no push. Verifique o login do Git ^(GitHub^) e rode este arquivo manualmente uma vez.
  exit /b 0
)

echo [publicar] Dados publicados no GitHub com sucesso.
exit /b 0
