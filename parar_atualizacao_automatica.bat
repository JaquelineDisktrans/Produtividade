@echo off
setlocal
set "NOME_TAREFA=Controle CS - Atualizacao Outlook"
echo Removendo a tarefa "%NOME_TAREFA%"...
schtasks /Delete /TN "%NOME_TAREFA%" /F
echo.
echo A atualizacao automatica foi desativada.
pause
