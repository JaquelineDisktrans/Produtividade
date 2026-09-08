@echo off
setlocal
schtasks /Delete /TN "Disktrans - Nao lidos 18h" /F
schtasks /Delete /TN "Disktrans - Nao lidos 2359" /F
echo Tarefas de nao lidos removidas.
pause
