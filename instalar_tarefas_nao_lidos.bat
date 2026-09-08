@echo off
setlocal
cd /d "%~dp0"
echo Criando as tarefas agendadas de nao lidos (18h e 23:59)...

schtasks /Create /TN "Disktrans - Nao lidos 18h" /TR "\"%~dp0nao_lidos_18h.bat\"" /SC DAILY /ST 18:00 /F
schtasks /Create /TN "Disktrans - Nao lidos 2359" /TR "\"%~dp0nao_lidos_2359.bat\"" /SC DAILY /ST 23:59 /F

echo.
echo Tarefas criadas. Elas rodam todo dia nesses horarios (com voce logado e o Outlook aberto).
echo Para remover depois: use parar_tarefas_nao_lidos.bat
pause
