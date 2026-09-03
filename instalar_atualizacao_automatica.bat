@echo off
setlocal
cd /d "%~dp0"

set "NOME_TAREFA=Controle CS - Atualizacao Outlook"
set "SCRIPT=%~dp0atualizar_automatico.bat"

echo Criando a atualizacao automatica a cada 30 minutos...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$action = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument ('/c ' + [char]34 + $env:SCRIPT + [char]34); $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration (New-TimeSpan -Days 3650); $settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable; Register-ScheduledTask -TaskName $env:NOME_TAREFA -Action $action -Trigger $trigger -Settings $settings -Description 'Atualiza a caixa de Trocas do Outlook localmente' -Force | Out-Null"
if errorlevel 1 (
  echo.
  echo Nao foi possivel criar a tarefa. Execute este arquivo com o Outlook aberto.
  pause
  exit /b 1
)
echo.
echo Pronto. O Windows tentara atualizar a cada 30 minutos enquanto o usuario estiver conectado.
echo O Outlook classico precisa estar aberto e sincronizado.
echo O log fica em saida\atualizacao.log
pause
