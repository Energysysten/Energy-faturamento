@echo off
REM Agenda o backup do banco no Agendador de Tarefas do Windows (12:30 e 19:00, todo dia).
REM Rodar uma vez, na pasta do sistema (C:\Faturamento), com o usuario que tem acesso ao \\Zion2.
cd /d "%~dp0"

for /f "delims=" %%P in ('where python') do (set "PY=%%P" & goto :achou)
echo Python nao encontrado no PATH.
exit /b 1
:achou

schtasks /Create /F /SC DAILY /ST 12:30 /TN "Faturamento - Backup 12h30" /TR "\"%PY%\" \"%~dp0backup_db.py\""
schtasks /Create /F /SC DAILY /ST 19:00 /TN "Faturamento - Backup 19h" /TR "\"%PY%\" \"%~dp0backup_db.py\""

echo.
echo Rodando um backup de teste agora...
"%PY%" "%~dp0backup_db.py"
echo Codigo de saida: %ERRORLEVEL%  (0 = ok, 2 = falhou em algum destino, 1 = falhou)
pause
