@echo off
REM Atualiza a Planilha Google de prazos (DJEN). Agendador de Tarefas: diariamente as 07:30 e 13:00,
REM marcando "Executar tarefa assim que possivel apos perder um inicio agendado".
REM Para testar sem gravar: iniciar_robos.bat ensaio
cd /d "%~dp0..\.."
set PYTHONIOENCODING=utf-8
if not exist logs mkdir logs
set LOG=logs\robo_planilha.log
set MODO=
if /i "%1"=="ensaio" set MODO=--ensaio
echo ==== [%date% %time%] inicio %MODO% ====>>"%LOG%"
.venv\Scripts\python.exe -m automacoes.robos.robo_planilha %MODO% >>"%LOG%" 2>&1
echo ---- codigo %ERRORLEVEL% (0 = ok, 1 = DJEN indisponivel) ---->>"%LOG%"
