@echo off
REM Roda os robos em sequencia (nunca juntos: os dois gravam a mesma planilha).
REM Agendador de Tarefas: "Ao fazer logon" com atraso de 5 min e
REM "Executar tarefa assim que possivel apos perder um inicio agendado".
REM Caminho da planilha e OAB ficam em config_local.ini (UTF-8, aceita acentos).
cd /d "%~dp0..\.."
set PYTHONIOENCODING=utf-8
if not exist logs mkdir logs
set LOG=logs\robos.log
echo ==== [%date% %time%] inicio ====>>"%LOG%"
python -m automacoes.robos.resgate_emails >>"%LOG%" 2>&1
echo ---- resgate_emails: codigo %ERRORLEVEL% ---->>"%LOG%"
python -m automacoes.robos.robo_prazos >>"%LOG%" 2>&1
echo ---- robo_prazos: codigo %ERRORLEVEL% (1 = DJEN falhou, DataJud usado) ---->>"%LOG%"
