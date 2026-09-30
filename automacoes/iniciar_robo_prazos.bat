@echo off
REM Copie este arquivo para a MESMA pasta do robo_prazos.py (ex.: ...\Planilhas).
REM %~dp0 = pasta deste .bat: dispensa o caminho fixo com acento e "&".
REM Agendador de Tarefas: disparar "Ao fazer logon" com atraso de 5 min (Drive precisa montar o G:)
REM e marcar "Executar tarefa assim que possivel apos perder um inicio agendado".
cd /d "%~dp0"
if not exist logs mkdir logs
set LOG=logs\robo_prazos.log

REM Espera ate 5 min o Google Drive disponibilizar o script.
set /a TENTATIVAS=0
:espera
if exist robo_prazos.py goto roda
set /a TENTATIVAS+=1
if %TENTATIVAS% GEQ 10 (echo [%date% %time%] robo_prazos.py indisponivel; Drive montado?>>"%LOG%" & exit /b 2)
timeout /t 30 /nobreak >nul
goto espera

:roda
echo ==== [%date% %time%] inicio ====>>"%LOG%"
python robo_prazos.py >>"%LOG%" 2>&1
set RC=%ERRORLEVEL%
echo ==== [%date% %time%] fim, codigo %RC% ====>>"%LOG%"
exit /b %RC%
