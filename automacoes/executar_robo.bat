@echo off
REM Agendador de Tarefas do Windows: acao "Iniciar programa" apontando para este arquivo.
REM Marque "Executar tarefa assim que possivel apos perder um inicio agendado".
REM Usa o Python do .venv diretamente (dispensa activate) e grava log datado.
cd /d "%~dp0.."
if not exist logs mkdir logs
set LOG=logs\robo_%date:~-4%%date:~3,2%%date:~0,2%.log
.venv\Scripts\python.exe -m automacoes.diagnostico_djen --oab 29573 --uf PB >> "%LOG%" 2>&1
REM Quando o notebook do Robo Diario for portado para script local, troque a linha acima por:
REM .venv\Scripts\python.exe -m automacoes.robo_djen >> "%LOG%" 2>&1
