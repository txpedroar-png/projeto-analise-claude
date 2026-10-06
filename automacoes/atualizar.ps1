# Atualiza o código do robô a partir do GitHub, preservando .venv, config_local.ini e logs.
# Uso (na pasta do projeto):  powershell -ExecutionPolicy Bypass -File automacoes\atualizar.ps1
$ErrorActionPreference = "Stop"
$projeto = Split-Path -Parent $PSScriptRoot
$url = "https://github.com/txpedroar-png/projeto-analise-claude/archive/refs/heads/claude/great-bell-rivb0k.zip"
$temp = Join-Path $env:TEMP "robo-atualizacao"
$zip = "$temp.zip"
Remove-Item -Recurse -Force $temp, $zip -ErrorAction SilentlyContinue
Write-Host "Baixando $url"
Invoke-WebRequest $url -OutFile $zip -UseBasicParsing
Expand-Archive $zip -DestinationPath $temp -Force
$novo = (Get-ChildItem $temp -Recurse -Depth 3 -Filter requirements.txt | Select-Object -First 1).DirectoryName
if (-not $novo) { throw "O arquivo baixado não contém o projeto." }
robocopy $novo $projeto /E /XD .venv logs /XF config_local.ini /NFL /NDL /NJH /NJS | Out-Null
if ($LASTEXITCODE -ge 8) { throw "Falha ao copiar os arquivos (robocopy $LASTEXITCODE)." }
Remove-Item -Recurse -Force $temp, $zip
& (Join-Path $projeto ".venv\Scripts\python.exe") -m pip install -q -r (Join-Path $projeto "requirements.txt")
Write-Host "Atualizado em $projeto"
