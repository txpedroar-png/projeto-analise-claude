# Instalação do robô de prazos no computador do escritório

Tempo total: cerca de 45 minutos, uma única vez. Precisa de um computador **Windows** que fique ligado nos
horários agendados e use a internet do escritório (o DJEN recusa conexões de fora do Brasil).

## 1. Credencial do robô no Google (≈ 20 min)

1. Acesse <https://console.cloud.google.com> com a conta do escritório e crie um projeto (ex.: `robo-prazos`).
2. Menu **APIs e serviços → Biblioteca** → procure **Google Sheets API** → **Ativar**.
3. **APIs e serviços → Credenciais → Criar credenciais → Conta de serviço**. Nome: `robo-prazos`. Conclua.
4. Abra a conta criada → aba **Chaves** → **Adicionar chave → Criar nova chave → JSON**. Um arquivo é baixado.
5. Crie a pasta `C:\robo` e mova o arquivo para lá com o nome `credencial-robo.json`.
   **Não** coloque esse arquivo no Google Drive, em e-mail ou no chat: ele dá acesso de edição às planilhas
   compartilhadas com o robô.
6. Copie o e-mail da conta de serviço (termina em `iam.gserviceaccount.com`) e, na Planilha Google,
   **Compartilhar** com esse e-mail como **Editor** (desmarque "Notificar").

## 2. Python e o repositório (≈ 15 min)

Os comandos abaixo são para o **PowerShell** (no Prompt de Comando funcionam iguais, exceto onde indicado).

1. Instale o Python 3.11 ou mais novo em <https://www.python.org/downloads/>.
   Na primeira tela do instalador, marque **"Add python.exe to PATH"**.
2. Baixe o código **do ramo do robô** (o botão "Download ZIP" da página inicial baixa o ramo principal, que
   ainda não tem o robô). Com o navegador logado no GitHub, abra:
   <https://github.com/txpedroar-png/projeto-analise-claude/archive/refs/heads/claude/great-bell-rivb0k.zip>
   O arquivo `projeto-analise-claude-claude-great-bell-rivb0k.zip` vai para a pasta Downloads.
3. Extraia e renomeie (se já existir uma pasta `C:\robo\projeto-analise-claude` de tentativa anterior,
   o primeiro comando a renomeia para `projeto-antigo`, sem apagar nada):
   ```
   cd C:\robo
   if (Test-Path projeto-analise-claude) { Rename-Item projeto-analise-claude projeto-antigo }
   Expand-Archive "$env:USERPROFILE\Downloads\projeto-analise-claude-claude-great-bell-rivb0k.zip" -DestinationPath C:\robo
   Rename-Item C:\robo\projeto-analise-claude-claude-great-bell-rivb0k projeto-analise-claude
   cd C:\robo\projeto-analise-claude
   dir requirements.txt, config_local.exemplo.ini
   ```
   O último comando deve listar os dois arquivos. Se disser que não existem, pare e envie a saída de `dir`.
4. Ambiente e configuração:
   ```
   python -m venv .venv
   .venv\Scripts\pip install -r requirements.txt
   copy config_local.exemplo.ini config_local.ini
   notepad config_local.ini
   ```
   No Bloco de Notas, confira `id`, `credencial`, `oab` e `uf`, e salve.

## 3. Ensaio (sem gravar nada)

```
.\automacoes\robos\iniciar_robos.bat ensaio
Get-Content logs\robo_planilha.log
```
(No PowerShell o `.\` no início é obrigatório; no Prompt de Comando, use `type` no lugar de `Get-Content`.)
O log mostra o que o robô **faria**: linhas que inseriria, prazos calculados e processos novos.
Envie o log para conferência antes de rodar de verdade.

- `HTTP 403` no DJEN: a conexão não está saindo do Brasil (VPN? rede de outro lugar?).
- `PermissionError`/`403` do Google: a planilha não foi compartilhada com o e-mail da conta de serviço.

## 4. Execução real e agendamento

1. Rode uma vez: `.\automacoes\robos\iniciar_robos.bat` e confira a planilha (aba Prazos e a aba Auditoria,
   bloco "Registro de execuções do robô").
2. **Agendador de Tarefas → Criar tarefa básica** → diariamente às **07:30** → Ação: Iniciar programa →
   `C:\robo\projeto-analise-claude\automacoes\robos\iniciar_robos.bat`.
   Depois, em Propriedades: aba Disparadores, acrescente **13:00**; aba Configurações, marque
   **"Executar tarefa assim que possível após perder um início agendado"**.

## Regras para a equipe

- O robô **nunca** altera Status, Data do protocolo, Observações, Link, CPF nem linhas encerradas.
- Prazo lançado à mão: preencha Processo, Disponibilização e Prazo (d.u.); o robô calcula os vencimentos.
- Para impedir que o robô recalcule uma linha (ex.: suspensão por portaria), escreva **ajuste manual** em
  Observações da equipe.
- Feriado confirmado: marque **Conferido = S** na aba Feriados (com a fonte); os prazos abertos são
  recalculados na execução seguinte e a mudança fica registrada em Conferência.
