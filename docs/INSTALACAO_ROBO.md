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

1. Instale o Python 3.11 ou mais novo em <https://www.python.org/downloads/>.
   Na primeira tela do instalador, marque **"Add python.exe to PATH"**.
2. Baixe o repositório `projeto-analise-claude` do GitHub (botão **Code → Download ZIP**) e extraia em
   `C:\robo\projeto-analise-claude`. (Fora do Google Drive.)
3. Abra o **Prompt de Comando** e rode:
   ```
   cd C:\robo\projeto-analise-claude
   python -m venv .venv
   .venv\Scripts\pip install -r requirements.txt
   copy config_local.exemplo.ini config_local.ini
   ```
4. Abra `config_local.ini` no Bloco de Notas e confira `id`, `credencial`, `oab` e `uf`.

## 3. Ensaio (sem gravar nada)

```
automacoes\robos\iniciar_robos.bat ensaio
type logs\robo_planilha.log
```
O log mostra o que o robô **faria**: linhas que inseriria, prazos calculados e processos novos.
Envie o log para conferência antes de rodar de verdade.

- `HTTP 403` no DJEN: a conexão não está saindo do Brasil (VPN? rede de outro lugar?).
- `PermissionError`/`403` do Google: a planilha não foi compartilhada com o e-mail da conta de serviço.

## 4. Execução real e agendamento

1. Rode uma vez: `automacoes\robos\iniciar_robos.bat` e confira a planilha (aba Prazos e a aba Auditoria,
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
