# Automações jurídicas (Python)

Ferramentas determinísticas para tirar da IA o que não exige leitura: datas, prazos e validação de números CNJ.
Toda saída é **minuta sujeita à revisão de advogado** (manual, seção 2).

| Módulo | O que faz | Comando |
|---|---|---|
| `prazos.py` | Prazo em dias úteis com a regra conservadora do escritório, recesso do art. 220 e feriados do CSV. Devolve vencimento interno **e** termo legal. | `python -m automacoes.prazos 2026-10-09 15 --comarca 0231` |
| `cnj.py` | Valida o dígito (Res. CNJ 65/2008) e a estrutura; audita minutas `.docx`/`.txt`; sai com código 1 se houver número inválido. | `python -m automacoes.cnj minuta.docx` |
| `triagem.py` | Pré-triagem de publicações: resolve em Python as que têm prazo expresso e isola em `--para-ia` só as que exigem leitura. | `python -m automacoes.triagem triagem.csv -o prazos.csv --para-ia revisar.txt` |
| `diagnostico_djen.py` | Uma consulta à API do DJEN e registro em log, para diagnosticar o HTTP 403. | `python -m automacoes.diagnostico_djen --oab 29573 --uf PB` |

Os quatro módulos acima usam só a biblioteca padrão. Os robôs abaixo precisam de `requests` e `openpyxl` (`pip install -r requirements.txt`).

## Robôs (`automacoes/robos/`)

| Robô | Substitui | Comando |
|---|---|---|
| `robo_planilha.py` | **Robô atual**: DJEN → Planilha Google (modelo v2), sem duplicar, sem tocar nas colunas da equipe | `python -m automacoes.robos.robo_planilha --ensaio` |
| `robo_prazos.py` | Robô para planilha .xlsx local (reserva) | `python -m automacoes.robos.robo_prazos` |
| `calculadora_bacen.py` | Calculadora Bacen | `python -m automacoes.robos.calculadora_bacen Calculo_Valores.xlsx --citacao 28/05/2019 --dobra` |
| `iniciar_robos.bat` | `iniciar_robo.bat` | Agendador de Tarefas: roda o robô da Planilha Google, log em `logs\robo_planilha.log`; `iniciar_robos.bat ensaio` não grava |

### Instalação no computador do escritório (uma vez)

1. Clonar o repositório **fora** da pasta do Google Drive e criar o ambiente: `python -m venv .venv`, `.venv\Scripts\activate`, `pip install -r requirements.txt`.
2. Copiar `config_local.exemplo.ini` para `config_local.ini` e ajustar o caminho da planilha e o e-mail.
3. Apontar o Agendador de Tarefas para `automacoes\robos\iniciar_robos.bat`.

Os robôs mantêm até 10 backups por robô na subpasta `backups\`, ao lado da planilha. O robô de prazos grava ao lado da planilha o `revisar_ia.txt`, só com as publicações que exigem leitura.

## `feriados_forenses.csv`

Colunas: `data,descricao,abrangencia,tipo,conferido,fonte`.

- `abrangencia`: `NACIONAL`, sigla do tribunal (`TJPB`) ou tribunal e comarca (`TJPB:0231`, código de origem = 4 últimos dígitos do CNJ).
- `tipo`: `FERIADO` ou `EXPEDIENTE_REDUZIDO` (este só gera aviso, art. 224, § 1º).
- `conferido`: **só `S` entra na contagem.** Linha `N` nunca adia o vencimento, apenas avisa. Um feriado não confirmado não pode empurrar o prazo para além do termo real.
- Hoje só os feriados nacionais fixados em lei estão como `S`. Carnaval, Semana Santa, Corpus Christi e feriados do TJPB estão como `N` até alguém conferir a portaria do TJPB de 2026 e trocar para `S`, indicando a fonte.
- Para anos sem nenhuma linha no CSV (ex.: 2027), o cálculo avisa que o vencimento pode estar antecipado.

## Instalação e testes

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows  (Linux/Mac: source .venv/bin/activate)
pip install -r requirements-dev.txt
python -m pytest -q
```

## Fora do escopo, por decisão do escritório

Prazo em dobro (arts. 180, 183, 186 e 229 do CPC); ciência ficta por portal (art. 5º, § 3º, da Lei 11.419/2006); prazos em horas ou meses. Quando aparecem no texto, a triagem marca a linha para revisão.

## Planilha de gestão (modelo v2)

`automacoes/planilha/gerar_modelo.py` gera o modelo da planilha e migra os dados da anterior:

```bash
python -m automacoes.planilha.gerar_modelo exportacao_da_planilha_antiga.xlsx -o modelo_v2.xlsx \
    --cadastro exportacao_do_cadastro_de_clientes.xlsx
```

Abas: Painel · Prazos (base diária, por tribunal → comarca → vencimento) · Vistas por tribunal (somente leitura) ·
Processos (com CPF do cliente) · Clientes (Cadastro Central, CPF como chave) · Decisões (jurimetria) · Publicações · Feriados · Comarcas · Auditoria · Leia-me.
As datas de prazo vêm do motor `prazos.py`; as fórmulas cuidam só do que muda com o dia. O arquivo gerado contém
dados de clientes: não versionar.

Passo a passo de instalação no computador do escritório: [`docs/INSTALACAO_ROBO.md`](../docs/INSTALACAO_ROBO.md).
