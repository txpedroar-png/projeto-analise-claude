# Automações jurídicas (Python, só biblioteca padrão)

Ferramentas determinísticas para tirar da IA o que não exige leitura: datas, prazos e validação de números CNJ.
Toda saída é **minuta sujeita à revisão de advogado** (manual, seção 2).

| Módulo | O que faz | Comando |
|---|---|---|
| `prazos.py` | Prazo em dias úteis com a regra conservadora do escritório, recesso do art. 220 e feriados do CSV. Devolve vencimento interno **e** termo legal. | `python -m automacoes.prazos 2026-10-09 15 --comarca 0231` |
| `cnj.py` | Valida o dígito (Res. CNJ 65/2008) e a estrutura; audita minutas `.docx`/`.txt`; sai com código 1 se houver número inválido. | `python -m automacoes.cnj minuta.docx` |
| `triagem.py` | Pré-triagem de publicações: resolve em Python as que têm prazo expresso e isola em `--para-ia` só as que exigem leitura. | `python -m automacoes.triagem triagem.csv -o prazos.csv --para-ia revisar.txt` |
| `diagnostico_djen.py` | Uma consulta à API do DJEN e registro em log, para diagnosticar o HTTP 403 do Robô Diário. | `python -m automacoes.diagnostico_djen --oab 29573 --uf PB` |
| `executar_robo.bat` | Execução pelo Agendador de Tarefas do Windows, com log datado em `logs/`. | — |

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
