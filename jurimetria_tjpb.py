"""Jurimetria das ações de tarifas bancárias no TJPB via API pública do DataJud (CNJ).

Pesquisa: "A presunção judicial de má-fé e o risco profissional" (OAB/PB, 2026).

Etapas (rode na ordem, cada uma retoma de onde parou):

  python jurimetria_tjpb.py diagnostico      # mostra 2 processos brutos e os campos existentes
  python jurimetria_tjpb.py coletar          # universo TJPB do assunto Tarifas (11807), com movimentos
  python jurimetria_tjpb.py enriquecer --base base_consolidada_bradesco_sem_nomes.csv
                                             # busca no DataJud os processos da base própria
  python jurimetria_tjpb.py analisar         # tabelas, testes e gráficos -> Excel + resumo

Decisões de método (documentadas para o artigo):
  * A API pública do DataJud não traz as partes. Não há como filtrar Bradesco por ela;
    o recorte Bradesco vem da base própria e da leitura dos autos no PJe.
  * O desfecho é lido pelo NOME do movimento processual (Tabela Processual Unificada),
    não por uma lista fixa de códigos. A aba "movimentos_inventario" lista todo par
    código/nome encontrado e a categoria atribuída, para conferência humana.
  * O DataJud guarda um registro por grau (G1, JE, G2, TR...). O desfecho de 1º grau
    vem dos registros G1/JE; o de 2º grau, dos registros G2/TR.
  * Paginação por search_after (o "from" do Elasticsearch para em 10.000 resultados).
"""

import argparse
import gzip
import json
import os
import re
import sys
import time
import unicodedata
from datetime import datetime, timezone

import requests

URL = "https://api-publica.datajud.cnj.jus.br/api_publica_tjpb/_search"
# Chave pública divulgada pelo CNJ na documentação da API (não é segredo pessoal).
API_KEY = os.environ.get(
    "DATAJUD_API_KEY", "cDZHYzlZa0JadVREZDJCendQbXY6SkJlTzNjLV9TRENyQk1RdnFKZGRQdw=="
)
HEADERS = {"Authorization": f"APIKey {API_KEY}", "Content-Type": "application/json"}

ASSUNTO_TARIFAS = 11807
TAMANHO_LOTE = 100
MAX_TENTATIVAS = 6
TIMEOUT = 60

# Marcos temporais (datas conferidas nas fontes do projeto).
MARCOS = [
    ("Rec. CNJ 159/2024", "2024-10-23"),
    ("Tema 1198 STJ", "2025-03-13"),
    ("IRDR TJPB inadmitido", "2026-02-02"),
]
PERIODOS = [
    ("1. Até Rec. 159 (até 22/10/2024)", None, "2024-10-23"),
    ("2. Rec. 159 → Tema 1198", "2024-10-23", "2025-03-13"),
    ("3. Tema 1198 → IRDR TJPB", "2025-03-13", "2026-02-02"),
    ("4. Após IRDR TJPB", "2026-02-02", None),
]

MUNICIPIOS_URL = (
    "https://raw.githubusercontent.com/txpedroar-png/projeto-analise-claude/"
    "claude/tjpb-datajud-api-review-jyag25/Munic%C3%ADpios_BR_Recortes_Territoriais_2022.CSV"
)

DIR = os.environ.get("JURIMETRIA_DIR", "dados_jurimetria")


# ----------------------------------------------------------------------------- rede

def _post(payload):
    espera = 2.0
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            r = requests.post(URL, headers=HEADERS, json=payload, timeout=TIMEOUT)
        except requests.exceptions.RequestException as exc:
            if tentativa == MAX_TENTATIVAS:
                raise
            print(f"  conexão falhou ({exc.__class__.__name__}); nova tentativa em {espera:.0f}s")
            time.sleep(espera)
            espera *= 2
            continue
        if r.status_code == 200:
            return r.json()
        if r.status_code == 429 or r.status_code >= 500:
            if tentativa == MAX_TENTATIVAS:
                r.raise_for_status()
            pausa = float(r.headers.get("Retry-After") or espera)
            print(f"  HTTP {r.status_code}; nova tentativa em {pausa:.0f}s")
            time.sleep(pausa)
            espera *= 2
            continue
        raise requests.exceptions.HTTPError(f"HTTP {r.status_code}: {r.text[:500]}")
    raise RuntimeError("tentativas esgotadas")


# ----------------------------------------------------------------------------- utilidades

def so_digitos(x):
    return re.sub(r"\D", "", str(x or ""))


def formatar_cnj(n):
    """20 dígitos -> NNNNNNN-DD.AAAA.J.TR.OOOO (máscara padrão do CNJ).

    Usar só nas planilhas/CSVs de saída, nunca na coluna interna "numero"
    usada para casar processos: uma coluna só com dígitos, aberta no Excel,
    é lida como número e perde os zeros à esquerda -- corrompendo o número
    do processo exatamente onde alguém for usá-lo para localizar o caso.
    """
    n = str(n or "")
    if len(n) != 20 or not n.isdigit():
        return n
    return f"{n[0:7]}-{n[7:9]}.{n[9:13]}.{n[13]}.{n[14:16]}.{n[16:20]}"


def norm(txt):
    t = unicodedata.normalize("NFKD", str(txt or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().upper()


def data_iso(valor):
    """Aceita '2023-02-01T10:20:00.000Z', '20230201102000', '2023-02-01'. Retorna 'AAAA-MM-DD' ou None."""
    s = str(valor or "").strip()
    if not s:
        return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.match(r"^(\d{4})(\d{2})(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return None


def movimentos_de(source):
    # A API usa "movimentos"; aceita "movimentacoes" por segurança.
    return source.get("movimentos") or source.get("movimentacoes") or []


def abrir_jsonl(caminho, modo):
    return gzip.open(caminho, modo + "t", encoding="utf-8")


def ler_jsonl(caminho):
    if not os.path.exists(caminho):
        return
    with abrir_jsonl(caminho, "r") as f:
        for linha in f:
            linha = linha.strip()
            if linha:
                yield json.loads(linha)


# ----------------------------------------------------------------------------- 1. diagnóstico

def cmd_diagnostico(args):
    q = {"query": {"terms": {"assuntos.codigo": args.assuntos}}, "size": 2}
    d = _post(q)
    total = d.get("hits", {}).get("total", {})
    print(f"Total informado pela API para o assunto {args.assuntos}: {total}")
    for h in d.get("hits", {}).get("hits", []):
        s = h.get("_source", {})
        print("\nCampos de primeiro nível:", sorted(s.keys()))
        print("Tem campo de partes?", any(k.lower().startswith("parte") or "polo" in k.lower() for k in s))
        movs = movimentos_de(s)
        print(f"Movimentos: {len(movs)}; exemplo:", json.dumps(movs[:3], ensure_ascii=False)[:800])
        print("orgaoJulgador:", s.get("orgaoJulgador"))
        print("grau:", s.get("grau"), "| dataAjuizamento:", s.get("dataAjuizamento"))


# ----------------------------------------------------------------------------- 2. coleta do universo

def cmd_coletar(args):
    os.makedirs(DIR, exist_ok=True)
    bruto = os.path.join(DIR, "universo_tarifas.jsonl.gz")
    ckpt = os.path.join(DIR, "universo_checkpoint.json")
    estado = {"search_after": None, "lidos": 0, "sort_desempate": True, "concluido": False}
    if os.path.exists(ckpt):
        estado.update(json.load(open(ckpt)))
        if estado["concluido"] and not args.refazer:
            print(f"Coleta já concluída ({estado['lidos']} registros). Use --refazer para recomeçar.")
            return
        if args.refazer:
            estado = {"search_after": None, "lidos": 0, "sort_desempate": True, "concluido": False}
            if os.path.exists(bruto):
                os.remove(bruto)
        else:
            print(f"Retomando: {estado['lidos']} registros já gravados.")

    query = {"terms": {"assuntos.codigo": args.assuntos}}
    total = None
    with abrir_jsonl(bruto, "a") as saida:
        while True:
            sort = [{"@timestamp": "asc"}]
            if estado["sort_desempate"]:
                sort.append({"numeroProcesso": "asc"})
            payload = {"query": query, "size": TAMANHO_LOTE, "sort": sort, "track_total_hits": True}
            if estado["search_after"] is not None:
                payload["search_after"] = estado["search_after"]
            try:
                d = _post(payload)
            except requests.exceptions.HTTPError as exc:
                if estado["sort_desempate"] and "fielddata" in str(exc).lower():
                    print("numeroProcesso não é ordenável; seguindo só com @timestamp.")
                    estado["sort_desempate"] = False
                    continue
                raise
            hits = d.get("hits", {}).get("hits", [])
            if total is None:
                total = d.get("hits", {}).get("total", {}).get("value")
                print(f"Registros no DataJud (todos os graus) para {args.assuntos}: {total}")
            if not hits:
                break
            for h in hits:
                saida.write(json.dumps(h.get("_source", {}), ensure_ascii=False) + "\n")
            saida.flush()
            estado["lidos"] += len(hits)
            estado["search_after"] = hits[-1].get("sort")
            json.dump(estado, open(ckpt, "w"))
            print(f"  {estado['lidos']} / {total}")
            if len(hits) < TAMANHO_LOTE:
                break
            time.sleep(0.4)
    estado["concluido"] = True
    json.dump(estado, open(ckpt, "w"))
    print(f"Coleta concluída: {estado['lidos']} registros em {bruto}")


# ----------------------------------------------------------------------------- 3. enriquecimento da base própria

def cmd_enriquecer(args):
    import pandas as pd

    os.makedirs(DIR, exist_ok=True)
    saida_path = os.path.join(DIR, "base_propria_datajud.jsonl.gz")
    base = pd.read_csv(args.base, dtype=str)
    col = next((c for c in base.columns if norm(c) in ("CNJ", "NUMERO CNJ", "NUMERO_CNJ", "NUMEROPROCESSO")), None)
    if col is None:
        sys.exit(f"Não achei a coluna do número CNJ em {args.base}. Colunas: {list(base.columns)}")
    # Guarda o texto original (com pontuação) de cada número, para o caso de
    # o índice armazenar numeroProcesso pontuado em vez de só dígitos --
    # não há como confirmar isso sem consultar a API a partir deste ambiente.
    original_por_digitos = {}
    for x in base[col]:
        d = so_digitos(x)
        if len(d) == 20:
            original_por_digitos.setdefault(d, str(x).strip())
    numeros = sorted(original_por_digitos)
    ja = {so_digitos(s.get("numeroProcesso")) for s in ler_jsonl(saida_path)}
    pendentes = [n for n in numeros if n not in ja]
    print(f"Base própria: {len(numeros)} números; já consultados {len(numeros) - len(pendentes)}; faltam {len(pendentes)}.")
    achados = set(ja)
    usar_formato_original = False
    testou_formato = False
    with abrir_jsonl(saida_path, "a") as saida:
        for i in range(0, len(pendentes), 50):
            lote_digitos = pendentes[i:i + 50]
            lote = [original_por_digitos[n] for n in lote_digitos] if usar_formato_original else lote_digitos
            d = _post({"query": {"terms": {"numeroProcesso": lote}}, "size": 500})
            hits = d.get("hits", {}).get("hits", [])
            if i == 0 and not hits and not testou_formato:
                lote_alt = [original_por_digitos[n] for n in lote_digitos]
                d_alt = _post({"query": {"terms": {"numeroProcesso": lote_alt}}, "size": 500})
                if d_alt.get("hits", {}).get("hits"):
                    print("Aviso: numeroProcesso parece estar indexado com pontuação;"
                          " passo a consultar no formato original do CNJ.")
                    usar_formato_original = True
                    d, hits = d_alt, d_alt.get("hits", {}).get("hits", [])
                testou_formato = True
            for h in hits:
                s = h.get("_source", {})
                saida.write(json.dumps(s, ensure_ascii=False) + "\n")
                achados.add(so_digitos(s.get("numeroProcesso")))
            saida.flush()
            print(f"  {min(i + 50, len(pendentes))} / {len(pendentes)} consultados")
            time.sleep(0.4)
    faltando = sorted(set(numeros) - achados)
    with open(os.path.join(DIR, "base_propria_nao_encontrados.txt"), "w") as f:
        f.write("\n".join(faltando))
    print(f"Encontrados no DataJud: {len(set(numeros) & achados)}; não encontrados: {len(faltando)} "
          f"(lista em {DIR}/base_propria_nao_encontrados.txt)")


# ----------------------------------------------------------------------------- 4. classificação de movimentos

REGRAS_1G = [
    ("Parcialmente procedente", r"^PROCEDENCIA EM PARTE|^PROCEDENTE EM PARTE|PARCIAL PROCEDENCIA"),
    ("Procedente", r"^PROCEDENCIA\b|^PROCEDENTE\b"),
    ("Improcedente", r"^IMPROCEDENCIA|^IMPROCEDENTE"),
    ("Prescrição/decadência", r"PRESCRICAO|DECADENCIA"),
    ("Acordo homologado", r"HOMOLOGACAO DE TRANSACAO|HOMOLOGACAO DE ACORDO|TRANSACAO HOMOLOGADA"),
    ("Sem mérito: indeferimento da inicial", r"INDEFERIMENTO DA PETICAO INICIAL|INDEFERIDA A PETICAO INICIAL"),
    ("Sem mérito: abandono", r"ABANDONO"),
    ("Sem mérito: desistência", r"DESISTENCIA"),
    # ILEGITIMIDADE/INEPCIA também aparecem em movimentos que excluem só uma
    # parte ou emendam a inicial no meio do processo, sem encerrá-lo. Não dá
    # para distinguir isso do nome do movimento sozinho -- confira sempre as
    # linhas com esses termos na aba movimentos_inventario antes de confiar
    # nos números de "Extinção sem mérito".
    ("Sem mérito: condições/pressupostos", r"AUSENCIA DAS CONDICOES|AUSENCIA DE PRESSUPOSTOS|FALTA DE INTERESSE|ILEGITIMIDADE|INEPCIA"),
    # CONEXAO foi removido de propósito: reunião de processos por conexão não
    # encerra o caso (ele segue tramitando, junto com o outro), diferente de
    # litispendência/coisa julgada/perempção. Incluir CONEXAO aqui inflaria
    # artificialmente a taxa de "extinção sem mérito" -- exatamente a métrica
    # central da hipótese da pesquisa.
    ("Sem mérito: litispendência/coisa julgada/perempção", r"LITISPENDENCIA|COISA JULGADA|PEREMPCAO"),
    ("Sem mérito: outra extinção", r"^EXTINCAO|EXTINTO O PROCESSO|SEM RESOLUCAO DO MERITO"),
]
REGRAS_2G = [
    ("Provido em parte", r"^PROVIMENTO EM PARTE|PROVIDO EM PARTE"),
    ("Não provido", r"^NAO[- ]?PROVIMENTO|^NAO PROVIDO|DESPROVIMENTO"),
    ("Provido", r"^PROVIMENTO\b|^PROVIDO\b"),
    ("Não conhecido", r"^NAO[- ]?CONHECIMENTO|NAO CONHECIDO"),
    ("Anulação/cassação", r"ANULACAO|CASSACAO"),
]
MARCADORES = [
    ("transito_julgado", r"TRANSITO EM JULGADO"),
    ("arquivamento_definitivo", r"ARQUIVAMENTO DEFINITIVO|ARQUIVADO DEFINITIVAMENTE|BAIXA DEFINITIVA"),
    ("alvara_levantamento", r"ALVARA|LEVANTAMENTO"),
    ("remessa_2g", r"REMETIDOS OS AUTOS.*INSTANCIA SUPERIOR|REMESSA.*INSTANCIA SUPERIOR"),
    ("redistribuicao", r"REDISTRIBU"),
    ("determinacao_emenda", r"EMENDA"),
]
_R1 = [(c, re.compile(p)) for c, p in REGRAS_1G]
_R2 = [(c, re.compile(p)) for c, p in REGRAS_2G]
_RM = [(c, re.compile(p)) for c, p in MARCADORES]


def classificar_nome(nome, regras):
    n = norm(nome)
    for cat, rx in regras:
        if rx.search(n):
            return cat
    return None


def grupo(cat):
    if cat is None:
        return None
    if cat.startswith("Sem mérito"):
        return "Extinção sem mérito"
    if cat == "Prescrição/decadência":
        return "Improcedente"
    return cat


def eh_2g(grau):
    # Estes são os valores esperados do campo "grau" (G1/JE = 1º grau;
    # G2/TR/SUP/STJ/TRU/TNU = 2º grau ou superior), mas não há como
    # confirmar isso sem consultar a API a partir deste ambiente. Confira
    # o "grau:" que aparece na saída de `diagnostico` contra esta lista
    # antes de confiar na separação 1º grau x 2º grau desta análise.
    return str(grau or "").upper() in ("G2", "TR", "SUP", "STJ", "TRU", "TNU")


def resumir_registro(s):
    """Uma linha por registro (processo x grau)."""
    movs = sorted(movimentos_de(s), key=lambda m: str(m.get("dataHora") or ""))
    grau = s.get("grau")
    regras = _R2 if eh_2g(grau) else _R1
    decisoes = []
    marcas = {k: False for k, _ in MARCADORES}
    for m in movs:
        nome = m.get("nome") or ""
        cat = None
        n = norm(nome)
        for c, rx in regras:
            if rx.search(n):
                cat = c
                break
        if cat:
            decisoes.append((data_iso(m.get("dataHora")), cat, nome, m.get("codigo")))
        for k, rx in _RM:
            if rx.search(n):
                marcas[k] = True
    oj = s.get("orgaoJulgador") or {}
    classe = s.get("classe") or {}
    ass = s.get("assuntos") or []
    ass_flat = []
    for a in ass:
        if isinstance(a, list):
            ass_flat.extend(a)
        else:
            ass_flat.append(a)
    primeira = decisoes[0] if decisoes else (None, None, None, None)
    ultima = decisoes[-1] if decisoes else (None, None, None, None)
    return {
        "numero": so_digitos(s.get("numeroProcesso")),
        "grau": grau,
        "classe": classe.get("nome"),
        "classe_codigo": classe.get("codigo"),
        "orgao": oj.get("nome"),
        "orgao_codigo": oj.get("codigo"),
        "municipio_ibge": oj.get("codigoMunicipioIBGE"),
        "origem_cnj": so_digitos(s.get("numeroProcesso"))[-4:],
        "data_ajuizamento": data_iso(s.get("dataAjuizamento")),
        "assuntos": "; ".join(f"{a.get('codigo')}-{a.get('nome')}" for a in ass_flat if isinstance(a, dict)),
        "n_movimentos": len(movs),
        "n_decisoes": len(decisoes),
        "decisao_primeira": primeira[1],
        "data_primeira_decisao": primeira[0],
        "decisao_ultima": ultima[1],
        "data_ultima_decisao": ultima[0],
        "grupo_ultima": grupo(ultima[1]) if not eh_2g(grau) else ultima[1],
        "decisoes_todas": " | ".join(f"{d[0]}:{d[1]}" for d in decisoes),
        **marcas,
    }


# ----------------------------------------------------------------------------- 5. análise

def periodo_de(data):
    if not data:
        return None
    for rot, ini, fim in PERIODOS:
        if (ini is None or data >= ini) and (fim is None or data < fim):
            return rot
    return None


def carregar_municipios():
    import pandas as pd
    for fonte in (os.path.join(DIR, "municipios.csv"), MUNICIPIOS_URL):
        try:
            m = pd.read_csv(fonte, sep=";", encoding="latin1", dtype=str)
            if fonte != os.path.join(DIR, "municipios.csv"):
                m.to_csv(os.path.join(DIR, "municipios.csv"), sep=";", encoding="latin1", index=False)
            return dict(zip(m["CD_MUN"], m["NM_MUN"]))
        except Exception:
            continue
    print("Aviso: não consegui carregar a tabela de municípios; uso o nome do órgão.")
    return {}


def wilson(k, n, z=1.96):
    if not n:
        return (None, None)
    p = k / n
    den = 1 + z * z / n
    c = (p + z * z / (2 * n)) / den
    h = z * ((p * (1 - p) / n + z * z / (4 * n * n)) ** 0.5) / den
    return (round(100 * (c - h), 1), round(100 * (c + h), 1))


def cmd_analisar(args):
    import pandas as pd
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    try:
        from scipy.stats import chi2_contingency
    except ImportError:
        chi2_contingency = None

    univ_path = os.path.join(DIR, "universo_tarifas.jsonl.gz")
    prop_path = os.path.join(DIR, "base_propria_datajud.jsonl.gz")
    registros, inventario = [], {}
    fontes = [("universo", univ_path), ("base_propria", prop_path)]
    vistos = set()
    for rotulo, caminho in fontes:
        for s in ler_jsonl(caminho):
            chave = (so_digitos(s.get("numeroProcesso")), s.get("grau"), (s.get("orgaoJulgador") or {}).get("codigo"))
            regras = _R2 if eh_2g(s.get("grau")) else _R1
            for m in movimentos_de(s):
                k = (m.get("codigo"), m.get("nome"))
                if k not in inventario:
                    inventario[k] = [0, classificar_nome(m.get("nome"), regras) or ""]
                inventario[k][0] += 1
            if chave in vistos:
                continue
            vistos.add(chave)
            r = resumir_registro(s)
            r["fonte"] = rotulo
            registros.append(r)
    if not registros:
        sys.exit("Nada para analisar: rode 'coletar' e/ou 'enriquecer' antes.")

    df = pd.DataFrame(registros)
    muni = carregar_municipios()
    df["municipio"] = df["municipio_ibge"].astype(str).map(muni).fillna(df["orgao"])
    df["tipo_unidade"] = df["orgao"].map(lambda o: "Juizado" if "JUIZADO" in norm(o) else ("2º grau/Turma" if any(t in norm(o) for t in ("CAMARA", "TURMA", "GABINETE", "SECAO", "DESEMBARG")) else "Vara"))
    df["ano_ajuizamento"] = df["data_ajuizamento"].str[:4]

    proprios = set()
    if args.base and os.path.exists(args.base):
        b = pd.read_csv(args.base, dtype=str)
        colb = next(c for c in b.columns if norm(c) in ("CNJ", "NUMERO CNJ", "NUMERO_CNJ", "NUMEROPROCESSO"))
        proprios = {so_digitos(x) for x in b[colb]}
        if "Status_Resultado" in b.columns:
            st = dict(zip(b[colb].map(so_digitos), b["Status_Resultado"]))
            df["status_base_propria"] = df["numero"].map(st)
    df["base_propria"] = df["numero"].isin(proprios)

    g1 = df[~df["grau"].map(eh_2g)].copy()
    g2 = df[df["grau"].map(eh_2g)].copy()
    g1d = g1[g1["grupo_ultima"].notna()].copy()
    g1d["periodo"] = g1d["data_ultima_decisao"].map(periodo_de)
    g1d["semestre"] = g1d["data_ultima_decisao"].map(lambda d: f"{d[:4]}-S{1 if int(d[5:7]) <= 6 else 2}" if d else None)
    g1d["extinto_sem_merito"] = g1d["grupo_ultima"].eq("Extinção sem mérito")
    g1d["exito_autor"] = g1d["grupo_ultima"].isin(["Procedente", "Parcialmente procedente", "Acordo homologado"])

    ordem = ["Procedente", "Parcialmente procedente", "Acordo homologado", "Improcedente", "Extinção sem mérito"]

    def dist(frame, por):
        t = pd.crosstab(frame[por], frame["grupo_ultima"]).reindex(columns=ordem, fill_value=0)
        t["Total"] = t.sum(axis=1)
        pct = t[ordem].div(t["Total"], axis=0).mul(100).round(1)
        pct.columns = [f"% {c}" for c in ordem]
        return pd.concat([t, pct], axis=1)

    tabelas = {}
    tabelas["1_universo"] = pd.DataFrame({
        "Indicador": [
            "Registros DataJud (processo x grau)", "Processos distintos",
            "Registros de 1º grau (G1/JE)", "Registros de 2º grau/Turma",
            "1º grau com decisão identificada", "Processos da base própria encontrados",
        ],
        "Valor": [
            len(df), df["numero"].nunique(), len(g1), len(g2), len(g1d),
            df.loc[df["base_propria"], "numero"].nunique(),
        ],
    })
    tabelas["2_ajuizamento_ano"] = pd.crosstab(g1["ano_ajuizamento"], g1["tipo_unidade"], margins=True, margins_name="Total")
    tabelas["3_resultado_periodo"] = dist(g1d, "periodo")
    tabelas["4_resultado_semestre"] = dist(g1d, "semestre")
    top = g1d["municipio"].value_counts().head(20).index
    tabelas["5_resultado_municipio"] = dist(g1d[g1d["municipio"].isin(top)], "municipio")
    tabelas["6_resultado_tipo_unidade"] = dist(g1d, "tipo_unidade")
    sm = g1d[g1d["extinto_sem_merito"]]
    ext = pd.crosstab(sm["periodo"], sm["decisao_ultima"])
    tabelas["7_tipos_extincao_periodo"] = ext

    linhas = []
    for rot, _, _ in PERIODOS:
        f = g1d[g1d["periodo"] == rot]
        n, k = len(f), int(f["extinto_sem_merito"].sum())
        ke = int(f["exito_autor"].sum())
        lo, hi = wilson(k, n)
        lo2, hi2 = wilson(ke, n)
        linhas.append({"Período": rot, "Decisões": n, "Extinções sem mérito": k,
                       "% extinção": round(100 * k / n, 1) if n else None, "IC95% extinção": f"{lo}–{hi}" if n else "",
                       "Êxito do autor (proc.+parcial+acordo)": ke,
                       "% êxito": round(100 * ke / n, 1) if n else None, "IC95% êxito": f"{lo2}–{hi2}" if n else ""})
    teste = pd.DataFrame(linhas)
    if chi2_contingency is not None:
        ct = pd.crosstab(g1d["periodo"], g1d["extinto_sem_merito"])
        if ct.shape == (len(ct.index), 2) and len(ct.index) > 1:
            chi2, p, dof, _ = chi2_contingency(ct)
            teste.loc[len(teste)] = {"Período": f"Qui-quadrado extinção x período: χ²={chi2:.2f}, gl={dof}, p={p:.4f}"}
    tabelas["8_teste_marcos"] = teste

    if len(g2):
        g2["ano_decisao"] = g2["data_ultima_decisao"].str[:4]
        g2d = g2[g2["decisao_ultima"].notna()]
        tabelas["9_2grau_ano"] = pd.crosstab(g2d["ano_decisao"], g2d["decisao_ultima"], margins=True, margins_name="Total")
        tabelas["10_2grau_orgao"] = pd.crosstab(g2d["orgao"], g2d["decisao_ultima"], margins=True, margins_name="Total")

    g1t = g1d[g1d["data_ajuizamento"].notna() & g1d["data_primeira_decisao"].notna()].copy()
    if len(g1t):
        g1t["dias"] = (pd.to_datetime(g1t["data_primeira_decisao"]) - pd.to_datetime(g1t["data_ajuizamento"])).dt.days
        g1t = g1t[g1t["dias"] >= 0]
        tabelas["11_tempo_1a_decisao"] = g1t.groupby("periodo")["dias"].describe(percentiles=[.25, .5, .75]).round(0)

    if proprios:
        comp = g1d.assign(grupo_base=g1d["base_propria"].map({True: "Base própria (Bradesco, 2 escritórios)", False: "Demais ações de tarifas no TJPB"}))
        tabelas["12_base_propria_x_demais"] = dist(comp, "grupo_base")
        if "status_base_propria" in g1d.columns:
            arq = g1[g1["base_propria"] & g1.get("status_base_propria", pd.Series(dtype=str)).fillna("").str.startswith("ARQ")]
            tabelas["13_arquivados_pre_classificados"] = arq["grupo_ultima"].fillna("Sem decisão identificada").value_counts().rename_axis("Desfecho no DataJud").to_frame("Processos")
            arq_export = arq[["numero", "orgao", "data_ajuizamento", "decisao_ultima", "data_ultima_decisao", "decisoes_todas",
                 "transito_julgado", "alvara_levantamento", "arquivamento_definitivo"]].copy()
            arq_export["numero"] = arq_export["numero"].map(formatar_cnj)
            arq_export.to_csv(
                os.path.join(DIR, "arquivados_pre_classificados.csv"), index=False, encoding="utf-8-sig")

    inv = pd.DataFrame([{"codigo": k[0], "nome": k[1], "ocorrencias": v[0], "categoria_atribuida": v[1]} for k, v in inventario.items()])
    inv = inv.sort_values("ocorrencias", ascending=False)

    xlsx = os.path.join(DIR, "jurimetria_tjpb_tarifas.xlsx")
    with pd.ExcelWriter(xlsx) as w:
        pd.DataFrame({"Leia-me": [
            "Fonte: API pública do DataJud (CNJ), TJPB, assunto Tarifas (11807) e processos da base própria.",
            f"Gerado em {datetime.now(timezone.utc).strftime('%d/%m/%Y %H:%M')} UTC.",
            "Desfecho de 1º grau = última decisão identificada nos registros G1/JE, pelo nome do movimento (TPU).",
            "Êxito do autor = procedência, procedência em parte ou acordo homologado.",
            "Confira a aba movimentos_inventario: se algum movimento de decisão ficou sem categoria, ajuste as REGRAS no script.",
            "Atenção na aba movimentos_inventario: movimentos com ILEGITIMIDADE/INEPCIA podem excluir só uma parte ou emendar"
            " a inicial no meio do processo, sem encerrá-lo -- confira antes de aceitar como 'Extinção sem mérito'.",
            "A separação 1º grau x 2º grau depende do campo 'grau' (G1/JE vs. G2/TR/...); confira contra a saída de"
            " 'diagnostico' antes de confiar nas tabelas por grau.",
            "A API pública não traz partes: estas tabelas cobrem ações de tarifas contra qualquer banco, exceto a aba 12.",
            "Marcos: Rec. CNJ 159 (23/10/2024); Tema 1198 STJ (13/03/2025); IRDR TJPB 0812102-56.2025 inadmitido (02/02/2026).",
        ]}).to_excel(w, sheet_name="leia-me", index=False)
        for nome, t in tabelas.items():
            t.to_excel(w, sheet_name=nome[:31])
        inv.to_excel(w, sheet_name="movimentos_inventario", index=False)
        df_export = df.copy()
        df_export["numero"] = df_export["numero"].map(formatar_cnj)
        df_export.to_excel(w, sheet_name="processos_por_grau", index=False)

    # Gráficos
    sem = tabelas["4_resultado_semestre"]
    sem = sem[sem["Total"] >= args.min_n]
    if len(sem):
        fig, ax = plt.subplots(figsize=(11, 5))
        cores = ["#2F4B7C", "#6A8CC7", "#8FB9A8", "#C9A227", "#B5473A"]
        base_y = None
        for c, cor in zip(ordem, cores):
            v = sem[f"% {c}"].values
            ax.bar(sem.index, v, bottom=base_y, label=c, color=cor)
            base_y = v if base_y is None else base_y + v
        ax.set_ylabel("% das decisões de 1º grau")
        ax.set_title("Ações de tarifas no TJPB: desfecho em 1º grau por semestre da decisão")
        idx = list(sem.index)
        for rot, d in MARCOS:
            s = f"{d[:4]}-S{1 if int(d[5:7]) <= 6 else 2}"
            if s in idx:
                x = idx.index(s)
                ax.axvline(x - 0.5, color="black", ls="--", lw=1)
                ax.text(x - 0.45, 101, rot, fontsize=8, rotation=0)
        ax.set_ylim(0, 110)
        ax.legend(fontsize=8, ncol=5, loc="lower center", bbox_to_anchor=(0.5, -0.25))
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(os.path.join(DIR, "fig_desfecho_semestre.png"), dpi=200)
        plt.close()

    # Resumo em texto
    linhas = ["# Resumo automático — jurimetria TJPB (tarifas)", ""]
    for _, r in tabelas["1_universo"].iterrows():
        linhas.append(f"- {r['Indicador']}: {r['Valor']}")
    try:
        tabela_marcos_txt = tabelas["8_teste_marcos"].to_markdown(index=False)
    except ImportError:
        # to_markdown() existe sempre, mas só funciona com o pacote opcional
        # "tabulate" instalado -- sem ele, cai aqui em vez de quebrar.
        tabela_marcos_txt = tabelas["8_teste_marcos"].to_string(index=False)
    linhas += ["", "## Extinção sem mérito e êxito do autor por período (1º grau)", "", tabela_marcos_txt]
    sem_cat = inv[(inv["categoria_atribuida"] == "") & inv["nome"].fillna("").str.upper().str.contains("PROCED|EXTIN|JULGAD|JULGAMENTO|PROVIMENTO|HOMOLOG") & ~inv["nome"].fillna("").str.upper().str.contains("TRANSITO|TRÂNSITO")]
    if len(sem_cat):
        linhas += ["", "## Atenção: movimentos de decisão sem categoria (revisar regras)", ""]
        linhas += [f"- {r.codigo} {r.nome} ({r.ocorrencias})" for r in sem_cat.head(30).itertuples()]
    open(os.path.join(DIR, "resumo.md"), "w", encoding="utf-8").write("\n".join(linhas))
    print("\n".join(linhas))
    print(f"\nArquivos em {DIR}/: jurimetria_tjpb_tarifas.xlsx, resumo.md, fig_desfecho_semestre.png"
          + (", arquivados_pre_classificados.csv" if proprios else ""))


# ----------------------------------------------------------------------------- CLI

def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    for nome in ("diagnostico", "coletar"):
        sp = sub.add_parser(nome)
        sp.add_argument("--assunto", type=int, action="append", dest="assuntos")
        if nome == "coletar":
            sp.add_argument("--refazer", action="store_true", help="apaga a coleta anterior e recomeça")
    sp = sub.add_parser("enriquecer")
    sp.add_argument("--base", required=True, help="CSV com a coluna CNJ (ex.: base_consolidada_bradesco_sem_nomes.csv)")
    sp = sub.add_parser("analisar")
    sp.add_argument("--base", default=None, help="mesmo CSV da base própria, para marcar e comparar")
    sp.add_argument("--min-n", type=int, default=20, help="mínimo de decisões por semestre no gráfico")
    a = p.parse_args()
    if getattr(a, "assuntos", None) is None and a.cmd in ("diagnostico", "coletar"):
        a.assuntos = [ASSUNTO_TARIFAS]
    {"diagnostico": cmd_diagnostico, "coletar": cmd_coletar,
     "enriquecer": cmd_enriquecer, "analisar": cmd_analisar}[a.cmd](a)


if __name__ == "__main__":
    main()
