import datetime as dt

from automacoes import cnj
from automacoes.prazos import Calendario, calcular_prazo
from automacoes.robos import robo_planilha as rp
from automacoes.unidades import Unidades, unidade_pelo_orgao
from tests.test_modelo_planilha import numero
from tests.test_robo_planilha import HOJE, item, linhas_prazos, planilha  # noqa: F401  (fixture)

COMARCAS = [
    {"Tribunal": "TJPB", "Código de origem": "0121", "Comarca": "Santa Rita", "Unidade atual": "Santa Rita/Bayeux"},
    {"Tribunal": "TJPB", "Código de origem": "0751", "Comarca": "Bayeux", "Unidade atual": "Santa Rita/Bayeux"},
    {"Tribunal": "TJPB", "Código de origem": 231, "Comarca": "Mamanguape"},          # código lido como número
    {"Tribunal": "TJPB", "Código de origem": "0000", "Comarca": "TJPB – 2º grau (originário)"},
]


def test_orgao_reconhece_municipio_capital_e_nucleo():
    assert unidade_pelo_orgao("1ª Vara da Comarca de Canguaretama", "TJRN") == "Canguaretama"
    assert unidade_pelo_orgao("1ª Vara de Família de Campina Grande", "TJPB") == "Campina Grande"
    assert unidade_pelo_orgao("Juizado Especial Cível de Baía da Traição", "TJPB") == "Baía da Traição"
    assert unidade_pelo_orgao("1ª V. Cumprimentos da Capital", "TJPB") == "João Pessoa (Capital)"
    assert unidade_pelo_orgao("Núcleo de Justiça 4.0 de Saúde Suplementar", "TJPB") == "Núcleo 4.0 – Saúde Suplementar"
    assert unidade_pelo_orgao("Vara de Feitos Especiais", "TJPB") is None          # "Feitos" não é município
    assert unidade_pelo_orgao("3ª Câmara Cível", "TJPB") is None                  # 2º grau: vale a origem


def test_resolver_unificacao_e_origem():
    u = Unidades(COMARCAS)
    assert u.resolver("TJPB", "0751", "1ª Vara Mista de Bayeux") == "Santa Rita/Bayeux"     # nome antigo no órgão
    assert u.resolver("TJPB", "0751", "") == "Santa Rita/Bayeux"                          # só o código
    assert u.resolver("TJPB", "0121", "2ª Vara Mista de Santa Rita") == "Santa Rita/Bayeux"
    assert u.resolver("TJPB", "0231", "3ª Câmara Cível") == "Mamanguape"                  # 2º grau: comarca de origem
    assert u.resolver("TJPB", "0231", "Núcleo de Justiça 4.0 de Saúde Suplementar") == "Núcleo 4.0 – Saúde Suplementar"
    assert u.resolver("TJPR", "0194", "") == "Cód. 0194"


def test_feriado_da_unidade_unificada_vale_para_os_dois_codigos():
    feriados = [{"data": dt.date(2026, 10, 13), "descricao": "Feriado municipal", "abrangencia": "TJPB:Santa Rita/Bayeux",
                 "tipo": "FERIADO", "conferido": "S"}]
    for cod in ("0121", "0751"):
        cal = Calendario.de_registros(feriados, "TJPB", cod, unidade="Santa Rita/Bayeux")
        assert calcular_prazo(dt.date(2026, 10, 9), 2, cal, HOJE).vencimento_interno == "2026-10-14"
    outra = Calendario.de_registros(feriados, "TJPB", "0231", unidade="Mamanguape")
    assert calcular_prazo(dt.date(2026, 10, 9), 2, outra, HOJE).vencimento_interno == "2026-10-13"


def _planilha_antiga(p):
    """Simula a planilha já em uso: Comarcas sem 'Unidade atual' e sem Santa Rita/Bayeux."""
    p.abas["Comarcas"] = [l[:5] for l in p.abas["Comarcas"]
                          if not (len(l) > 1 and str(l[1]) in ("0121", "0751"))]


def test_robo_cria_coluna_registra_unificacao_e_corrige(planilha):
    _planilha_antiga(planilha)
    linhas, aba = linhas_prazos(planilha)
    vazia = max(l["_n"] for l in linhas) + 1
    antigo = numero("0000020", "15", "0751")
    planilha.atualizar("Prazos", [{"range": f"{aba.letra['Processo']}{vazia}", "values": [[antigo]]},
                                  {"range": f"{aba.letra['Comarca']}{vazia}", "values": [["Cód. 0751"]]},
                                  {"range": f"{aba.letra['Tribunal']}{vazia}", "values": [["TJPB"]]}])
    robo = rp.Robo(planilha, HOJE)
    robo.atualizar_unidades()
    com = rp.Aba("Comarcas", planilha.ler("Comarcas"), 1)
    assert "Unidade atual" in com.letra
    assert {(str(l["Código de origem"]), l["Unidade atual"]) for l in com.linhas if l.get("Unidade atual")} == \
        {("0121", "Santa Rita/Bayeux"), ("0751", "Santa Rita/Bayeux")}
    corrigida = next(l for l in linhas_prazos(planilha)[0] if l["Processo"] == antigo)
    assert corrigida["Comarca"] == "Santa Rita/Bayeux" and robo.rel.comarcas_atualizadas >= 1


def test_djen_com_codigo_novo_e_painel(planilha):
    robo = rp.Robo(planilha, HOJE)
    robo.atualizar_unidades()
    tjpr = f"0000626-{cnj.calcular_dv('0000626', '2023', '8', '16', '0194')}.2023.8.16.0194"
    bayeux = numero("0000030", "15", "0751")
    robo.processar_djen([item(tjpr, "2026-10-05", "Intime-se.", "Vara Cível de Ponta Grossa"),
                         item(bayeux, "2026-10-05", "Manifeste-se em 5 dias.", "1ª Vara Mista de Bayeux")])
    robo.atualizar_painel()
    com = rp.Aba("Comarcas", planilha.ler("Comarcas"), 1)
    novo = next(l for l in com.linhas if str(l["Código de origem"]) == "0194")
    assert (novo["Tribunal"], novo["Comarca"], novo["Conferido"]) == ("TJPR", "Ponta Grossa", "N")
    linhas = linhas_prazos(planilha)[0]
    assert next(l for l in linhas if l["Processo"] == bayeux)["Comarca"] == "Santa Rita/Bayeux"
    painel = planilha.ler("Painel")
    combos = {(l[0], l[1]) for l in painel if len(l) > 1}
    assert ("TJPB", "Santa Rita/Bayeux") in combos and ("TJPR", "Ponta Grossa") in combos
