import datetime as dt

import pytest
from openpyxl import Workbook, load_workbook

from automacoes import cnj
from automacoes.robos import calculadora_bacen as calc
from automacoes.robos import robo_prazos as robo
from automacoes.robos.comum import COL_DATA_1, COL_PROC_1, LINHA_INICIAL, mapear_processos

D = dt.date.fromisoformat
N1 = f"0000001-{cnj.calcular_dv('0000001', '2026', '8', '15', '0231')}.2026.8.15.0231"
N2 = f"0000002-{cnj.calcular_dv('0000002', '2026', '8', '15', '0231')}.2026.8.15.0231"


def planilha(tmp_path, processos=(N1,)):
    wb = Workbook()
    ws = wb.active
    ws.title = "TJPB"
    for i, n in enumerate(processos):
        ws.cell(row=LINHA_INICIAL + i, column=COL_PROC_1, value=n)
    caminho = tmp_path / "Prazos Processuais.xlsx"
    wb.save(caminho)
    return caminho


# ---------------- calculadora ----------------

@pytest.mark.parametrize("entrada, esperado", [
    ("1.500", 1500.0), ("1.500,90", 1500.90), ("15,90", 15.90), ("R$ 2.000.000,00", 2_000_000.0),
    (12.5, 12.5), ("12.5", 12.5), ("1500", 1500.0),
])
def test_ler_valor(entrada, esperado):
    assert calc.ler_valor(entrada) == esperado


def test_fator_correcao_e_mes_faltante():
    indices = {(2024, 1): 0.01, (2024, 2): 0.02}
    assert calc.fator_correcao(D("2024-01-15"), D("2024-03-01"), indices) == pytest.approx(1.01 * 1.02)
    with pytest.raises(ValueError, match="03/2024"):
        calc.fator_correcao(D("2024-01-15"), D("2024-04-01"), indices)


def test_desconto_posterior_a_data_base_nao_trava():
    with pytest.raises(ValueError, match="posterior"):
        calc.fator_correcao(D("2062-01-01"), D("2024-03-01"), {})


def test_data_base_padrao():
    assert calc.data_base_padrao({(2026, 7): 0.0, (2026, 8): 0.0}) == D("2026-09-01")
    assert calc.data_base_padrao({(2025, 12): 0.0}) == D("2026-01-01")


def test_linha_com_dobra_e_juros_da_citacao():
    p = calc.Parametros("INPC", D("2024-03-01"), D("2024-02-10"), "citacao", 0.01, dobra=True)
    r = calc.calcular_linha(D("2024-01-15"), "100,00", p, {(2024, 1): 0.01, (2024, 2): 0.02})
    assert r["Valor base (dobro)"] == 200.0
    assert r["Valor corrigido (INPC)"] == 206.04   # 200 x 1,01 x 1,02
    assert r["Meses de juros"] == 1                 # citação 02/2024 -> data-base 03/2024
    assert (r["Juros"], r["Total"]) == (2.06, 208.10)


def test_memorial_completo(tmp_path):
    ent = tmp_path / "e.xlsx"
    wb = Workbook()
    wb.active.append(["Data", "Valor", "Rubrica"])
    wb.active.append(["15/01/2024", "1.500", "RMC"])
    wb.active.append(["15/05/2024", "10,00", "Seguro"])   # sem índice de 05/2024 -> recusada
    wb.active.append([None, None, None])
    wb.save(ent)
    p = calc.Parametros("INPC", D("2024-03-01"), None, "desconto", 0.0, dobra=False)
    cab, dados, idx = calc.ler_planilha(ent)
    n_ok, n_erro, total = calc.processar(cab, dados, idx, tmp_path / "s.xlsx", p, {(2024, 1): 0.0, (2024, 2): 0.0}, 188)
    assert (n_ok, n_erro, total) == (1, 1, 1500.0)
    out = load_workbook(tmp_path / "s.xlsx")
    assert out.sheetnames == ["Cálculo", "Índices", "Premissas"]
    assert "RECUSADA" in out["Cálculo"].cell(row=3, column=len(cab) + 7).value


def test_premissas_avisam_lei_14905():
    p = calc.Parametros("INPC", D("2026-09-01"), D("2019-05-28"), "citacao", 0.01, dobra=True)
    textos = " ".join(v for _, v in calc.premissas(p, 188, 1, 0))
    assert "14.905" in textos and "art. 42" in textos


# ---------------- robô de prazos ----------------

class Resp:
    def __init__(self, status, dados=None):
        self.status_code, self._dados, self.text = status, dados, ""

    def json(self):
        return self._dados


class SessaoFalsa:
    def __init__(self, respostas):
        self.respostas, self.chamadas = list(respostas), []

    def get(self, url, **kw):
        self.chamadas.append(kw.get("params"))
        return self.respostas.pop(0)


def test_djen_pagina_ate_lote_incompleto():
    s = SessaoFalsa([Resp(200, {"items": [{}] * 100}), Resp(200, {"items": [{}] * 30})])
    itens = robo.consultar_djen(s, "1", "PB", D("2026-09-24"), D("2026-09-30"))
    assert len(itens) == 130 and [c["pagina"] for c in s.chamadas] == [1, 2]


def test_djen_403_levanta_falha_sem_repetir(monkeypatch):
    monkeypatch.setattr(robo.time, "sleep", lambda s: None)
    s = SessaoFalsa([Resp(403)])
    with pytest.raises(robo.FalhaConsulta, match="403"):
        robo.consultar_djen(s, "1", "PB", D("2026-09-24"), D("2026-09-30"))
    assert len(s.chamadas) == 1


def test_erro_de_rede_tem_limite(monkeypatch):
    monkeypatch.setattr(robo.time, "sleep", lambda s: None)

    class Caida:
        n = 0

        def get(self, *a, **k):
            Caida.n += 1
            raise robo.requests.ConnectionError("sem rede")
    with pytest.raises(robo.FalhaConsulta):
        robo.consultar_djen(Caida(), "1", "PB", D("2026-09-24"), D("2026-09-30"))
    assert Caida.n == robo.TENTATIVAS


def test_data_mais_recente_vence_ordem_da_api(tmp_path):
    wb = load_workbook(planilha(tmp_path))
    mapa = mapear_processos(wb)
    itens = [
        {"numero_processo": cnj.digitos(N1), "texto": "Intime-se no prazo de 5 dias.", "data_disponibilizacao": "2026-09-18"},
        {"numero_processo": cnj.digitos(N1), "texto": "Ciência.", "data_disponibilizacao": "2026-09-10"},
    ]
    triagem = robo.trilha_djen(wb, mapa, itens, {}, D("2026-09-21"))
    assert robo.ler_data(wb["TJPB"].cell(row=LINHA_INICIAL, column=COL_DATA_1).value) == D("2026-09-18")
    calculada = next(l for l in triagem if l["Prazo (dias)"])
    assert calculada["Vencimento interno"] == D("2026-09-25") and calculada["Criticidade"] == "URGENTE"


def test_horas_e_anos_nao_viram_dias():
    for texto in ("no prazo de 48 horas", "prazo de 1 (um) ano"):
        l = robo.linha_triagem(D("2026-09-21"), "Vara", N1, texto, {}, D("2026-09-21"))
        assert l["Prazo (dias)"] == "" and l["Revisar"]


def test_hash_e_aliases():
    assert robo.hash_movimento({"dataHora": "2026-01-01", "codigo": 123})
    assert robo.alias_datajud(N1) == "api_publica_tjpb"
    assert robo.alias_datajud("0000001-00.2026.4.05.0001".replace("00.2026", f"{cnj.calcular_dv('0000001','2026','4','05','0001')}.2026")) == "api_publica_trf5"


def test_main_cai_para_datajud_quando_djen_falha(tmp_path, monkeypatch):
    caminho = planilha(tmp_path)
    cfg = tmp_path / "c.ini"
    cfg.write_text(f"[planilha]\ncaminho = {caminho}\n[djen]\noab = 1\nuf = PB\ndias_busca = 7\n", encoding="utf-8")

    def djen_falha(*a, **k):
        raise robo.FalhaConsulta("HTTP 403")
    monkeypatch.setattr(robo, "consultar_djen", djen_falha)
    monkeypatch.setattr(robo, "consultar_datajud", lambda s, n, c: {"dataHora": "2026-09-15T10:00:00", "codigo": 51, "nome": "Conclusão"})
    assert robo.main(["--config", str(cfg), "--hoje", "2026-09-30"]) == 1
    wb = load_workbook(caminho)
    assert wb["Triagem_Diaria"].cell(row=2, column=4).value == "NOVO MOVIMENTO: Conclusão"
    assert (tmp_path / "revisar_ia.txt").exists()
    assert any((tmp_path / "backups").iterdir())


def test_processo_novo_do_djen_e_cadastrado(tmp_path):
    wb = load_workbook(planilha(tmp_path, [N1]))
    mapa = mapear_processos(wb)
    itens = [
        {"numero_processo": cnj.digitos(N2), "texto": "Intime-se no prazo de 5 dias.",
         "data_disponibilizacao": "2026-09-18", "nomeOrgao": "1ª Vara de Mamanguape"},
        {"numero_processo": cnj.digitos(N2), "texto": "Ciência.", "data_disponibilizacao": "2026-09-10"},
        {"numero_processo": "0000001-35.2020.8.15.0001", "texto": "Ciência.", "data_disponibilizacao": "2026-09-10"},
    ]
    triagem = robo.trilha_djen(wb, mapa, itens, {}, D("2026-09-21"))
    ws = wb["TJPB"]
    assert ws.cell(row=LINHA_INICIAL + 1, column=COL_PROC_1).value == N2
    assert ws.cell(row=LINHA_INICIAL + 2, column=COL_PROC_1).value is None  # CNJ inválido não entra
    assert robo.ler_data(ws.cell(row=LINHA_INICIAL + 1, column=COL_DATA_1).value) == D("2026-09-18")
    assert "processo novo cadastrado" in triagem[0]["Revisar"]
    assert "processo novo" not in triagem[1]["Revisar"]  # cadastrado uma vez só


def test_cadastro_de_novos_desligavel(tmp_path):
    wb = load_workbook(planilha(tmp_path, [N1]))
    itens = [{"numero_processo": cnj.digitos(N2), "texto": "Ciência.", "data_disponibilizacao": "2026-09-10"}]
    robo.trilha_djen(wb, mapear_processos(wb), itens, {}, D("2026-09-21"), cadastrar_novos=False)
    assert wb["TJPB"].cell(row=LINHA_INICIAL + 1, column=COL_PROC_1).value is None
