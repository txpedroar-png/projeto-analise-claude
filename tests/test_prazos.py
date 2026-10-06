import datetime as dt

import pytest

from automacoes.prazos import Calendario, calcular_prazo, criticidade

D = dt.date.fromisoformat
CAB = "data,descricao,abrangencia,tipo,conferido,fonte\n"


@pytest.fixture
def cal():
    return Calendario()  # CSV versionado do repositório


def csv_temp(tmp_path, linhas):
    p = tmp_path / "f.csv"
    p.write_text(CAB + "\n".join(linhas) + "\n", encoding="utf-8")
    return p


def test_recesso_limites():
    assert Calendario.em_recesso(D("2026-12-20")) and Calendario.em_recesso(D("2027-01-20"))
    assert not Calendario.em_recesso(D("2026-12-19")) and not Calendario.em_recesso(D("2027-01-21"))


def test_prazo_atravessa_recesso(cal):
    # 17/12 (qui) D0; 18/12 = 1; recesso 20/12-20/01; 21/01 = 2, 22 = 3, 25 = 4, 26 = 5
    r = calcular_prazo(D("2026-12-17"), 5, cal, hoje=D("2026-12-17"))
    assert (r.inicio_contagem, r.vencimento_interno, r.termo_legal) == ("2026-12-18", "2027-01-26", "2027-01-27")
    assert any("recesso" in a for a in r.avisos)
    assert any("2027" in a for a in r.avisos)  # CSV não cobre 2027


def test_disponibilizacao_sexta_com_feriado_conferido(cal):
    # 09/10 (sex) D0; 12/10 feriado nacional; 13 = 1 ... 19 = 5
    r = calcular_prazo(D("2026-10-09"), 5, cal, hoje=D("2026-10-09"))
    assert (r.inicio_contagem, r.vencimento_interno, r.termo_legal) == ("2026-10-13", "2026-10-19", "2026-10-20")


def test_feriado_nao_conferido_nao_adia_mas_avisa(cal):
    # Carnaval 16-17/02 com conferido=N conta como dia útil (regra conservadora)
    r = calcular_prazo(D("2026-02-13"), 3, cal, hoje=D("2026-02-13"))
    assert r.vencimento_interno == "2026-02-18"
    assert sum("NÃO CONFERIDO" in a for a in r.avisos) == 3


def test_feriado_conferido_adia(tmp_path):
    p = csv_temp(tmp_path, ["2026-02-16,Carnaval,TJPB,FERIADO,S,x", "2026-02-17,Carnaval,TJPB,FERIADO,S,x"])
    r = calcular_prazo(D("2026-02-13"), 3, Calendario(p), hoje=D("2026-02-13"))
    assert r.vencimento_interno == "2026-02-20"


def test_feriado_municipal_so_na_comarca(tmp_path):
    p = csv_temp(tmp_path, ["2026-10-13,Padroeira,TJPB:0231,FERIADO,S,x"])
    # CSV temporário sem o 12/10: 12 = 1, 13 = 2 (sem a comarca) ou 14 = 2 (com a comarca)
    sem = calcular_prazo(D("2026-10-09"), 2, Calendario(p), hoje=D("2026-10-09"))
    com = calcular_prazo(D("2026-10-09"), 2, Calendario(p, comarca="0231"), hoje=D("2026-10-09"))
    assert (sem.vencimento_interno, com.vencimento_interno) == ("2026-10-13", "2026-10-14")


def test_feriado_de_outro_tribunal_ignorado(tmp_path):
    p = csv_temp(tmp_path, ["2026-10-13,Feriado RN,TJRN,FERIADO,S,x"])
    assert calcular_prazo(D("2026-10-12"), 1, Calendario(p), hoje=D("2026-10-12")).vencimento_interno == "2026-10-13"


def test_csv_ausente_falha_explicitamente(tmp_path):
    with pytest.raises(FileNotFoundError):
        Calendario(tmp_path / "nao_existe.csv")


def test_csv_com_data_invalida_falha(tmp_path):
    with pytest.raises(ValueError, match="linha 2"):
        Calendario(csv_temp(tmp_path, ["2026-13-45,Erro,NACIONAL,FERIADO,S,x"]))


@pytest.mark.parametrize("dias", [0, -1, 2.5, True])
def test_dias_prazo_invalido(cal, dias):
    with pytest.raises(ValueError):
        calcular_prazo(D("2026-10-09"), dias, cal)


def test_dias_uteis_restantes_e_criticidade(cal):
    venc = D("2026-10-19")  # segunda
    assert cal.dias_uteis_entre(D("2026-10-16"), venc) == 1  # sexta -> segunda
    assert cal.dias_uteis_entre(venc, venc) == 0
    assert cal.dias_uteis_entre(D("2026-10-21"), venc) == -2
    assert [criticidade(n) for n in (6, 5, 1, 0, -1)] == ["NORMAL", "URGENTE", "URGENTE", "CRITICO", "VENCIDO"]


def test_resultado_criticidade_integrada(cal):
    r = calcular_prazo(D("2026-10-09"), 5, cal, hoje=D("2026-10-19"))
    assert (r.dias_uteis_restantes, r.criticidade, r.cor) == (0, "CRITICO", "#f4cccc")
