"""Critérios de aceite da auditoria de 29/set/2026 (valores medidos por recálculo independente
sobre a posição de 11/set/2026, só Ação 21EM)."""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from gerar_dashboard import etl  # noqa: E402

FIX = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures", "credito_disp_11set2026.xlsx")


@pytest.fixture(scope="module")
def catr():
    return etl(FIX)[3]


def test_totais_21em(catr):
    t = catr["totais"]
    assert round(t["prov"], 2) == 10407342.69
    assert round(t["emp"], 2) == 8950534.87
    assert round(t["liq"], 2) == 4132406.05
    assert round(t["pag"], 2) == 3889473.08
    assert round(t["cred"], 2) == 1456807.82


def test_ugs_21em(catr):
    assert {u["cod"] for u in catr["por_ug"]} == {
        "160006", "160007", "160014", "160016", "160021", "160352", "160353", "160482", "160907"}


def test_so_21em(catr):
    assert all(l["acao"] == "21EM" for l in catr["linhas"])


def test_saldo_por_pi(catr):
    assert {k: round(v, 2) for k, v in catr["cred_por_pi"].items()} == {
        "OCS90001000": 1224797.12, "OCS90001001": 217961.26, "OCS90001002": 14049.44}


def test_saldo_nc_fecha_com_celula(catr):
    # soma dos saldos atribuídos às NCs == saldo exato de cada (UG, ND)
    for u in catr["por_ug"]:
        exato = {n["nd"]: n["cred"] for n in u["nds_list"]}
        atrib = {}
        for n in u["ncs"]:
            atrib[n["nd"]] = atrib.get(n["nd"], 0) + n["cred"]
        for nd, v in exato.items():
            assert abs(v - atrib.get(nd, 0)) < 0.01, (u["cod"], nd)


def test_chave_unica(catr):
    chaves = [f'{l["nc"]}|{l["ug"]}|{l["nd"]}' for l in catr["linhas"]]
    assert len(chaves) == len(set(chaves))


def test_integridade(catr):
    assert catr["integridade"]["ok"], catr["integridade"]["falhas"]


def test_quadro_prazos_fecha_com_disponivel(catr):
    assert abs(sum(catr["prazos"]["total"].values()) - catr["totais"]["cred"]) < 0.01


def test_correlatos_fora_dos_totais(catr):
    # 4 NCs da Ação 2000 (R$ 88.121,54 na fixture) ficam num quadro à parte e não somam
    assert round(sum(c["cred"] for c in catr["correlatos"]), 2) == 88121.54
    assert all(c["acao"] != "21EM" for c in catr["correlatos"])


def test_sem_cifra_literal_no_codigo():
    """Nenhuma cifra 'R$ <dígito>' fixa no HTML/JS do gerador (a exceção é o rótulo de filtro 'R$ 0')."""
    import re
    src = open(os.path.join(os.path.dirname(FIX), "..", "..", "gerar_dashboard.py"), encoding="utf-8").read()
    achados = [m.group(0) for m in re.finditer(r"R\$ ?\d[\d.,]*", src) if m.group(0).replace(" ", "") != "R$0"]
    assert not achados, achados[:10]
