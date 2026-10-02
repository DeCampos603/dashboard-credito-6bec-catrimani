# -*- coding: utf-8 -*-
"""Regras de leitura do objeto da NC (prazo de empenho e travas), usadas pelo dashboard e pelo e-mail.

O prazo vem do texto livre do campo "NC - Descrição". A referência de "hoje" é sempre
informada por quem chama (data da execução), nunca lida do relógio aqui dentro.
"""
import datetime
import re

MESES = {'JAN': 1, 'FEV': 2, 'MAR': 3, 'ABR': 4, 'MAI': 5, 'JUN': 6,
         'JUL': 7, 'AGO': 8, 'SET': 9, 'OUT': 10, 'NOV': 11, 'DEZ': 12}
RE_PRAZO = re.compile(
    r"(?:EMPENHO\s+AT[EÉ]?|PRAZO\s+(?:DE\s+)?EMPENHO:?)\s+(?:O\s+DIA\s+)?"
    r"(\d{1,2})\s*(?:DE\s+)?([A-ZÇ]{3,9})\s*(?:DE\s+)?(\d{2,4})")
RE_IMEDIATO = re.compile(r"EMPENHO\s+IMEDIATO")
RE_TRAVA = re.compile(r"N[ÃA]O\s+DEVE\s+ALTERAR\s+ND/?UGR|SOMENTE\s+P/?\s*COTER|APLICACAO\s+RESTRITA|USO\s+RESTRITO")

SIT_VENCIDO = "VENCIDO"
SIT_7D = "VENCE EM ≤7 DIAS"
SIT_30D = "VENCE EM ≤30 DIAS"
SIT_LONGO = "PRAZO > 30 DIAS"
SIT_IMEDIATO = "EMPENHO IMEDIATO"
SIT_SEM = "SEM PRAZO NA NC"
SITUACOES = (SIT_VENCIDO, SIT_7D, SIT_30D, SIT_IMEDIATO, SIT_SEM, SIT_LONGO)


def prazo_empenho(obj):
    """Devolve (data_limite | None, tipo) com tipo em DATA | IMEDIATO | SEM_PRAZO."""
    s = (obj or "").upper()
    m = RE_PRAZO.search(s)
    if m:
        d, mes, a = m.groups()
        a = int(a)
        a = a + 2000 if a < 100 else a
        try:
            return datetime.date(a, MESES[mes[:3]], int(d)), "DATA"
        except (KeyError, ValueError):
            pass
    if RE_IMEDIATO.search(s):
        return None, "IMEDIATO"
    return None, "SEM_PRAZO"


def situacao_prazo(prazo, tipo, hoje):
    if tipo == "IMEDIATO":
        return SIT_IMEDIATO
    if prazo is None:
        return SIT_SEM
    dias = (prazo - hoje).days
    if dias < 0:
        return SIT_VENCIDO
    if dias <= 7:
        return SIT_7D
    if dias <= 30:
        return SIT_30D
    return SIT_LONGO


def tem_trava(obj):
    return bool(RE_TRAVA.search((obj or "").upper()))


def anotar_prazo(registro, hoje):
    """Grava prazo, situação, dias e trava num registro de NC (dict) a partir do campo obj."""
    prazo, tipo = prazo_empenho(registro.get("obj"))
    registro["prazo"] = prazo.isoformat() if prazo else ""
    registro["prazo_tipo"] = tipo
    registro["situacao_prazo"] = situacao_prazo(prazo, tipo, hoje)
    registro["dias_para_prazo"] = (prazo - hoje).days if prazo else None
    registro["trava_nd_ugr"] = tem_trava(registro.get("obj"))
    return registro
