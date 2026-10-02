# -*- coding: utf-8 -*-
"""Análises de apoio ao D10 da Op CATRIMANI II (Passos 10.1–10.4 da auditoria de 29/set/2026).

Funções puras sobre os dados do ETL (catrimani_data). O saldo por NC é ESTIMADO: o saldo exato é o da célula UG·PI·ND.
"""
import datetime
import json
import os

import regras_nc

ARQ_METAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "metas_coter.json")
SALDO_MIN = 0.005
_ORDEM = {regras_nc.SIT_VENCIDO: 0, regras_nc.SIT_7D: 1, regras_nc.SIT_IMEDIATO: 2,
          regras_nc.SIT_30D: 3, regras_nc.SIT_SEM: 4, regras_nc.SIT_LONGO: 5}


def _dt(s):
    try:
        d, m, a = str(s).split("/")
        return datetime.date(int(a), int(m), int(d))
    except Exception:
        return None


def fila_acao(linhas):
    """10.1 — NCs com saldo, da mais urgente para a menos: situação do prazo e depois dias restantes."""
    itens = [l for l in linhas if l["cred"] > SALDO_MIN and l.get("status_slug") != "canc"]
    itens.sort(key=lambda l: (_ORDEM.get(l.get("situacao_prazo"), 9),
                              l["dias_para_prazo"] if l.get("dias_para_prazo") is not None else 9999,
                              -l["cred"]))
    return itens


def texto_cobranca(sigla, cod, itens, hoje, limite=8):
    """10.1 — texto pronto para colar (WhatsApp/e-mail) cobrando uma UG. Só afirma o que está na fonte."""
    total = sum(i["cred"] for i in itens)
    venc = [i for i in itens if i.get("situacao_prazo") == regras_nc.SIT_VENCIDO]
    urg = [i for i in itens if i.get("situacao_prazo") == regras_nc.SIT_7D]
    brl = lambda v: "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    L = [f"*{sigla}* (UG {cod}) — Op CATRIMANI II / Ação 21EM — posição de {hoje.strftime('%d/%m/%Y')}",
         f"Saldo estimado a empenhar: {brl(total)} em {len(itens)} NC(s)."]
    if venc:
        L.append(f"PRAZO VENCIDO: {brl(sum(i['cred'] for i in venc))} em {len(venc)} NC(s) — regularizar com o COTER.")
    if urg:
        L.append(f"Vence em até 7 dias: {brl(sum(i['cred'] for i in urg))} em {len(urg)} NC(s).")
    for i in itens[:limite]:
        pz = i.get("prazo") or ""
        dd = f"até {pz[8:10]}/{pz[5:7]}" if pz else i.get("situacao_prazo", "").lower()
        L.append(f"- NC {i['nc'][-6:]} · ND {i['nd']} · {brl(i['cred'])} · {dd}")
    if len(itens) > limite:
        L.append(f"- … e mais {len(itens) - limite} NC(s)")
    L.append("Lembrete: informar na observação da NE o número da NC do COTER.")
    L.append("(Saldo por NC é estimado; o exato é o da célula UG · PI · ND.)")
    return "\n".join(L)


def recolhimentos(linhas):
    """10.2 — anulações/recolhimentos de descentralização: valor devolvido por mês e por emitente.

    Devolver saldo não é achado negativo: mostra quanto a operação devolveu.
    """
    por_mes, por_emit, total, n = {}, {}, 0.0, 0
    for l in linhas:
        for ln in l["linhas"]:
            v = ln.get("prov", 0.0) - ln.get("conc", 0.0)
            anula = "ANULA" in (ln.get("op") or "").upper() or "CANCEL" in (ln.get("op") or "").upper()
            if v < -SALDO_MIN and (anula or not l.get("is_det")):
                d = _dt(ln.get("dia") or l.get("dia"))
                mes = d.strftime("%Y-%m") if d else "sem data"
                por_mes[mes] = por_mes.get(mes, 0.0) + (-v)
                por_emit[l["emit"]] = por_emit.get(l["emit"], 0.0) + (-v)
                total += -v
                n += 1
    return {"total": total, "n": n, "por_mes": dict(sorted(por_mes.items())), "por_emit": por_emit}


def idade_ociosa(linhas, hoje):
    """10.3 — idade do saldo parado, por UG: dias desde a emissão da NC (média ponderada pelo saldo e a mais antiga).

    A idade até o PRIMEIRO EMPENHO da célula (indicador da DGOF) exige a data da NE — consulta do Passo 9.
    """
    por_ug = {}
    for l in linhas:
        if l["cred"] <= SALDO_MIN or l.get("is_det"):
            continue
        d = _dt(l.get("dia"))
        if not d:
            continue
        idade = (hoje - d).days
        u = por_ug.setdefault(l["ug"], {"saldo": 0.0, "peso": 0.0, "max": 0, "nc_antiga": ""})
        u["saldo"] += l["cred"]
        u["peso"] += l["cred"] * idade
        if idade >= u["max"]:
            u["max"], u["nc_antiga"] = idade, l["nc"]
    for u in por_ug.values():
        u["media"] = u["peso"] / u["saldo"] if u["saldo"] else 0.0
    return por_ug


def carregar_metas(caminho=None):
    """10.4 — metas do COTER como PARÂMETRO (data/metas_coter.json), nunca fixas no código."""
    try:
        with open(caminho or ARQ_METAS, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"fonte": "", "marcos": []}


def meta_vigente(metas, hoje):
    for m in metas.get("marcos", []):
        try:
            if datetime.date.fromisoformat(m["de"]) <= hoje <= datetime.date.fromisoformat(m["ate"]):
                return m
        except Exception:
            continue
    return None


def desvio_metas(por_ug, meta):
    """Desvio (p.p.) do % empenhado e do % liquidado de cada UG frente à meta vigente (None se a meta não traz o marco)."""
    out = {}
    for u in por_ug:
        dot = u["prov"] - u["conc"]
        pe = u["emp"] / dot * 100 if dot > 0 else 0.0
        pl = u["liq"] / u["emp"] * 100 if u["emp"] > 0 else 0.0
        out[u["cod"]] = {
            "pct_emp": pe, "pct_liq": pl,
            "d_emp": (pe - meta["emp"]) if meta and meta.get("emp") is not None else None,
            "d_liq": (pl - meta["liq"]) if meta and meta.get("liq") is not None else None,
        }
    return out
