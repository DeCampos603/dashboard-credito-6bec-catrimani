# -*- coding: utf-8 -*-
"""
Dashboard Interativo de Crédito — 6º Batalhão de Engenharia de Construção (6º BEC) & Operação Catrimani II
UASGs: 160353 (OGU) e 167353 (Fundo do Exército) · Boa Vista / RR
Acompanhamento Orçamentário Multi-UGs da Ação Governamental 21EM (Operação Catrimani II)

- Lê o export do Tesouro Gerencial 'CRÉDITO DISP 160353.xlsx'.
- Detalhamento integral do 6º BEC (160353 / 167353) em par OGU/FEx.
- Aba dedicada exclusiva para a Operação Catrimani II com as 10 UGs executoras.
- Validação anti-falha: Crédito Disponível = Recebido − Concedido − Empenhado.
- Escreve site/index.html (autocontido: CSS Moderno + Google Fonts + SVG + Tabela + Excel) e site/data/history.json.
"""
import os, sys, json, argparse, datetime, urllib.request, tempfile, html, math, re, shutil
import unicodedata, csv
import openpyxl

HDR_ROW, DATA_ROW = 8, 9
UNIDADES = [
    {"sigla": "6º BEC", "nome": "6º Batalhão de Engenharia de Construção", "ogu": "160353", "fex": "167353", "logo": "6BEC.png", "accent": "#15803D", "key": "BEC6"},
]

OMDS_COMPARATIVO = [
    {"sigla": "6º BEC", "nome": "6º Batalhão de Engenharia de Construção", "ogu": "160353", "fex": "167353", "logo": "6BEC.png", "accent": "#15803D", "key": "BEC6"},
    {"sigla": "7º BIS", "nome": "Comando de Fronteira Roraima / 7º BIS", "ogu": "160352", "fex": "167352", "logo": "12RM.png", "accent": "#047857", "key": "BIS7"},
    {"sigla": "1ª Bda Inf Sl", "nome": "Comando 1ª Brigada de Infantaria de Selva", "ogu": "160482", "fex": "167482", "logo": "12RM.png", "accent": "#1D4ED8", "key": "BDA1"},
    {"sigla": "4º B Av Ex", "nome": "4º Batalhão de Aviação do Exército", "ogu": "160007", "fex": "167007", "logo": "12RM.png", "accent": "#D97706", "key": "BAV4"},
    {"sigla": "1º B Log Sl", "nome": "1º Batalhão Logístico de Selva", "ogu": "160907", "fex": "167907", "logo": "12RM.png", "accent": "#B91C1C", "key": "BLOG1"},
    {"sigla": "Pq R Mnt/12", "nome": "Parque Regional de Manutenção da 12ª RM", "ogu": "160021", "fex": "167021", "logo": "12RM.png", "accent": "#6D28D9", "key": "PQ12"},
    {"sigla": "1º BIS (AMV)", "nome": "1º Batalhão de Infantaria de Selva (Amv)", "ogu": "160006", "fex": "167006", "logo": "12RM.png", "accent": "#0F766E", "key": "BIS1"},
    {"sigla": "Cmdo 12ª RM", "nome": "Comando da 12ª Região Militar", "ogu": "160014", "fex": "167014", "logo": "12RM.png", "accent": "#991B1B", "key": "RM12"},
    {"sigla": "Cmdo CMA", "nome": "Comando Militar da Amazônia", "ogu": "160016", "fex": "167016", "logo": "12RM.png", "accent": "#1E3A8A", "key": "CMA"}
]

UASG_TO_OMDS = {}
for _u in OMDS_COMPARATIVO:
    UASG_TO_OMDS[_u["ogu"]] = (_u, "OGU")
    UASG_TO_OMDS[_u["fex"]] = (_u, "FEx")

def _par(u):
    return [(u["ogu"], f'{u["sigla"]} · OGU'), (u["fex"], f'{u["sigla"]} · FEx')]

ALVOS = [p for u in UNIDADES for p in _par(u)]
FONTE_CURTA = {"160353": "160", "167353": "167"}
DEFAULT_FILE_ID = None
DEFAULT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTVtnLCf2tvVO1-PFklLro4Y-ijBqw9h3psRi2y3Q69_1TSX75OPmph7yPK3zmANA/pub?gid=991377463&single=true&output=csv"
HERE = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(HERE, "site")
DATA = os.path.join(HERE, "data")
HISTFILE = os.path.join(DATA, "history.json")
DEFAULT_SRC = os.path.join(DATA, "CRÉDITO DISP 160353.xlsx")

# ---------------- leitura ----------------
def norm(s):
    if s is None: return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("utf-8")
    return s.strip().upper()

def to_num(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace("'", "").replace('"', '')
    if s in ("", "-", "-9", "NAO SE APLICA", "NÃO SE APLICA"): return 0.0
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    try:
        val = float(s.replace(".", "").replace(",", "."))
        return -val if neg else val
    except ValueError:
        return 0.0

def disp(v):
    s = "" if v is None else str(v).strip().replace("'", "")
    return "" if s in ("-9", "NAO SE APLICA", "NÃO SE APLICA") else s

def baixar(target=None):
    if target and os.path.exists(target):
        return target

    url = target if (target and target.startswith(("http://", "https://"))) else (os.environ.get("SHEETS_CSV_URL") or DEFAULT_CSV_URL)

    if target and not target.startswith(("http://", "https://")) and len(target) > 15 and not os.path.exists(target):
        url = f"https://docs.google.com/spreadsheets/d/{target}/export?format=xlsx"

    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (dashboard-6bec)"})
    ext = ".csv" if "output=csv" in url else ".xlsx"
    tmp = os.path.join(tempfile.gettempdir(), f"credito_disp_160353_download{ext}")
    try:
        with urllib.request.urlopen(req, timeout=120) as r, open(tmp, "wb") as f:
            f.write(r.read())
        if os.path.getsize(tmp) < 500:
            if os.path.exists(DEFAULT_SRC):
                print(f"[AVISO] Download pequeno ({os.path.getsize(tmp)} bytes). Usando fallback: {DEFAULT_SRC}")
                return DEFAULT_SRC
            raise SystemExit("Download muito pequeno — verifique o link público da planilha.")
        return tmp
    except Exception as e:
        if os.path.exists(DEFAULT_SRC):
            print(f"[AVISO] Falha no download ({e}). Usando fallback: {DEFAULT_SRC}")
            return DEFAULT_SRC
        raise

def ler_linhas(path):
    is_csv = str(path).lower().endswith(".csv")
    if not is_csv and not str(path).lower().endswith((".xlsx", ".xlsm")):
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                first = f.read(500)
                if "," in first or ";" in first:
                    is_csv = True
        except Exception:
            pass

    if is_csv:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            sample = f.read(2048)
            f.seek(0)
            delim = ";" if sample.count(";") > sample.count(",") else ","
            reader = csv.reader(f, delimiter=delim)
            for r in reader:
                yield tuple(r)
    else:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = None
        for nm in wb.sheetnames:
            if "SALDO R$" in norm(nm) or "CREDITO DISP" in norm(nm) or "CRÉDITO DISP" in norm(nm):
                ws = wb[nm]; break
        if ws is None: ws = wb.active
        for r in ws.iter_rows(values_only=True):
            yield tuple(r)


# ---------------- ETL ----------------
def _dt_br(s):
    try:
        dd, mm, aa = str(s).split("/")
        return datetime.date(int(aa), int(mm), int(dd))
    except Exception:
        return None

def anotar_trocas_nd(res):
    for cod, d in res.items():
        linhas = d["linhas"]
        if not linhas:
            continue
        por_nc = {}
        for L in linhas:
            if L["nc"]:
                por_nc.setdefault(L["nc"], []).append(L)
        trocas = {}
        for nc, ls in por_nc.items():
            grupos = {}
            for L in ls:
                grupos.setdefault((L["acao"], L["pi"]), []).append(L)
            for k, gl in grupos.items():
                pos = [x for x in gl if x["cred"] > 0.005]
                neg = [x for x in gl if x["cred"] < -0.005]
                if pos and neg and abs(sum(x["cred"] for x in gl)) <= 0.05:
                    trocas.setdefault(nc, {})[k] = (neg[0]["nd"], pos[0]["nd"], pos[0]["cred"])
        if not trocas:
            continue
        entradas = {}
        for L in linhas:
            if L["cred"] > 0.005 and L["nc"]:
                entradas.setdefault((L["acao"], L["pi"], L["nd"]), []).append(L)

        def rastrear(nc_alt, acao, pi, nd_org, valor, prof=0, visto=None):
            if visto is None:
                visto = set()
            if prof > 6 or nc_alt in visto:
                return None
            visto.add(nc_alt)
            dt_alt = None
            for L in por_nc.get(nc_alt, []):
                dt_alt = _dt_br(L["dia"])
                if dt_alt:
                    break
            cands = [L for L in entradas.get((acao, pi, nd_org), [])
                     if L["nc"] != nc_alt
                     and (_dt_br(L["dia"]) or datetime.date.min) <= (dt_alt or datetime.date.max)]
            if not cands:
                return None
            exatos = [c for c in cands if abs(c["cred"] - valor) < 0.05]
            cands = exatos or sorted(cands, key=lambda c: (_dt_br(c["dia"]) or datetime.date.min), reverse=True)
            orig = cands[0]
            t = trocas.get(orig["nc"], {}).get((acao, pi))
            if t:
                acima = rastrear(orig["nc"], acao, pi, t[0], t[2], prof + 1, visto)
                if acima:
                    return acima
            return orig

        for nc, mp in trocas.items():
            for (acao, pi), (nd_org, nd_dst, val) in mp.items():
                orig = rastrear(nc, acao, pi, nd_org, val)
                for L in por_nc.get(nc, []):
                    if L["acao"] == acao and L["pi"] == pi and L["cred"] > 0.005:
                        L["nd_de"] = nd_org
                        if orig:
                            L["nc_origem"] = {"nc": orig["nc"], "obj": orig["obj"], "dia": orig["dia"],
                                              "emit": orig["emit"], "op": orig["op"]}

def etl(path):
    all_rows = list(ler_linhas(path))
    if not all_rows:
        raise SystemExit(f"Arquivo vazio ou ilegível: {path}")

    header_rows = all_rows[:35]
    hdr_row = None
    for r_idx, row_vals in enumerate(header_rows, 1):
        row_norm = [norm(v) for v in row_vals if v is not None]
        if "CREDITO DISPONIVEL" in row_norm or "PROVISAO RECEBIDA" in row_norm:
            hdr_row = r_idx; break
    if not hdr_row:
        raise SystemExit("Não encontrei o cabeçalho (PROVISAO RECEBIDA / CREDITO DISPONIVEL) na planilha.")

    hdr_line = header_rows[hdr_row - 1]
    hdr = {norm(val): c_idx for c_idx, val in enumerate(hdr_line) if norm(val)}

    def col(name, req=True):
        c = hdr.get(norm(name))
        if c is None and req:
            raise SystemExit(f"Coluna '{name}' não encontrada na planilha")
        return c

    C = dict(prov=col("PROVISAO RECEBIDA"), cred=col("CREDITO DISPONIVEL"),
             emp=col("DESPESAS EMPENHADAS"), liq=col("DESPESAS LIQUIDADAS"), pag=col("DESPESAS PAGAS"),
             conc=col("PROVISAO CONCEDIDA", req=False))

    hdrL = {}
    start_search = max(0, hdr_row - 5)
    for row_vals in header_rows[start_search:hdr_row]:
        for c_idx, val in enumerate(row_vals):
            nn = norm(val)
            if nn and nn not in hdrL:
                hdrL[nn] = c_idx

    def colL(name, fb_0idx):
        c = hdrL.get(norm(name))
        return c if c is not None else fb_0idx

    CA    = colL("ACAO GOVERNO", 5)
    CPI   = colL("PI", 6)
    CPID  = CPI + 1
    CND   = colL("NATUREZA DESPESA", 8)
    CNDD  = CND + 1
    CNC   = colL("NC", 4)
    COBJ  = colL("NC - DESCRICAO", 10)
    COP   = colL("NC - OPERACAO (TIPO)", 11)
    CDIA  = colL("NC - DIA EMISSAO", 12)
    CEMIT = colL("EMITENTE - UG", 0)

    data_row = None
    for r_idx, row_vals in enumerate(all_rows[hdr_row:], hdr_row + 1):
        if len(row_vals) > 2 and row_vals[2] is not None:
            val3 = norm(row_vals[2]).replace("'", "")
            if val3.isdigit():
                data_row = r_idx; break
    if not data_row:
        data_row = hdr_row + 1

    periodo = None
    for row_vals in all_rows[max(0, hdr_row - 4):min(len(all_rows), hdr_row + 2)]:
        for v in row_vals:
            if v:
                vs = str(v).strip().upper()
                if re.match(r"^[A-Z]{3}/\d{4}$", vs):
                    periodo = vs; break
        if periodo:
            break

    res = {c: {"prov": 0.0, "conc": 0.0, "cred": 0.0, "emp": 0.0, "liq": 0.0, "pag": 0.0,
               "nome": label, "linhas": [], "celulas": {}, "por_acao": {}, "por_nd": {}, "nd_nome": {}, "n": 0}
           for c, label in ALVOS}
    codigos_alvo = set(res.keys())
    ugs_presentes = set()

    # Benchmarking e Comparativo de OMDS
    omds_totais = {
        u["key"]: {
            "key": u["key"], "sigla": u["sigla"], "nome": u["nome"], "logo": u["logo"],
            "accent": u["accent"], "ogu": u["ogu"], "fex": u["fex"],
            "prov": 0.0, "conc": 0.0, "emp": 0.0, "liq": 0.0, "pag": 0.0, "cred": 0.0,
            "n_linhas": 0, "n_ncs": set()
        } for u in OMDS_COMPARATIVO
    }

    # Operação Catrimani
    catrimani_totais = {"prov": 0.0, "conc": 0.0, "dot": 0.0, "cred": 0.0, "emp": 0.0, "liq": 0.0, "pag": 0.0, "count": 0}
    catrimani_por_ug = {}
    catrimani_por_nd = {}
    catrimani_linhas = []

    nomes_padrao_ugs = {
        "160352": "CMDO FRON RR / 7º BIS",
        "160482": "CMDO 1ª BDA INF SL",
        "160353": "6º B E CNST",
        "160016": "CMDO C M A",
        "160907": "1º B LOG SL",
        "160007": "4º B AV EX",
        "160006": "1º B I S (AMV)",
        "167482": "CMDO 1ª BDA INF SL (FEx)",
        "160014": "CMDO 12ª RM",
        "160021": "PQ R MNT/12",
        "160329": "BCMS",
        "160238": "BA AP LOG",
    }

    siglas_padrao_ugs = {
        "160007": "4º BAVEX",
        "160353": "6º BEC",
        "160352": "7º BIS",
        "160482": "1ª BDA INF SL",
        "160907": "1º B LOG SL",
        "160016": "CMDO CMA",
        "160006": "1º BIS (AMV)",
        "167482": "1ª BDA (FEx)",
        "160014": "CMDO 12ª RM",
        "160021": "PQ R MNT/12",
        "160329": "BCMS",
        "160238": "BA AP LOG",
    }

    def get_val(row_tup, idx):
        return row_tup[idx] if idx is not None and 0 <= idx < len(row_tup) else None

    for row in all_rows[data_row - 1:]:
        if not row: continue
        ug_raw = get_val(row, 2)
        if ug_raw is None: continue
        ug = str(ug_raw).strip().replace("'", "")
        if not ug.isdigit(): continue
        ugs_presentes.add(ug)

        prov = to_num(get_val(row, C["prov"]))
        conc = to_num(get_val(row, C["conc"])) if C["conc"] is not None else 0.0
        cred = to_num(get_val(row, C["cred"]))
        emp  = to_num(get_val(row, C["emp"]))
        liq  = to_num(get_val(row, C["liq"]))
        pag  = to_num(get_val(row, C["pag"]))

        acao      = disp(get_val(row, CA))
        pi        = disp(get_val(row, CPI))
        pi_nome   = disp(get_val(row, CPID))
        nd        = disp(get_val(row, CND))
        nd_nome   = disp(get_val(row, CNDD))
        nc        = disp(get_val(row, CNC))
        dia       = disp(get_val(row, CDIA))
        emit      = disp(get_val(row, CEMIT))
        emit_nome = disp(get_val(row, CEMIT + 1 if CEMIT is not None else None))
        fav_nome  = disp(get_val(row, 3))
        op        = disp(get_val(row, COP))
        obj       = disp(get_val(row, COBJ))

        if not obj:
            for ci in (14, 15, 16, 17, 18, 19):
                if ci < len(row):
                    cand = disp(row[ci])
                    if len(cand) > 10 and not cand.replace(".", "").replace(",", "").replace("-", "").isdigit():
                        obj = cand; break

        is_nc = bool(nc and nc not in ("-9", "NAO SE APLICA", "NÃO SE APLICA"))

        # Regra Mestre de Sanidade Orçamentária SIAFI/TG (Agente-Execucao-Orcamentaria/conhecimento/09):
        # Se a linha não é Nota de Crédito (is_nc == False), somente processa se o emitente for a própria UG executora (emit == ug).
        # Linhas sem NC onde emit != ug (ex: Emit=160907 -> Exec=160482) são cruzamentos interunidades/subtotais de controle
        # no relatório do TG que duplicam empenhos e violam a invariante fiscal (Empenho <= Dotação).
        if not is_nc and emit != ug:
            continue

        # Acúmulo de dados para o Benchmarking das OMDS da Amazônia
        if ug in UASG_TO_OMDS:
            u_info, u_fonte = UASG_TO_OMDS[ug]
            ot = omds_totais[u_info["key"]]
            ot["prov"] += prov
            ot["conc"] += conc
            ot["cred"] += cred
            ot["emp"]  += emp
            ot["liq"]  += liq
            ot["pag"]  += pag
            ot["n_linhas"] += 1
            if nc and nc not in ("-9", "NAO SE APLICA", "NÃO SE APLICA"):
                ot["n_ncs"].add(nc)

        # 1. 6º BEC (160353 OGU / 167353 FEx) - Visão Integral
        if ug in codigos_alvo:
            d = res[ug]
            d["prov"] += prov; d["conc"] += conc; d["cred"] += cred
            d["emp"]  += emp;  d["liq"]  += liq;  d["pag"]  += pag
            d["n"]    += 1

            d["linhas"].append(dict(
                acao=acao, pi=pi, pi_nome=pi_nome, nd=nd, nd_desc=nd_nome,
                nc=nc, dia=dia, emit=emit, emit_nome=emit_nome, fav_nome=fav_nome or d["nome"], op=op, obj=obj,
                prov=prov, conc=conc, cred=cred, emp=emp, liq=liq, pag=pag))

            if acao: d["por_acao"][acao] = d["por_acao"].get(acao, 0.0) + cred
            if nd:
                d["por_nd"][nd] = d["por_nd"].get(nd, 0.0) + cred
                if nd_nome and nd not in d["nd_nome"]: d["nd_nome"][nd] = nd_nome

            k_cel = (acao, pi, nd)
            if k_cel not in d["celulas"]:
                d["celulas"][k_cel] = {"acao": acao, "pi": pi, "pi_nome": pi_nome, "nd": nd, "nd_nome": nd_nome,
                                       "prov": 0.0, "conc": 0.0, "cred": 0.0, "emp": 0.0, "liq": 0.0, "pag": 0.0,
                                       "cpos": 0.0, "cneg": 0.0, "n_nc": 0, "ncs": []}
            cel = d["celulas"][k_cel]
            cel["prov"] += prov; cel["conc"] += conc; cel["cred"] += cred
            cel["emp"]  += emp;  cel["liq"]  += liq;  cel["pag"]  += pag
            if cred >= 0: cel["cpos"] += cred
            else:         cel["cneg"] += -cred
            if nc:
                cel["n_nc"] += 1
                cel["ncs"].append(dict(nc=nc, dia=dia, emit=emit, emit_nome=emit_nome, op=op,
                                       prov=prov, conc=conc, cred=cred, emp=emp, liq=liq, pag=pag, obj=obj))

        # 2. Operação Catrimani (Ação 21EM ou texto Catrimani)
        row_full = f"{acao} {pi} {pi_nome} {obj}".upper()
        if acao == "21EM" or "CATRIMANI" in row_full:
            has_fin = any(abs(v) > 0.005 for v in [prov, conc, cred, emp, liq, pag])
            is_nc = bool(nc and nc not in ("-9", "NAO SE APLICA", "NÃO SE APLICA"))

            # Descarta linhas puramente 'fantasmas' do relatório do Tesouro (sem NC e com tudo zerado)
            if not has_fin and not is_nc:
                continue

            catrimani_totais["prov"] += prov
            catrimani_totais["conc"] += conc
            catrimani_totais["cred"] += cred
            catrimani_totais["emp"]  += emp
            catrimani_totais["liq"]  += liq
            catrimani_totais["pag"]  += pag
            catrimani_totais["count"] += 1

            if ug not in catrimani_por_ug:
                nom = nomes_padrao_ugs.get(ug, fav_nome or f"UG {ug}")
                sigla_m = siglas_padrao_ugs.get(ug, ug)
                catrimani_por_ug[ug] = {
                    "cod": ug, "sigla": sigla_m, "nome": nom, "prov": 0.0, "conc": 0.0,
                    "cred": 0.0, "emp": 0.0, "liq": 0.0, "pag": 0.0, "count": 0,
                    "ncs": [], "nds": {}
                }
            u_catr = catrimani_por_ug[ug]
            u_catr["prov"] += prov; u_catr["conc"] += conc; u_catr["cred"] += cred
            u_catr["emp"]  += emp;  u_catr["liq"]  += liq;  u_catr["pag"]  += pag
            u_catr["count"] += 1

            if nd:
                if nd not in catrimani_por_nd:
                    catrimani_por_nd[nd] = {"nd": nd, "nome": nd_nome or nd, "prov": 0.0, "cred": 0.0, "emp": 0.0, "liq": 0.0}
                nd_c = catrimani_por_nd[nd]
                nd_c["prov"] += prov; nd_c["cred"] += cred; nd_c["emp"] += emp; nd_c["liq"] += liq

                if nd not in u_catr["nds"]:
                    u_catr["nds"][nd] = {"nd": nd, "nome": nd_nome or nd, "prov": 0.0, "emp": 0.0, "cred": 0.0, "liq": 0.0, "pag": 0.0}
                u_nd = u_catr["nds"][nd]
                u_nd["prov"] += prov; u_nd["emp"] += emp; u_nd["cred"] += cred; u_nd["liq"] += liq; u_nd["pag"] += pag

            # Apenas linhas com Nota de Crédito real compõem o Extrato de NCs (evita poluir com empenhos negativos sem NC)
            if is_nc:
                nc_reg = {
                    "nc": nc, "dia": dia, "ug": ug, "ug_nome": nomes_padrao_ugs.get(ug, fav_nome or ug),
                    "emit": emit, "emit_nome": emit_nome, "acao": acao, "pi": pi, "pi_nome": pi_nome,
                    "nd": nd, "nd_desc": nd_nome, "op": op, "obj": obj,
                    "prov": prov, "conc": conc, "cred": cred, "emp": emp, "liq": liq, "pag": pag
                }
                catrimani_linhas.append(nc_reg)
                u_catr["ncs"].append(nc_reg)

    total_linhas = sum(d["n"] for d in res.values())
    if total_linhas == 0:
        raise SystemExit("Nenhuma linha das UASGs do 6º BEC na planilha.")

    for d in res.values():
        for cel in d["celulas"].values():
            cel["aloc"] = cel["cred"] + cel["emp"]

    anotar_trocas_nd(res)

    alertas = []
    for cod, d in res.items():
        if d["n"] == 0: continue
        saldo_calc = d["prov"] - d["conc"] - d["emp"]
        if abs(saldo_calc - d["cred"]) > 0.05:
            alertas.append(f"{d['nome']}: Crédito Disponível ({d['cred']:.2f}) difere de Recebido−Concedido−Empenhado ({saldo_calc:.2f}).")
        if d["emp"] < -0.01:
            alertas.append(f"{d['nome']}: Empenhado negativo ({d['emp']:.2f}).")

    # Cálculos Catrimani
    catrimani_totais["dot"] = catrimani_totais["prov"] - catrimani_totais["conc"]
    catrimani_totais["pct_emp"] = (catrimani_totais["emp"] / catrimani_totais["dot"] * 100.0) if catrimani_totais["dot"] > 0 else 0.0
    catrimani_totais["pct_liq"] = (catrimani_totais["liq"] / catrimani_totais["emp"] * 100.0) if catrimani_totais["emp"] > 0 else 0.0

    ugs_catr_list = list(catrimani_por_ug.values())
    for u_c in ugs_catr_list:
        dot_u = u_c["prov"] - u_c["conc"]
        u_c["dot"] = dot_u
        u_c["pct_emp"] = (u_c["emp"] / dot_u * 100.0) if dot_u > 0 else 0.0
        u_c["pct_liq"] = (u_c["liq"] / u_c["emp"] * 100.0) if u_c["emp"] > 0 else 0.0
        u_c["n_ncs"] = len(u_c["ncs"])
        u_c["nds_list"] = sorted(list(u_c["nds"].values()), key=lambda x: x["prov"], reverse=True)
    ugs_catr_list.sort(key=lambda x: x["prov"], reverse=True)

    nds_catr_list = list(catrimani_por_nd.values())
    nds_catr_list.sort(key=lambda x: x["emp"], reverse=True)

    catrimani_ncs_distintas = len(set(x["nc"] for x in catrimani_linhas if x.get("nc")))

    for ot in omds_totais.values():
        ot["n_ncs_cnt"] = len(ot["n_ncs"])
        ot["n_ncs"] = ot["n_ncs_cnt"]

    catrimani_data = {
        "totais": catrimani_totais,
        "por_ug": ugs_catr_list,
        "por_nd": nds_catr_list,
        "linhas": catrimani_linhas,
        "ncs_distintas": catrimani_ncs_distintas,
        "total_linhas_brutas": len(catrimani_linhas)
    }

    return res, periodo, alertas, catrimani_data, omds_totais

def atualizar_historico(res, data_str):
    os.makedirs(DATA, exist_ok=True)
    hist = []
    if os.path.exists(HISTFILE):
        try:
            with open(HISTFILE, "r", encoding="utf-8") as f:
                hist = json.load(f)
        except Exception:
            hist = []
    hist = [h for h in hist if h.get("data") != data_str]
    snap = {"data": data_str}
    for cod, _ in ALVOS:
        snap[cod] = {k: round(res[cod][k], 2) for k in ("prov", "conc", "cred", "emp", "liq", "pag")}
    snap["total"] = {k: round(sum(res[c][k] for c, _ in ALVOS), 2) for k in ("prov", "conc", "cred", "emp", "liq", "pag")}
    hist.append(snap)
    hist.sort(key=lambda h: h.get("data", ""))
    with open(HISTFILE, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=1)
    return hist

# ---------------- formatação ----------------
def _fmt(v):
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s

def brl(v):
    return ("−R$ " if v < 0 else "R$ ") + _fmt(v)

def num(v):
    return ("−" if v < 0 else "") + _fmt(v)

def pct(a, b): return (100.0 * a / b) if b else 0.0

def abrev(v):
    a = abs(v); s = "−" if v < 0 else ""
    if a >= 1e6: return s + f"{a/1e6:.1f}".replace(".", ",") + " mi"
    if a >= 1e3: return s + f"{a/1e3:.0f} mil"
    return s + f"{a:.0f}"

def esc(s): return html.escape(str(s))

# ---------------- SVG ----------------
def _r(x, y, w, h, var, extra=""):
    fallback = "#10B981" if "success" in var else ("#F59E0B" if "warning" in var else ("#EF4444" if "danger" in var else "#2563EB"))
    return f'<rect x="{x:.1f}" y="{y:.1f}" width="{max(0,w):.1f}" height="{h:.1f}" style="fill:var(--{var}, {fallback})" {extra}/>'

def svg_util(recebido, empenhado, disponivel):
    """Barra empilhada de utilização: Recebido = Empenhado + Disponível."""
    if recebido <= 0:
        return '<p class="vazio">Sem provisão recebida</p>'
    x0, x1, y, h, W, H = 4, 636, 40, 32, 640, 92
    plot = x1 - x0
    fe = max(0.0, min(1.0, empenhado / recebido))
    we = plot * fe
    wd = plot - we
    pe, pd = pct(empenhado, recebido), pct(disponivel, recebido)
    return f'''<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="Utilização do crédito recebido">
<title>Provisão Recebida {brl(recebido)}: Empenhado {brl(empenhado)} ({pe:.1f}%) + Crédito Disponível {brl(disponivel)} ({pd:.1f}%)</title>
<text x="{x0}" y="20" class="s-lbl">PROVISÃO RECEBIDA TOTAL · {esc(brl(recebido))}</text>
<line x1="{x0}" y1="26" x2="{x1}" y2="26" class="s-brk"/><line x1="{x0}" y1="26" x2="{x0}" y2="32" class="s-brk"/><line x1="{x1}" y1="26" x2="{x1}" y2="32" class="s-brk"/>
{_r(x0, y, we, h, "warning-main", 'rx="6"')}
{_r(x0+we, y, wd, h, "success-main", 'rx="6"')}
<line x1="{x0+we:.1f}" y1="{y}" x2="{x0+we:.1f}" y2="{y+h}" style="stroke:var(--bg-surface);stroke-width:2.5"/>
<text x="{x0+8}" y="{y+h+17}" class="s-seg">Empenhado {esc(brl(empenhado))} · {pe:.1f}%</text>
<text x="{x1-8}" y="{y+h+17}" text-anchor="end" class="s-seg s-seg-ok">Crédito Disponível {esc(brl(disponivel))} · {pd:.1f}%</text>
</svg>'''

def svg_waterfall(recebido, empenhado, disponivel, mini=False):
    """Waterfall horizontal: Recebida → (−)Empenhado → (=)Disponível."""
    if recebido <= 0:
        return '<p class="vazio">Sem dados</p>'
    if mini:
        W, H, xL, rh, gap, fs = 360, 118, 112, 26, 10, 10
        L1, L2, L3 = "Recebido", "(−) Empenhado", "(=) Disponível"
    else:
        W, H, xL, rh, gap, fs = 720, 196, 190, 42, 16, 13
        L1, L2, L3 = "Provisão Recebida", "(−) Empenhado", "(=) Crédito Disponível"
    xR = W - 24
    plot = xR - xL
    R = recebido
    def X(v): return xL + (v / R) * plot
    y1 = 14; y2 = y1 + rh + gap; y3 = y2 + rh + gap
    xd = X(max(disponivel, 0))
    parts = [f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="Composição do crédito">',
             f'<title>Provisão Recebida {brl(recebido)} menos Empenhado {brl(empenhado)} igual a Crédito Disponível {brl(disponivel)}</title>',
             f'<desc>{esc(brl(recebido))} − {esc(brl(empenhado))} = {esc(brl(disponivel))}</desc>']
    # linha 1 — Recebida
    parts.append(_r(xL, y1, plot, rh, "primary-600", 'rx="5"'))
    parts.append(f'<text x="{xL-10}" y="{y1+rh/2+4:.0f}" text-anchor="end" class="s-cat">{esc(L1)}</text>')
    parts.append(f'<text x="{xR-8}" y="{y1+rh/2+4:.0f}" text-anchor="end" class="s-val s-on">{esc(brl(recebido))}</text>')
    # linha 2 — Empenhado (flutuante)
    we = plot - (xd - xL)
    parts.append(_r(xd, y2, we, rh, "warning-main", 'rx="5"'))
    parts.append(f'<text x="{xL-10}" y="{y2+rh/2+4:.0f}" text-anchor="end" class="s-cat">{esc(L2)}</text>')
    parts.append(f'<text x="{xd+8:.1f}" y="{y2+rh/2+4:.0f}" class="s-val s-on">−{esc(brl(empenhado).replace("R$ ","R$ "))}</text>')
    # linha 3 — Disponível
    parts.append(_r(xL, y3, xd - xL, rh, "success-main", 'rx="5"'))
    parts.append(f'<text x="{xL-10}" y="{y3+rh/2+4:.0f}" text-anchor="end" class="s-cat s-cat-ok">{esc(L3)}</text>')
    parts.append(f'<text x="{xd+8:.1f}" y="{y3+rh/2+4:.0f}" class="s-val s-ok">{esc(brl(disponivel))}</text>')
    # conectores tracejados
    parts.append(f'<line x1="{xR:.1f}" y1="{y1+rh}" x2="{xR:.1f}" y2="{y2}" class="s-conn"/>')
    parts.append(f'<line x1="{xd:.1f}" y1="{y2+rh}" x2="{xd:.1f}" y2="{y3}" class="s-conn"/>')
    parts.append('</svg>')
    return "".join(parts)

def svg_diverg(itens, titulo, max_itens=8):
    itens = [(k, v) for k, v in itens if round(v, 2) != 0]
    itens = sorted(itens, key=lambda x: abs(x[1]), reverse=True)
    resto = sum(v for _, v in itens[max_itens:])
    itens = itens[:max_itens]
    if round(resto, 2) != 0:
        itens.append(("Outras", resto))
    if not itens:
        return f'<div class="card chart"><div class="eyebrow">{esc(titulo)}</div><p class="vazio">Sem valores no período</p></div>'
    vmax = max(abs(v) for _, v in itens) or 1
    labW, rh, pad = 150, 32, 10
    zx = labW + 246
    half = 230
    W = zx + half + 70
    H = pad * 2 + rh * len(itens) + 16
    el = [f'<line x1="{zx}" y1="{pad}" x2="{zx}" y2="{pad+rh*len(itens):.0f}" class="s-zero"/>']
    for i, (k, v) in enumerate(itens):
        y = pad + i * rh
        w = abs(v) / vmax * half
        lbl = k if len(k) <= 24 else k[:23] + "…"
        el.append(f'<text x="{labW-8}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-cat">{esc(lbl)}</text>')
        if v >= 0:
            el.append(_r(zx, y+5, w, rh-10, "success-main", f'rx="4"><title>{esc(k)}: {esc(brl(v))}</title></rect'.replace("/>", ">")))
            el.append(f'<text x="{zx+w+8:.1f}" y="{y+rh/2+4:.0f}" class="s-num s-ok">{esc(num(v))}</text>')
        else:
            el.append(_r(zx-w, y+5, w, rh-10, "danger-main", f'rx="4"><title>{esc(k)}: {esc(brl(v))}</title></rect'.replace("/>", ">")))
            el.append(f'<text x="{zx-w-8:.1f}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-num s-neg">{esc(num(v))}</text>')
    svg = f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="{esc(titulo)}">{"".join(el)}</svg>'
    return f'<div class="card chart"><div class="eyebrow">{esc(titulo)}</div>{svg}</div>'

def svg_funil(cod, d):
    emp, liq, pag = d["emp"], d["liq"], d["pag"]
    base = emp or 1
    W, rh, gap, xL = 340, 26, 12, 110
    plot = W - xL - 75
    H = 18 + 3 * (rh + gap)
    rows = [("Empenhado", emp, "stg1", ""), ("Liquidado", liq, "stg2", f"{pct(liq,emp):.0f}% do emp."),
            ("Pago", pag, "stg3", f"{pct(pag,liq):.0f}% do liq.")]
    el = []
    for i, (nome, val, cls, conv) in enumerate(rows):
        y = 12 + i * (rh + gap)
        w = max(4, abs(val) / base * plot)
        el.append(f'<text x="{xL-10}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-cat">{nome}</text>')
        el.append(f'<rect x="{xL}" y="{y}" width="{w:.1f}" height="{rh}" rx="4" style="fill:var(--{cls})"><title>{nome}: {esc(brl(val))}</title></rect>')
        el.append(f'<text x="{xL+w+8:.1f}" y="{y+rh/2+4:.0f}" class="s-num s-on2">{esc(num(val))}</text>')
        if conv:
            el.append(f'<text x="{W-4}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-conv">{conv}</text>')
    svg = f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="Estágios {cod}">{"".join(el)}</svg>'
    return f'<div class="card chart"><div class="eyebrow">Estágios da Despesa · {esc(cod)} {esc(FONTE_CURTA[cod])}</div>{svg}</div>'

def svg_tendencia(hist):
    wk, order = {}, []
    for h in hist:
        try:
            y, m, dd = (int(x) for x in h["data"].split("-"))
            key = datetime.date(y, m, dd).isocalendar()[:2]
        except Exception:
            key = h.get("data")
        if key not in wk:
            order.append(key)
        wk[key] = h
    semanas = [wk[k] for k in order]
    pts = [(h["data"], h["total"]["cred"]) for h in semanas]
    W, H, pl, pb, pt, pr = 720, 220, 70, 36, 18, 80
    pw, ph = W - pl - pr, H - pb - pt
    vals = [v for _, v in pts]
    vmin, vmax = min(vals + [0]), max(vals + [1])
    rng = (vmax - vmin) or 1
    n = len(pts)
    def X(i): return pl + (pw * (i / (n - 1)) if n > 1 else pw / 2)
    def Y(v): return pt + ph - ((v - vmin) / rng * ph)
    el = []
    for t in range(4):
        val = vmin + rng * t / 3; y = Y(val)
        el.append(f'<line x1="{pl}" y1="{y:.1f}" x2="{W-pr}" y2="{y:.1f}" class="s-grid"/>')
        el.append(f'<text x="{pl-10}" y="{y+4:.1f}" text-anchor="end" class="s-ax">{esc(abrev(val))}</text>')
    if n > 1:
        line = "M" + " L".join(f"{X(i):.1f},{Y(v):.1f}" for i, (_, v) in enumerate(pts))
        area = f"M{X(0):.1f},{Y(vmin):.1f} L" + " L".join(f"{X(i):.1f},{Y(v):.1f}" for i, (_, v) in enumerate(pts)) + f" L{X(n-1):.1f},{Y(vmin):.1f} Z"
        el.append(f'<path d="{area}" class="s-area"/>')
        el.append(f'<path d="{line}" class="s-line"/>')
    step = max(1, n // 6)
    for i, (dt, v) in enumerate(pts):
        show = (i % step == 0 or i == n - 1)
        el.append(f'<circle cx="{X(i):.1f}" cy="{Y(v):.1f}" r="4" class="s-dot"><title>{esc(dt)}: {esc(brl(v))}</title></circle>')
        if show:
            el.append(f'<text x="{X(i):.1f}" y="{H-pb+18}" text-anchor="middle" class="s-ax">{dt[8:10]}/{dt[5:7]}</text>')
    if n >= 1:
        dt, v = pts[-1]
        el.append(f'<text x="{X(n-1):.1f}" y="{Y(v)-12:.1f}" text-anchor="end" class="s-num s-ok">{esc(brl(v))}</text>')
    nota = "" if n > 1 else '<p class="vazio">A curva semanal se desenvolve a partir da 2ª semana de histórico.</p>'
    svg = f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="Tendência semanal do Crédito Disponível">{"".join(el)}</svg>'
    return f'<div class="card chart wide"><div class="eyebrow">Tendência Histórica · Crédito Disponível Consolidado</div>{svg}{nota}</div>'

# ---------------- componentes HTML ----------------
def kpi_tile(label, valor, chip, cls, id_v="", onclick=""):
    chip_html = f'<span class="chip">{esc(chip)}</span>' if chip else ""
    id_attr = f' id="{esc(id_v)}"' if id_v else ""
    clk = onclick if onclick else f"bcmsModalKpi('{cls}')"
    return (f'<div class="kpi kpi-{cls}" tabindex="0" role="button" onclick="{clk}" '
            f'title="Clique para ver o detalhamento deste indicador (SIAFI)"><div class="kpi-l">{esc(label)}</div>'
            f'<div class="kpi-v num"{id_attr}>{esc(valor)}</div>{chip_html}</div>')

def uasg_card(cod, d):
    barp = pct(d["emp"], d["prov"])
    return (f'<div class="card uasg">'
            f'<div class="uasg-h"><span class="uasg-cod num">{esc(cod)}</span>'
            f'<span class="pill-fonte">{esc(FONTE_CURTA[cod])}</span>'
            f'<span class="uasg-nome">{esc(d["nome"].split("·")[1].strip())}</span></div>'
            f'<div class="uasg-disp"><span class="uasg-disp-l">Crédito Disponível</span>'
            f'<span class="uasg-disp-v num">{esc(brl(d["cred"]))}</span></div>'
            f'<div class="uasg-eq num">{esc(brl(d["prov"]))} <i>−</i> {esc(brl(d["emp"]))}</div>'
            f'{svg_waterfall(d["prov"], d["emp"], d["cred"], mini=True)}'
            f'<div class="uasg-exec"><div class="exec-l"><span>Empenhado / Recebido</span><span class="num">{barp:.1f}%</span></div>'
            f'<div class="exec-track"><div class="exec-fill" style="width:{min(barp,100):.1f}%"></div></div></div>'
            f'</div>')

def tabela_html(tid, celulas, com_fonte, ativo):
    """Relação de crédito EM TELA por célula orçamentária (saldo líquido positivo)."""
    cols = (["Fonte"] if com_fonte else []) + ["Ação", "PI", "ND", "Aplicação", "Recebido (líq)", "Empenhado", "Crédito Disp."]
    ths = []
    for c in cols:
        numc = c in ("Recebido (líq)", "Empenhado", "Crédito Disp.")
        cls = ' class="num"' if numc else ''
        ths.append(f'<th{cls} tabindex="0" role="button" aria-sort="none" onclick="bcmsSort(this)" onkeydown="if(event.key==\'Enter\'||event.key==\' \'){{event.preventDefault();bcmsSort(this)}}">{esc(c)}<span class="sort"></span></th>')
    body = []
    tot = sum(c["cred"] for c in celulas)
    for c in celulas:
        fonte = f'<td><span class="pill-fonte">{esc(FONTE_CURTA.get(c.get("uasg",""),""))}</span></td>' if com_fonte else ''
        aplic = c.get("nd_nome") or c.get("pi_nome") or ""
        cid = esc(c.get("cid", ""))
        body.append(
            f'<tr class="cel-row" tabindex="0" role="button" data-cel="{cid}" title="Ver as notas de crédito desta célula (descrição completa)" '
            f'onclick="bcmsCel(this)" onkeydown="if(event.key==\'Enter\'||event.key==\' \'){{event.preventDefault();bcmsCel(this)}}">'
            f'{fonte}<td>{esc(c["acao"])}</td><td class="mono2">{esc(c["pi"])}</td><td class="mono2">{esc(c["nd"])}</td>'
            f'<td class="obj" title="{esc(aplic)}">{esc(aplic[:60])}</td>'
            f'<td class="num" data-sort="{c["aloc"]:.2f}">{esc(brl(c["aloc"]))}</td>'
            f'<td class="num" data-sort="{c["emp"]:.2f}">{esc(brl(c["emp"]))}</td>'
            f'<td class="num anchor" data-sort="{c["cred"]:.2f}">{esc(brl(c["cred"]))}<i class="chev" aria-hidden="true">›</i></td></tr>')
    ncols = len(cols)
    tfoot = (f'<tfoot><tr><td colspan="{ncols-1}">TOTAL · {len(celulas)} célula(s) com crédito em tela</td>'
             f'<td class="num anchor">{esc(brl(tot))}</td></tr></tfoot>')
    disp_style = "" if ativo else ' style="display:none"'
    return (f'<div class="tabpanel" id="{tid}" role="tabpanel"{disp_style}>'
            f'<div class="tbl-tools"><label class="visually-hidden" for="q-{tid}">Buscar</label>'
            f'<input type="search" id="q-{tid}" class="tbl-search" placeholder="Buscar por ação, PI, ND ou aplicação…" oninput="bcmsSearch(this,\'{tid}\')">'
            f'<button type="button" class="btn-excel" onclick="bcmsExportTable(this,\'{tid}\',\'creditos_em_tela_{tid}\')" title="Baixar dados em planilha formatada para Excel"><span class="btn-excel-ic">📊</span> Exportar Excel</button>'
            f'<span class="tbl-count" id="cnt-{tid}" data-unit="célula(s)" aria-live="polite">{len(celulas)} células</span></div>'
            f'<div class="tbl-scroll"><table class="det"><thead><tr>{"".join(ths)}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody>{tfoot}</table></div></div>')

# ---------------- página da unidade ----------------
def conteudo_unidade(res, hist, data_str, periodo, u, u_hist_items=None):
    if u_hist_items is None:
        u_hist_items = []
    ALVOS = _par(u); sfx = u["key"]
    tot = {k: sum(res[c][k] for c, _ in ALVOS) for k in ("prov", "conc", "cred", "emp", "liq", "pag", "n")}
    ger = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M")
    posicao = periodo if periodo else (data_str[8:10] + "/" + data_str[5:7] + "/" + data_str[0:4])

    delta_html = ""
    if len(hist) >= 2:
        dv = hist[-1]["total"]["cred"] - hist[-2]["total"]["cred"]
        if round(dv, 2) != 0:
            seta = "▲" if dv > 0 else "▼"
            cls = "up" if dv > 0 else "down"
            delta_html = f'<div class="delta {cls}"><span>{seta}</span> {esc(num(dv))} <small>vs. dia anterior</small></div>'
        else:
            delta_html = '<div class="delta flat">Sem variação vs. dia anterior</div>'
    else:
        delta_html = '<div class="delta flat">1º dia de histórico</div>'

    kpis = (kpi_tile("Provisão Recebida", brl(tot["prov"]), "", "prov") +
            kpi_tile("Empenhado", brl(tot["emp"]), f'{pct(tot["emp"],tot["prov"]):.1f}% do recebido', "emp") +
            kpi_tile("Liquidado", brl(tot["liq"]), f'{pct(tot["liq"],tot["emp"]):.1f}% do empenhado', "liq") +
            kpi_tile("Pago", brl(tot["pag"]), f'{pct(tot["pag"],tot["liq"]):.1f}% do liquidado', "pag"))

    cards_uasg = "".join(uasg_card(cod, res[cod]) for cod, _ in ALVOS)

    acao, nd, nd_nome = {}, {}, {}
    for cod, _ in ALVOS:
        for k, v in res[cod]["por_acao"].items(): acao[k] = acao.get(k, 0) + v
        for k, v in res[cod]["por_nd"].items(): nd[k] = nd.get(k, 0) + v
        nd_nome.update(res[cod]["nd_nome"])
    nd_lbl = {(f'{k} {nd_nome.get(k,"")[:18]}').strip(): v for k, v in nd.items()}
    ch_acao = svg_diverg(acao.items(), "Crédito Disponível por Ação (R$)")
    ch_nd = svg_diverg(nd_lbl.items(), "Crédito Disponível por Natureza de Despesa (R$)")
    funis = "".join(svg_funil(cod, res[cod]) for cod, _ in ALVOS)
    trend = svg_tendencia(hist)

    lin_idx = {}
    for cod, _ in ALVOS:
        for L in res[cod]["linhas"]:
            lin_idx.setdefault((cod, L["acao"], L["pi"], L["nd"]), []).append(L)
    celdata = {}
    cid_map = {}
    cid_seq = [0]
    def cid_for(uasg, c):
        key = (uasg, c["acao"], c["pi"], c["nd"])
        cid = cid_map.get(key)
        if cid is None:
            cid_seq[0] += 1
            cid = "%s_c%d" % (sfx, cid_seq[0])
            cid_map[key] = cid
            ncs = sorted(([L["nc"], L["op"], round(L["cred"], 2), L["obj"], L.get("emit", ""), L.get("dia", "")]
                          for L in lin_idx.get(key, [])), key=lambda x: -x[2])
            celdata[cid] = {"t": f'{c["acao"]} · PI {c["pi"]} · ND {c["nd"]}', "nome": c.get("nd_nome", ""),
                            "u": FONTE_CURTA.get(uasg, ""), "uasg": uasg,
                            "acao": c["acao"], "pi": c["pi"], "pinome": c.get("pi_nome", ""),
                            "nd": c["nd"], "ndnome": c.get("nd_nome", ""),
                            "r": round(c["aloc"], 2), "e": round(c["emp"], 2), "l": round(c.get("liq", 0.0), 2),
                            "p": round(c.get("pag", 0.0), 2), "d": round(c["cred"], 2), "ncs": ncs}
        c["cid"] = cid
        return cid
    def celulas_pos(cod):
        cl = [c for c in res[cod]["celulas"].values() if c["cred"] > 0.005]
        cl.sort(key=lambda x: x["cred"], reverse=True)
        for c in cl:
            cid_for(cod, c)
        return cl
    cons_cel = []
    for cod, _ in ALVOS:
        for c in celulas_pos(cod):
            cons_cel.append({**c, "uasg": cod, "cid": cid_for(cod, c)})
    cons_cel.sort(key=lambda x: x["cred"], reverse=True)
    ogu_c, fex_c = ALVOS[0][0], ALVOS[1][0]
    abas = (f'<button class="tab on" role="tab" aria-selected="true" tabindex="0" onclick="bcmsTab(this,\'tab-cons-{sfx}\')" onkeydown="bcmsTabKey(event,this)">Consolidado</button>'
            f'<button class="tab" role="tab" aria-selected="false" tabindex="-1" onclick="bcmsTab(this,\'tab-{ogu_c}\')" onkeydown="bcmsTabKey(event,this)">{ogu_c} · OGU</button>'
            f'<button class="tab" role="tab" aria-selected="false" tabindex="-1" onclick="bcmsTab(this,\'tab-{fex_c}\')" onkeydown="bcmsTabKey(event,this)">{fex_c} · FEx</button>')
    tabs = (tabela_html(f"tab-cons-{sfx}", cons_cel, True, True) +
            tabela_html(f"tab-{ogu_c}", celulas_pos(ogu_c), False, False) +
            tabela_html(f"tab-{fex_c}", celulas_pos(fex_c), False, False))

    movs = []
    for cod, _ in ALVOS:
        for L in res[cod]["linhas"]:
            if not L["nc"]:
                continue
            try:
                dd, mm, yy = L.get("dia", "").split("/")
                dt = datetime.date(int(yy), int(mm), int(dd))
            except Exception:
                dt = None
            movs.append({**L, "uasg": cod, "dt": dt})
    ncdata = {}
    for i, m in enumerate(movs, 1):
        nid = "%s_n%d" % (sfx, i)
        m["nid"] = nid
        ncdata[nid] = {"nc": m["nc"], "u": FONTE_CURTA.get(m["uasg"], ""), "acao": m["acao"],
                       "pi": m["pi"], "nd": m["nd"], "ndn": m.get("nd_desc", ""), "op": m["op"],
                       "dia": m.get("dia", ""), "val": round(m["cred"], 2), "obj": m["obj"]}
    datas = sorted({m["dt"] for m in movs if m["dt"]}, reverse=True)
    max_date = datas[0] if datas else None
    fmt_d = lambda d: d.strftime("%d/%m/%Y") if d else "—"
    daily = sorted([m for m in movs if m["dt"] == max_date] if max_date else [], key=lambda x: x["cred"], reverse=True)
    rec_d = sum(m["cred"] for m in daily if m["cred"] > 0)
    red_d = sum(m["cred"] for m in daily if m["cred"] < 0)

    def _th(h, numc, sortable=True):
        cls = ' class="num"' if numc else ''
        if not sortable:
            return f'<th{cls}>{esc(h)}</th>'
        return (f'<th{cls} tabindex="0" role="button" aria-sort="none" onclick="bcmsSort(this)" '
                f'onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsSort(this)}}">{esc(h)}<span class="sort"></span></th>')

    def mov_row(m):
        neg = ' cell-neg' if m["cred"] < 0 else ''
        return (f'<tr class="cel-row" tabindex="0" role="button" data-nc="{esc(m["nid"])}" title="Detalhar a NC" '
                f'onclick="bcmsNC(this)" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsNC(this)}}">'
                f'<td class="mono2">{esc(m["nc"])}</td><td><span class="pill-fonte">{esc(FONTE_CURTA[m["uasg"]])}</span></td>'
                f'<td>{esc(m["acao"])}</td><td class="mono2">{esc(m["nd"])}</td>'
                f'<td class="obj" title="{esc(m["op"])}">{esc(m["op"][:28])}</td>'
                f'<td class="num anchor{neg}" data-sort="{m["cred"]:.2f}">{esc(brl(m["cred"]))}<i class="chev" aria-hidden="true">›</i></td></tr>')

    def mov_tabela(tid, lst):
        if not lst:
            return '<p class="vazio">Sem movimentação de NC neste período.</p>'
        ths = _th("NC", False) + _th("Fonte", False) + _th("Ação", False) + _th("ND", False) + _th("Operação", False) + _th("Valor", True)
        body = "".join(mov_row(m) for m in lst)
        return (f'<div class="tbl-tools"><label class="visually-hidden" for="q-{tid}">Buscar</label>'
                f'<input type="search" id="q-{tid}" class="tbl-search" placeholder="Buscar por NC, ação, ND ou operação…" oninput="bcmsSearch(this,\'{tid}\')">'
                f'<button type="button" class="btn-excel" onclick="bcmsExportTable(this,\'{tid}\',\'movimentacao_{tid}\')" title="Baixar movimentação em Excel"><span class="btn-excel-ic">📊</span> Exportar Excel</button>'
                f'<span class="tbl-count" id="cnt-{tid}" data-unit="NC(s)" aria-live="polite">{len(lst)} NC(s)</span></div>'
                f'<div class="tbl-scroll" id="{tid}"><table class="det"><thead><tr>{ths}</tr></thead><tbody>{body}</tbody></table></div>')

    week_days = datas[:7]
    week_rows, tot_rec, tot_red, n_week = [], 0.0, 0.0, 0
    daydata = {}
    for idx, d in enumerate(week_days):
        dm = [m for m in movs if m["dt"] == d]
        rec = sum(m["cred"] for m in dm if m["cred"] > 0)
        red = sum(m["cred"] for m in dm if m["cred"] < 0)
        tot_rec += rec; tot_red += red; n_week += len(dm)
        dk = "%s_d%d" % (sfx, idx)
        daydata[dk] = {"d": fmt_d(d), "n": len(dm), "rec": round(rec, 2), "red": round(red, 2), "liq": round(rec + red, 2),
                       "ncs": sorted([[m["nc"], FONTE_CURTA[m["uasg"]], m["op"], round(m["cred"], 2), m["obj"]] for m in dm],
                                     key=lambda x: -x[3])}
        week_rows.append((d, len(dm), rec, red, rec + red, dk))

    def semana_tabela(rows):
        if not rows:
            return '<p class="vazio">Sem movimentação na última semana.</p>'
        heads = _th("Dia", False, False) + _th("Nº NC", True, False) + _th("Recebido (+)", True, False) + _th("Reduções (−)", True, False) + _th("Líquido", True, False)
        body = ""
        for d, n, rec, red, liq, dk in rows:
            body += (f'<tr class="cel-row" tabindex="0" role="button" data-day="{dk}" title="Ver as NCs deste dia" '
                     f'onclick="bcmsDay(this)" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsDay(this)}}">'
                     f'<td class="mono2">{fmt_d(d)}</td><td class="num">{n}</td>'
                     f'<td class="num col-pos">{esc(brl(rec))}</td><td class="num col-neg">{esc(brl(red))}</td>'
                     f'<td class="num anchor">{esc(brl(liq))}<i class="chev" aria-hidden="true">›</i></td></tr>')
        foot = (f'<tfoot><tr><td>TOTAL</td><td class="num">{n_week}</td>'
                f'<td class="num">{esc(brl(tot_rec))}</td><td class="num">{esc(brl(tot_red))}</td>'
                f'<td class="num anchor">{esc(brl(tot_rec + tot_red))}</td></tr></tfoot>')
        return f'<div class="tbl-scroll"><table class="det"><thead><tr>{heads}</tr></thead><tbody>{body}</tbody>{foot}</table></div>'

    asof = max_date or datetime.date.today()
    rec_por_cel = {}
    for cod, _ in ALVOS:
        for L in res[cod]["linhas"]:
            if L["cred"] > 0 and L.get("dia"):
                try:
                    dd, mm, yy = L["dia"].split("/"); dt = datetime.date(int(yy), int(mm), int(dd))
                except Exception:
                    dt = None
                if dt:
                    rec_por_cel.setdefault((cod, L["acao"], L["pi"], L["nd"]), []).append({**L, "dt": dt})

    emtela_ncs = []
    for cod, _ in ALVOS:
        for cl in celulas_pos(cod):
            recs = sorted(rec_por_cel.get((cod, cl["acao"], cl["pi"], cl["nd"]), []),
                          key=lambda L: L["dt"], reverse=True)
            restante = cl["cred"]
            for L in recs:
                if restante <= 0.005:
                    break
                val_nc = min(L["cred"], restante)
                restante -= val_nc
                dt = L["dt"]
                dias = (asof - dt).days if dt else None
                emtela_ncs.append({
                    "uasg": cod,
                    "fonte": FONTE_CURTA.get(cod, ""),
                    "nc": L.get("nc") or "(sem número)",
                    "acao": cl["acao"],
                    "pi": cl["pi"],
                    "pi_nome": cl.get("pi_nome", ""),
                    "nd": cl["nd"],
                    "nd_nome": cl.get("nd_nome", ""),
                    "obj": L.get("obj", "") or "(sem descrição)",
                    "op": L.get("op", ""),
                    "emit": L.get("emit", ""),
                    "dia": L.get("dia", ""),
                    "dt": dt,
                    "dias": dias,
                    "cred": val_nc,
                    "cid": cl.get("cid", ""),
                    "nd_de": L.get("nd_de", ""),
                    "nc_origem": L.get("nc_origem"),
                })
            if restante > 0.005:
                emtela_ncs.append({
                    "uasg": cod,
                    "fonte": FONTE_CURTA.get(cod, ""),
                    "nc": "(Saldo em tela)",
                    "acao": cl["acao"],
                    "pi": cl["pi"],
                    "pi_nome": cl.get("pi_nome", ""),
                    "nd": cl["nd"],
                    "nd_nome": cl.get("nd_nome", ""),
                    "obj": f"Saldo remanescente em tela da célula {cl['acao']} · PI {cl['pi']} · ND {cl['nd']}",
                    "op": "SALDO REMANESCENTE",
                    "emit": "",
                    "dia": "",
                    "dt": None,
                    "dias": None,
                    "cred": restante,
                    "cid": cl.get("cid", "")
                })

    emtela_ncs.sort(key=lambda x: (x["dias"] if x["dias"] is not None else -1, x["cred"]), reverse=True)
    # Dados por LINHA em tela (uma NC específica), para o detalhamento não misturar NCs
    # de objetos diferentes que apenas compartilham o mesmo PI.
    teladata = {}
    for _i, _c in enumerate(emtela_ncs, 1):
        _tid = "%s_t%d" % (sfx, _i)
        _c["tid"] = _tid
        _org = _c.get("nc_origem") or {}
        teladata[_tid] = {
            "nc": _c["nc"], "uasg": _c["uasg"], "u": _c.get("fonte", ""),
            "acao": _c["acao"], "pi": _c["pi"], "pinome": _c.get("pi_nome", ""),
            "nd": _c["nd"], "ndnome": _c.get("nd_nome", ""),
            "v": round(_c["cred"], 2), "dias": _c["dias"], "dia": _c.get("dia", ""),
            "obj": _c.get("obj", ""), "op": _c.get("op", ""), "emit": _c.get("emit", ""),
            "cid": _c.get("cid", ""), "nd_de": _c.get("nd_de", ""),
            "org": ({"nc": _org.get("nc", ""), "obj": _org.get("obj", ""), "dia": _org.get("dia", ""),
                     "emit": _org.get("emit", "")} if _org else None),
        }
    tot_emtela = sum(c["cred"] for c in emtela_ncs)
    idades = [c["dias"] for c in emtela_ncs if c["dias"] is not None]
    idade_media = round(sum(idades) / len(idades)) if idades else 0
    idade_max = max(idades) if idades else 0

    def et_row(c):
        aplic = c.get("nd_nome") or c.get("pi_nome") or ""
        acao_nd = f'{c["acao"]} · {c["nd"]}'
        dias = c["dias"]
        if dias is None:
            dcls, dtxt, dsort = "badge-age age-none", "—", -1
        else:
            if dias > 60:
                dcls = "badge-age age-red"
            elif dias > 30:
                dcls = "badge-age age-amber"
            else:
                dcls = "badge-age age-green"
            dtxt = f'{dias}d'
            dsort = dias
        refd = c["dt"].strftime("%d/%m/%y") if c["dt"] else "—"
        cid = esc(c.get("cid", ""))
        # Troca de ND: a descrição própria ("MUDANÇA DE ND", "DETALHAMENTO…") não diz o objeto.
        # Mostra o objeto REAL da NC de origem, sinalizando a troca.
        org = c.get("nc_origem")
        if org and org.get("obj"):
            desc_completa = esc(org["obj"])
            selo = f'<span class="tag-nd" title="Crédito recebido por mudança de ND (de {esc(c.get("nd_de",""))} para {esc(c["nd"])}) — objeto herdado da NC de origem {esc(org["nc"])}">↪ ND {esc(c.get("nd_de",""))}</span> '
        else:
            desc_completa = esc(c.get("obj", ""))
            selo = ""
        desc_resumo = desc_completa[:118] + ("…" if len(desc_completa) > 118 else "")
        desc_resumo = selo + desc_resumo
        # nº curto da NC (o completo fica no title e no modal): "160504…2026NC401667" -> "NC 401667 · 160504"
        nc_full = str(c["nc"] or "")
        m_nc = re.search(r"NC(\d+)$", nc_full)
        if m_nc:
            nc_lbl = f'<span class="nc-lbl-wrap"><span class="nc-num">NC {esc(m_nc.group(1))}</span> <span class="nc-ug">({esc(nc_full[:6])})</span></span>'
        else:
            nc_lbl = f'<span class="nc-num">{esc(nc_full)}</span>'
        tid = esc(c.get("tid", ""))
        faixa = "none" if dias is None else ("r" if dias > 60 else ("a" if dias > 30 else "v"))
        return (f'<tr class="cel-row" tabindex="0" role="button" data-tela="{tid}" data-cel="{cid}" '
                f'data-fonte="{esc(c.get("fonte",""))}" data-acao="{esc(c["acao"])}" data-nd="{esc(c["nd"])}" '
                f'data-faixa="{faixa}" data-val="{c["cred"]:.2f}" title="Clique para ver o detalhamento completo em tela desta NC" '
                f'onclick="bcmsTela(this)" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsTela(this)}}">'
                f'<td><span class="pill-fonte">{esc(c["fonte"])}</span></td>'
                f'<td class="mono2" style="white-space:nowrap;" title="{esc(nc_full)}">{nc_lbl}</td>'
                f'<td class="mono2" style="white-space:nowrap;">{esc(acao_nd)}</td>'
                f'<td class="obj" title="{desc_completa}" data-full-desc="{desc_completa}">{desc_resumo}</td>'
                f'<td class="mono2" style="white-space:nowrap;">{esc(refd)}</td>'
                f'<td class="num anchor" data-sort="{c["cred"]:.2f}" title="Crédito Disponível: {esc(brl(c["cred"]))} (Clique para detalhar)">{esc(brl(c["cred"]))}</td>'
                f'<td class="num" data-sort="{dsort}"><span class="{dcls}">{dtxt}</span><i class="chev" aria-hidden="true">›</i></td></tr>')

    et_ths = (
        _th("Fonte", False) +
        _th("NC", False) +
        _th("Ação · ND", False) +
        _th("Descrição do objeto da NC", False) +
        _th("Recebido em", False) +
        _th("Crédito em Tela", True) +
        _th("Idade", True)
    )

    # opções dos filtros — só o que existe de fato nesta unidade
    _fontes = sorted({c.get("fonte", "") for c in emtela_ncs if c.get("fonte")})
    _acoes = sorted({c["acao"] for c in emtela_ncs if c["acao"]})
    _nds = sorted({c["nd"] for c in emtela_ncs if c["nd"]})
    _nd_lbl = {}
    for _c in emtela_ncs:
        if _c["nd"] and _c["nd"] not in _nd_lbl and _c.get("nd_nome"):
            _nd_lbl[_c["nd"]] = _c["nd_nome"]
    opt_fonte = '<option value="">Fonte: todas</option>' + "".join(
        f'<option value="{esc(_f)}">{esc(_f)}{" · OGU (exercício corrente)" if _f == "160" else (" · FEx (exercícios anteriores)" if _f == "167" else "")}</option>' for _f in _fontes)
    opt_acao = '<option value="">Ação: todas</option>' + "".join(
        f'<option value="{esc(_a)}">{esc(_a)}</option>' for _a in _acoes)
    opt_nd = '<option value="">ND: todas</option>' + "".join(
        f'<option value="{esc(_n)}">{esc(_n)}{esc(" — " + _nd_lbl[_n][:26]) if _nd_lbl.get(_n) else ""}</option>'
        for _n in _nds)
    _cb = "bcmsFiltra('%s')" % sfx
    _cbl = "bcmsLimpaFiltros('%s')" % sfx

    emtela_html = (
        '<section class="sec"><div class="eyebrow">Créditos em tela — por Nota de Crédito (NC)</div>'
        '<div class="et-head">'
        f'<div class="et-kpi et-hero"><span>Crédito Disponível em tela</span><b class="num">{esc(brl(tot_emtela))}</b></div>'
        f'<div class="et-kpi"><span>Notas de Crédito</span><b class="num">{len(emtela_ncs)}</b></div>'
        f'<div class="et-kpi"><span>Idade média</span><b class="num">{idade_media} dias</b></div>'
        f'<div class="et-kpi"><span>Mais antigo</span><b class="num">{idade_max} dias</b></div>'
        f'<div class="et-action"><button type="button" class="btn-excel btn-excel-lg" onclick="bcmsExportTable(this,\'tab-emtela-{sfx}\',\'creditos_em_tela_nc_{sfx}\')" title="Baixar relatório detalhado de créditos por NC em planilha Excel"><span class="btn-excel-ic">📥</span> Baixar Relatório NC em Excel</button></div>'
        f'<div class="et-meta">Posição {esc(posicao)}<br><span class="rh-delay">⏱ dados com ~24h de defasagem</span></div></div>'
        '<p class="sec-nota">Relação dos <b>créditos disponíveis por Nota de Crédito (NC)</b> com descrição completa do objeto e <b>dias em tela</b> (desde o lançamento da NC). '
        '<b>Clique em uma linha</b> para abrir a ficha completa. Legenda de idade: <span class="badge-age age-green">≤30d</span> recente · <span class="badge-age age-amber">31–60d</span> atenção · <span class="badge-age age-red">&gt;60d</span> crítico.</p>'
        f'<div class="tbl-tools"><label class="visually-hidden" for="q-tab-emtela-{sfx}">Buscar</label>'
        f'<input type="search" id="q-tab-emtela-{sfx}" class="tbl-search" placeholder="Buscar por NC, Ação, ND ou palavras na descrição completa…" oninput="{_cb}">'
        f'<button type="button" class="btn-excel" onclick="bcmsExportTable(this,\'tab-emtela-{sfx}\',\'creditos_em_tela_nc_{sfx}\')" title="Exporta exatamente as linhas visíveis, conforme os filtros aplicados"><span class="btn-excel-ic">📊</span> Exportar Excel</button>'
        f'<span class="tbl-count" id="cnt-tab-emtela-{sfx}" data-unit="NC(s) em tela" aria-live="polite">{len(emtela_ncs)} NC(s) em tela</span></div>'
        f'<div class="tbl-filtros" id="flt-{sfx}" role="group" aria-label="Filtros da lista de créditos em tela">'
        '<span class="flt-lbl">Filtrar:</span>'
        f'<select class="flt" id="f-fonte-{sfx}" aria-label="Filtrar por fonte" onchange="{_cb}">{opt_fonte}</select>'
        f'<select class="flt" id="f-acao-{sfx}" aria-label="Filtrar por ação de governo" onchange="{_cb}">{opt_acao}</select>'
        f'<select class="flt" id="f-nd-{sfx}" aria-label="Filtrar por natureza de despesa" onchange="{_cb}">{opt_nd}</select>'
        f'<select class="flt" id="f-idade-{sfx}" aria-label="Filtrar por idade em tela" onchange="{_cb}">'
        '<option value="">Idade: todas</option><option value="v">≤ 30 dias</option>'
        '<option value="a">31 a 60 dias</option><option value="r">&gt; 60 dias</option></select>'
        f'<button type="button" class="flt-limpa" onclick="{_cbl}" title="Limpar todos os filtros">✕ Limpar</button>'
        f'<span class="flt-resumo" id="flt-res-{sfx}" aria-live="polite"></span></div>'
        f'<div class="tbl-scroll" id="tab-emtela-{sfx}"><table class="det det-compact"><thead><tr>{et_ths}</tr></thead>'
        f'<tbody>{"".join(et_row(c) for c in emtela_ncs)}</tbody>'
        f'<tfoot><tr><td colspan="5">TOTAL · {len(emtela_ncs)} Nota(s) de Crédito em tela</td><td class="num anchor">{esc(brl(tot_emtela))}</td><td>—</td></tr></tfoot></table></div></section>'
    )

    resumo_html = (
        emtela_html
        + f'<section class="sec"><div class="eyebrow">Movimentação de NC — {fmt_d(max_date)} (dia anterior)</div>'
        f'<p class="sec-nota">Notas de crédito com lançamento em <b>{fmt_d(max_date)}</b> (último dia com movimento — dados com ~24h de defasagem): '
        f'<b>{len(daily)}</b> NC(s) · Recebido <b>{esc(brl(rec_d))}</b> · Reduções <b>{esc(brl(red_d))}</b> · Líquido <b>{esc(brl(rec_d + red_d))}</b>. '
        'Clique em uma NC para detalhá-la.</p>'
        + mov_tabela(f"mov-dia-{sfx}", daily) + '</section>'
        '<section class="sec"><div class="eyebrow">Resumo semanal — últimos 7 dias com movimentação</div>'
        '<p class="sec-nota">Movimentação de NC por dia: recebimentos (+), reduções/anulações (−) e líquido. <b>Clique em um dia</b> para ver as NCs daquele dia.</p>'
        + semana_tabela(week_rows) + '</section>'
    )

    hero_eq = (
        f'<div class="hero-eq-box"><span class="eq-tag">RECEBIDO</span><span class="eq-val num">{esc(brl(tot["prov"]))}</span></div>'
        f'<span class="hero-eq-sign">−</span>'
        f'<div class="hero-eq-box"><span class="eq-tag">EMPENHADO</span><span class="eq-val num eq-emp">{esc(brl(tot["emp"]))}</span></div>'
        f'<span class="hero-eq-sign">=</span>'
        f'<div class="hero-eq-box eq-highlight"><span class="eq-tag">DISPONÍVEL</span><span class="eq-val num eq-disp">{esc(brl(tot["cred"]))}</span></div>'
    )
    # ---- Seção Histórico Completo da Unidade ----
    tot_u_ncs = len(set(x["nc"] for x in u_hist_items))
    tot_u_prov = sum(x["prov"] for x in u_hist_items)
    tot_u_emp  = sum(x["emp"]  for x in u_hist_items)
    tot_u_cred = sum(x["cred"] for x in u_hist_items)

    u_ptres_set = set(x["ptres"] for x in u_hist_items if x.get("ptres"))
    u_meses_set = set(x["mes"] for x in u_hist_items if x.get("mes"))
    opt_u_ptres = '<option value="">Ação: todas</option>' + "".join(
        f'<option value="{esc(p)}">{esc(p)}</option>' for p in sorted(u_ptres_set)
    )
    opt_u_meses = (
        '<option value="">Período: todos</option>'
        '<optgroup label="Trimestres">'
        '<option value="T1">1º Trimestre (Jan–Mar)</option>'
        '<option value="T2">2º Trimestre (Abr–Jun)</option>'
        '<option value="T3">3º Trimestre (Jul–Set)</option>'
        '<option value="T4">4º Trimestre (Out–Dez)</option>'
        '</optgroup>'
        '<optgroup label="Meses">'
        + "".join(f'<option value="{esc(m)}">{esc(m)}</option>' for m in sorted(u_meses_set, reverse=True))
        + '</optgroup>'
    )

    u_sigla = esc(u['sigla'])
    u_nome = esc(u['nome'])
    u_ogu = esc(str(u['ogu']))
    u_fex = esc(str(u['fex']))

    u_head_html = (
        f'<div class="et-head">'
        f'<div class="et-kpi et-hero"><span>Saldo Disponível</span><b class="num" id="kpi-uhist-cred-{sfx}">{esc(brl(tot_u_cred))}</b></div>'
        f'<div class="et-kpi"><span>Notas de Crédito</span><b class="num" id="kpi-uhist-total-{sfx}">{tot_u_ncs}</b></div>'
        f'<div class="et-kpi"><span>Provisão Recebida</span><b class="num" id="kpi-uhist-prov-{sfx}">{esc(brl(tot_u_prov))}</b></div>'
        f'<div class="et-kpi"><span>Total Empenhado</span><b class="num" id="kpi-uhist-emp-{sfx}">{esc(brl(tot_u_emp))}</b></div>'
        f'<div class="et-action"><button type="button" class="btn-excel btn-excel-lg" onclick="bcmsExportUHistExcel(\'{sfx}\',\'{u_sigla}\')" title="Baixar histórico completo de {u_sigla} em planilha Excel"><span class="btn-excel-ic">📥</span> Baixar Histórico em Excel</button></div>'
        f'<div class="et-meta">Posição {esc(posicao)}<br><span class="rh-delay">⏱ dados com ~24h de defasagem</span></div>'
        f'</div>'
    )

    u_tools_html = (
        f'<div class="tbl-tools">'
        f'<label class="visually-hidden" for="flt-uhist-busca-{sfx}">Buscar</label>'
        f'<input type="search" id="flt-uhist-busca-{sfx}" class="tbl-search" placeholder="Buscar por NC, Ação, ND, PI ou palavras na descrição completa…" oninput="bcmsFiltraUHist(\'{sfx}\')">'
        f'<button type="button" class="btn-excel" onclick="bcmsExportUHistExcel(\'{sfx}\',\'{u_sigla}\')" title="Exportar exatamente as linhas visíveis do histórico, conforme os filtros aplicados"><span class="btn-excel-ic">📊</span> Exportar Excel</button>'
        f'<span class="tbl-count" id="cnt-uhist-ncs-{sfx}" data-unit="NC(s) no histórico" aria-live="polite">{len(u_hist_items)} NC(s) no histórico</span>'
        f'</div>'
        f'<div class="tbl-filtros" id="flt-uhist-{sfx}" role="group" aria-label="Filtros do histórico de {u_sigla}">'
        f'<span class="flt-lbl">Filtrar:</span>'
        f'<select class="flt" id="flt-uhist-fonte-{sfx}" aria-label="Filtrar por fonte" onchange="bcmsFiltraUHist(\'{sfx}\')">'
        f'<option value="">Fonte: todas</option>'
        f'<option value="OGU">OGU (160 - Orçamento Geral da União)</option>'
        f'<option value="FEx">FEx (167 - Fundo do Exército)</option>'
        f'</select>'
        f'<select class="flt" id="flt-uhist-ptres-{sfx}" aria-label="Filtrar por ação de governo" onchange="bcmsFiltraUHist(\'{sfx}\')">{opt_u_ptres}</select>'
        f'<select class="flt" id="flt-uhist-periodo-{sfx}" aria-label="Filtrar por período" onchange="bcmsFiltraUHist(\'{sfx}\')">{opt_u_meses}</select>'
        f'<select class="flt" id="flt-uhist-faixa-{sfx}" aria-label="Filtrar por status do saldo" onchange="bcmsFiltraUHist(\'{sfx}\')">'
        f'<option value="">Status: todos</option>'
        f'<option value="saldo_pos">🟢 Com Saldo (&gt; R$ 0)</option>'
        f'<option value="parcial">🟡 Parcialmente Executadas</option>'
        f'<option value="zerada">⚪ Executadas / Zeradas</option>'
        f'<option value="canc">🔴 Anuladas / Canceladas</option>'
        f'</select>'
        f'<button type="button" class="flt-limpa" onclick="bcmsLimpaFiltrosUHist(\'{sfx}\')" title="Limpar todos os filtros">✕ Limpar</button>'
        f'<span class="flt-resumo" id="flt-uhist-res-{sfx}" aria-live="polite"></span>'
        f'</div>'
    )

    u_initial_rows = []
    for item in u_hist_items[:50]:
        status_cls = f"status-{item['status_slug']}"
        emit_nome_curto = (item["emit_nome"][:22] + "…") if len(item["emit_nome"]) > 22 else item["emit_nome"]
        desc_completa = item.get("obj", "")
        desc_resumo = desc_completa[:118] + ("…" if len(desc_completa) > 118 else "")
        nc_full = str(item["nc"] or "")
        m_nc = re.search(r"NC(\d+)$", nc_full)
        if m_nc:
            nc_lbl = f'<span class="nc-num">NC {esc(m_nc.group(1))}</span> <span class="nc-ug">· {esc(nc_full[:6])}</span>'
        else:
            nc_lbl = f'<span class="nc-num">{esc(nc_full)}</span>'
        acao_nd = f'{item["ptres"]} · {item["nd"]}' if item.get("ptres") and item.get("nd") else (item.get("ptres") or item.get("nd") or "—")

        u_initial_rows.append(
            f'<tr class="cel-row" data-hid="{item["hid"]}" tabindex="0" role="button" onclick="bcmsDetalheNC(\'{item["hid"]}\')" '
            f'title="Clique para abrir o detalhamento completo da NC {esc(nc_full)}" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsDetalheNC(\'{item["hid"]}\')}}">'
            f'<td><span class="pill-fonte">{esc(item["fonte"])}</span></td>'
            f'<td class="mono2" title="{esc(nc_full)}">{nc_lbl}</td>'
            f'<td class="mono2">{esc(acao_nd)}</td>'
            f'<td class="obj" title="{esc(desc_completa)}" data-full-desc="{esc(desc_completa)}">{esc(desc_resumo)}</td>'
            f'<td title="{esc(item["emit_nome"])}"><span class="ug-pill emit">{esc(item["emit_cod"])}</span> <small>{esc(emit_nome_curto)}</small></td>'
            f'<td class="mono2">{esc(item["dia"] or "—")}</td>'
            f'<td class="num" data-sort="{item["prov"]:.2f}">{esc(brl(item["prov"]))}</td>'
            f'<td class="num anchor" data-sort="{item["cred"]:.2f}">{esc(brl(item["cred"]))}</td>'
            f'<td class="num"><span class="pill-status {status_cls}">{esc(item["status"])}</span><i class="chev" aria-hidden="true">›</i></td>'
            f'</tr>'
        )

    u_pages = math.ceil(len(u_hist_items) / 50) if u_hist_items else 1
    u_table_html = (
        f'<div class="tbl-scroll" id="scroll-uhist-ncs-{sfx}">'
        f'<table class="det det-compact" id="tbl-uhist-ncs-{sfx}">'
        f'<thead>'
        f'<tr>'
        f'<th tabindex="0" role="button" aria-sort="none" onclick="bcmsSortUHist(\'fonte\', this, \'{sfx}\')" title="Ordenar por Fonte">Fonte <span class="sort"></span></th>'
        f'<th tabindex="0" role="button" aria-sort="none" onclick="bcmsSortUHist(\'nc\', this, \'{sfx}\')" title="Ordenar por Número da NC">NC <span class="sort"></span></th>'
        f'<th tabindex="0" role="button" aria-sort="none" onclick="bcmsSortUHist(\'ptres\', this, \'{sfx}\')" title="Ordenar por PTRES / Ação">Ação · ND <span class="sort"></span></th>'
        f'<th>Descrição do objeto da NC</th>'
        f'<th tabindex="0" role="button" aria-sort="none" onclick="bcmsSortUHist(\'emit\', this, \'{sfx}\')" title="Ordenar por UG Emitente">UG Emitente <span class="sort"></span></th>'
        f'<th tabindex="0" role="button" aria-sort="descending" onclick="bcmsSortUHist(\'dt\', this, \'{sfx}\')" title="Ordenar por Data">Recebido em <span class="sort">▼</span></th>'
        f'<th class="num" tabindex="0" role="button" aria-sort="none" onclick="bcmsSortUHist(\'prov\', this, \'{sfx}\')" title="Ordenar por Provisão Recebida">Provisão <span class="sort"></span></th>'
        f'<th class="num" tabindex="0" role="button" aria-sort="none" onclick="bcmsSortUHist(\'cred\', this, \'{sfx}\')" title="Ordenar por Saldo Disponível">Crédito Disp. <span class="sort"></span></th>'
        f'<th tabindex="0" role="button" aria-sort="none" onclick="bcmsSortUHist(\'status\', this, \'{sfx}\')" title="Ordenar por Status">Status <span class="sort"></span></th>'
        f'</tr>'
        f'</thead>'
        f'<tbody id="tbody-uhist-ncs-{sfx}">'
        f'{"".join(u_initial_rows)}'
        f'</tbody>'
        f'<tfoot>'
        f'<tr>'
        f'<td colspan="6" id="tf-uhist-label-{sfx}">TOTAL · {len(u_hist_items)} Nota(s) de Crédito no histórico</td>'
        f'<td class="num" id="tf-uhist-prov-{sfx}">{esc(brl(tot_u_prov))}</td>'
        f'<td class="num anchor" id="tf-uhist-cred-{sfx}">{esc(brl(tot_u_cred))}</td>'
        f'<td>—</td>'
        f'</tr>'
        f'</tfoot>'
        f'</table>'
        f'</div>'
        f'<div class="tbl-tools" style="margin-top:10px;justify-content:space-between;" id="paginacao-uhist-{sfx}">'
        f'<span class="pag-info" id="pag-info-uhist-{sfx}" style="font-size:0.8125rem;color:var(--ink-muted);font-weight:600;">'
        f'Página 1 de {u_pages} (Exibindo 1–{min(50, len(u_hist_items))} de {len(u_hist_items)} NCs)'
        f'</span>'
        f'<div style="display:flex;gap:8px;align-items:center;">'
        f'<button type="button" class="flt-limpa" id="btn-pag-uhist-ant-{sfx}" onclick="bcmsPaginaUHist(-1, \'{sfx}\')" disabled>‹ Anterior</button>'
        f'<button type="button" class="flt-limpa" id="btn-pag-uhist-prox-{sfx}" onclick="bcmsPaginaUHist(1, \'{sfx}\')"' + (' disabled' if u_pages <= 1 else '') + '>Próxima ›</button>'
        f'</div>'
        f'</div>'
    )

    historico_unidade_html = (
        f'<section class="sec">'
        f'<div class="eyebrow">Histórico Completo de Notas de Crédito — {u_sigla}</div>'
        f'{u_head_html}'
        f'<p class="sec-nota">Relação completa de todas as <b>Notas de Crédito (NC) recebidas</b> pela unidade <b>{u_nome} ({u_sigla})</b> no exercício corrente (UASGs {u_ogu} · OGU e {u_fex} · FEx). <b>Clique em uma linha</b> para abrir a ficha cadastral completa no modal de detalhamento.</p>'
        f'{u_tools_html}'
        f'{u_table_html}'
        f'</section>'
    )

    disp = "" if (u is UNIDADES[0]) else ' style="display:none"'
    frag = f"""<section class="unidade" id="secao-{sfx}" data-key="{sfx}" data-sigla="{esc(u['sigla'])}"{disp}>
  <div class="toptabs" role="tablist" aria-label="Visões do painel">
    <button class="toptab on" role="tab" aria-selected="true" onclick="bcmsView(this,'resumo')">📋 Resumo Executivo & Créditos em Tela</button>
    <button class="toptab" role="tab" aria-selected="false" onclick="bcmsView(this,'completo')">📊 Detalhamento Completo & Gráficos</button>
    <button class="toptab" role="tab" aria-selected="false" onclick="bcmsView(this,'historico')">📜 Histórico Completo</button>
  </div>
  <div class="view-resumo">
  {resumo_html}
  </div>
  <div class="view-completo" style="display:none">
  <section class="hero">
    <div class="hero-l">
      <div class="eyebrow">Crédito Disponível · Consolidado {esc(u['sigla'])}</div>
      <div class="hero-num num">{esc(brl(tot["cred"]))}</div>
      <div class="hero-eq">{hero_eq}</div>
    </div>
    <div class="hero-r">
      {delta_html}
      {svg_util(tot["prov"], tot["emp"], tot["cred"])}
    </div>
  </section>

  <section class="sec">
    <div class="eyebrow">Composição Visual da Disponibilidade</div>
    <div class="card">{svg_waterfall(tot["prov"], tot["emp"], tot["cred"])}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Indicadores Globais de Execução</div>
    <div class="kpis">{kpis}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Desdobramento por Fonte de Recurso</div>
    <div class="grid2">{cards_uasg}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Distribuição do Crédito Disponível</div>
    <div class="grid2">{ch_acao}{ch_nd}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Estágios da Despesa por Fonte</div>
    <div class="grid2">{funis}</div>
  </section>

  <section class="sec">{trend}</section>

  <section class="sec">
    <div class="eyebrow">Crédito Disponível em tela — por célula orçamentária</div>
    <p class="sec-nota">Saldo <b>líquido</b> por célula (Ação · PI · ND): <b>Recebido (líq) − Empenhado = Crédito Disponível</b> em cada linha. "Recebido (líq)" já compensa alterações de ND, detalhamentos e anulações (a alteração <b>não é somada</b> com a NC original). Só aparecem células com saldo &gt; 0; a soma fecha com o total consolidado. <b>Clique em uma linha</b> para ver as notas de crédito da célula com a descrição completa.</p>
    <div class="tabs" role="tablist" aria-label="Crédito em tela por UASG">{abas}</div>
    {tabs}
  </section>
  </div>
  <div class="view-historico" style="display:none">
  {historico_unidade_html}
  </div>
</section>"""
    return frag, celdata, ncdata, daydata, teladata

# ---------------- módulo Ranking e Comparativo OMDS ----------------
def svg_comparativo_barras(u_stats, metric="cred", titulo="Crédito Disponível por Unidade (R$)"):
    itens = [(u["sigla"], u[metric], u["accent"]) for u in u_stats if round(u[metric], 2) != 0]
    itens.sort(key=lambda x: x[1], reverse=True)
    if not itens:
        return f'<div class="card chart"><div class="eyebrow">{esc(titulo)}</div><p class="vazio">Sem valores no período</p></div>'
    vmax = max(x[1] for x in itens) or 1
    labW, rh, pad = 110, 36, 12
    zx = labW + 10
    plotW = 440
    W = zx + plotW + 140
    H = pad * 2 + rh * len(itens)
    el = []
    for i, (sigla, val, accent) in enumerate(itens):
        y = pad + i * rh
        w = max(4, (val / vmax) * plotW)
        el.append(f'<text x="{labW-8}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-cat" style="font-weight:700">{esc(sigla)}</text>')
        el.append(f'<rect x="{zx}" y="{y+4}" width="{w:.1f}" height="{rh-8}" rx="4" style="fill:{accent}"><title>{esc(sigla)}: {esc(brl(val))}</title></rect>')
        el.append(f'<text x="{zx+w+8:.1f}" y="{y+rh/2+4:.0f}" class="s-num s-on2" style="font-weight:700">{esc(brl(val))}</text>')
    svg = f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="{esc(titulo)}">{"".join(el)}</svg>'
    return f'<div class="card chart"><div class="eyebrow">{esc(titulo)}</div>{svg}</div>'

def svg_comparativo_exec(u_stats, media_cmd):
    itens = [(u["sigla"], u["exec_pct"], u["accent"], u["prov"], u["emp"]) for u in u_stats]
    itens.sort(key=lambda x: x[1], reverse=True)
    labW, rh, pad = 110, 36, 12
    zx = labW + 10
    plotW = 440
    W = zx + plotW + 90
    H = pad * 2 + rh * len(itens) + 24
    el = []
    x_media = zx + (min(100.0, media_cmd) / 100.0) * plotW
    el.append(f'<line x1="{x_media:.1f}" y1="{pad-4}" x2="{x_media:.1f}" y2="{H-pad-16}" stroke="var(--gold)" stroke-width="2" stroke-dasharray="4 3"/>')
    el.append(f'<text x="{x_media:.1f}" y="{H-pad-2}" text-anchor="middle" font-size="11" font-weight="700" fill="var(--gold)">Média do Comando: {media_cmd:.1f}%</text>')
    for i, (sigla, pct_v, accent, prov, emp) in enumerate(itens):
        y = pad + i * rh
        w = max(4, (min(100.0, pct_v) / 100.0) * plotW)
        el.append(f'<text x="{labW-8}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-cat" style="font-weight:700">{esc(sigla)}</text>')
        el.append(f'<rect x="{zx}" y="{y+4}" width="{plotW}" height="{rh-8}" rx="4" fill="var(--track)"/>')
        el.append(f'<rect x="{zx}" y="{y+4}" width="{w:.1f}" height="{rh-8}" rx="4" style="fill:{accent}"><title>{esc(sigla)}: {pct_v:.1f}% empenhado ({esc(brl(emp))} de {esc(brl(prov))})</title></rect>')
        el.append(f'<text x="{zx+w+8:.1f}" y="{y+rh/2+4:.0f}" class="s-num" font-weight="700">{pct_v:.1f}%</text>')
    svg = f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="Taxa de Execução Orçamentária por OMDS">{"".join(el)}</svg>'
    return f'<div class="card chart"><div class="eyebrow">Taxa de Execução Orçamentária (% Empenhado / Recebido)</div>{svg}</div>'

def secao_comparativo_omds(omds_totais, hist, data_str, periodo):
    u_stats = []
    for u in OMDS_COMPARATIVO:
        ot = omds_totais.get(u["key"], {})
        tot_prov = ot.get("prov", 0.0)
        tot_conc = ot.get("conc", 0.0)
        tot_emp  = ot.get("emp", 0.0)
        tot_liq  = ot.get("liq", 0.0)
        tot_pag  = ot.get("pag", 0.0)
        tot_cred = ot.get("cred", 0.0)
        n_cel    = ot.get("n_ncs_cnt", ot.get("n_ncs", 0))
        if isinstance(n_cel, (set, list)): n_cel = len(n_cel)
        exec_pct = pct(tot_emp, tot_prov)
        liq_pct  = pct(tot_liq, tot_emp)
        pag_pct  = pct(tot_pag, tot_liq)
        u_stats.append({
            "key": u["key"], "sigla": u["sigla"], "nome": u["nome"], "logo": u["logo"],
            "accent": u["accent"], "ogu": u["ogu"], "fex": u["fex"],
            "prov": tot_prov, "conc": tot_conc, "emp": tot_emp, "liq": tot_liq,
            "pag": tot_pag, "cred": tot_cred, "exec_pct": exec_pct,
            "liq_pct": liq_pct, "pag_pct": pag_pct, "n_cel": n_cel
        })
    
    cmd_prov = sum(x["prov"] for x in u_stats)
    cmd_emp = sum(x["emp"] for x in u_stats)
    cmd_liq = sum(x["liq"] for x in u_stats)
    cmd_pag = sum(x["pag"] for x in u_stats)
    cmd_cred = sum(x["cred"] for x in u_stats)
    cmd_n_cel = sum(x["n_cel"] for x in u_stats)
    cmd_exec_pct = pct(cmd_emp, cmd_prov)
    cmd_liq_pct = pct(cmd_liq, cmd_emp)
    cmd_pag_pct = pct(cmd_pag, cmd_liq)
    
    rank_exec = sorted(u_stats, key=lambda x: x["exec_pct"], reverse=True)
    
    podio_order = []
    if len(rank_exec) >= 2:
        podio_order.append((rank_exec[1], 2, "🥈 2º Lugar", "silver"))
    if len(rank_exec) >= 1:
        podio_order.append((rank_exec[0], 1, "🥇 1º Lugar", "gold"))
    if len(rank_exec) >= 3:
        podio_order.append((rank_exec[2], 3, "🥉 3º Lugar", "bronze"))
    
    podio_cards = []
    for u, pos, badge, cls in podio_order:
        podio_cards.append(
            f'<div class="podium-step podium-{cls}" onclick="' + (f'trocaOMDSPorKey(\'{u["key"]}\')' if u["key"] == "BEC6" else f'bcmsDetalheOMDS(\'{u["key"]}\')') + f'" title="Clique para abrir o detalhamento completo de {esc(u["sigla"])}">'
            f'<div class="podium-badge">{badge}</div>'
            f'<div class="podium-avatar-wrap"><img src="assets/logos/{u["logo"]}" alt="{esc(u["sigla"])}" class="podium-logo" onerror="this.src=\'assets/logos/12RM.png\'"></div>'
            f'<div class="podium-sigla">{esc(u["sigla"])}</div>'
            f'<div class="podium-nome">{esc(u["nome"])}</div>'
            f'<div class="podium-stat-pill"><span class="stat-l">Execução</span><b class="stat-v num">{u["exec_pct"]:.1f}%</b></div>'
            f'<div class="podium-substat">Disponível: <span class="num">{esc(brl(u["cred"]))}</span></div>'
            f'<button type="button" class="podium-btn" onclick="event.stopPropagation();' + (f'trocaOMDSPorKey(\'{u["key"]}\')' if u["key"] == "BEC6" else f'bcmsDetalheOMDS(\'{u["key"]}\')') + f'">Detalhar Unidade ›</button>'
            f'</div>'
        )
    
    kpis_cmd = (
        kpi_tile("Provisão Recebida (Comando)", brl(cmd_prov), "9 OMDS da Amazônia", "prov", onclick="bcmsModalKpi('prov')") +
        kpi_tile("Empenhado (Comando)", brl(cmd_emp), f"{cmd_exec_pct:.1f}% de execução", "emp", onclick="bcmsModalKpi('emp')") +
        kpi_tile("Liquidado (Comando)", brl(cmd_liq), f"{cmd_liq_pct:.1f}% do empenhado", "liq", onclick="bcmsModalKpi('liq')") +
        kpi_tile("Crédito Disponível", brl(cmd_cred), f"9 OMDS monitoradas", "pag", onclick="bcmsModalKpi('cred')")
    )
    
    ch_cred = svg_comparativo_barras(u_stats, "cred", "Crédito Disponível por Unidade (R$)")
    ch_exec = svg_comparativo_exec(u_stats, cmd_exec_pct)
    
    def _th_r(h, numc):
        cls = ' class="num"' if numc else ''
        return (f'<th{cls} tabindex="0" role="button" aria-sort="none" onclick="bcmsSort(this)" '
                f'onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsSort(this)}}">{esc(h)}<span class="sort"></span></th>')

    ths = (
        _th_r("Pos.", False) +
        _th_r("Organização Militar (OMDS)", False) +
        _th_r("UASGs", False) +
        _th_r("Provisão Recebida", True) +
        _th_r("Empenhado", True) +
        _th_r("% Execução", True) +
        _th_r("Crédito Disponível", True) +
        _th_r("Liquidado", True) +
        _th_r("% Liquidação", True) +
        _th_r("Pago", True) +
        _th_r("Células", True) +
        _th_r("Ação", False)
    )
    
    body_rows = []
    for pos, u in enumerate(rank_exec, 1):
        medalha = "🥇 1º" if pos == 1 else ("🥈 2º" if pos == 2 else ("🥉 3º" if pos == 3 else f"{pos}º"))
        bar_w = min(100.0, u["exec_pct"])
        body_rows.append(
            f'<tr class="cel-row tr-click" tabindex="0" role="button" onclick="' + (f'trocaOMDSPorKey(\'{u["key"]}\')' if u["key"] == "BEC6" else f'bcmsDetalheOMDS(\'{u["key"]}\')') + f'" '
            f'title="Clique para ir ao painel do {esc(u["sigla"])}" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();trocaOMDSPorKey(\'{u["key"]}\')}}">'
            f'<td class="mono2" style="font-weight:700">{medalha}</td>'
            f'<td><div class="tbl-om-cell"><img src="assets/logos/{u["logo"]}" alt="" class="tbl-om-logo" onerror="this.style.display=\'none\'"><b>{esc(u["sigla"])}</b> <span class="tbl-om-sub">{esc(u["nome"])}</span></div></td>'
            f'<td class="mono2">{esc(u["ogu"])} / {esc(u["fex"])}</td>'
            f'<td class="num" data-sort="{u["prov"]:.2f}">{esc(brl(u["prov"]))}</td>'
            f'<td class="num" data-sort="{u["emp"]:.2f}">{esc(brl(u["emp"]))}</td>'
            f'<td class="num" data-sort="{u["exec_pct"]:.2f}"><div class="tbl-pct-cell"><span style="font-weight:700">{u["exec_pct"]:.1f}%</span><div class="mini-track"><div class="mini-fill" style="width:{bar_w:.1f}%;background:{u["accent"]}"></div></div></div></td>'
            f'<td class="num anchor col-pos" data-sort="{u["cred"]:.2f}">{esc(brl(u["cred"]))}</td>'
            f'<td class="num" data-sort="{u["liq"]:.2f}">{esc(brl(u["liq"]))}</td>'
            f'<td class="num" data-sort="{u["liq_pct"]:.2f}">{u["liq_pct"]:.1f}%</td>'
            f'<td class="num" data-sort="{u["pag"]:.2f}">{esc(brl(u["pag"]))}</td>'
            f'<td class="num" data-sort="{u["n_cel"]}">{u["n_cel"]}</td>'
            f'<td><button type="button" class="tbl-action-btn" onclick="event.stopPropagation();' + (f'trocaOMDSPorKey(\'{u["key"]}\')' if u["key"] == "BEC6" else f'bcmsDetalheOMDS(\'{u["key"]}\')') + f'">Detalhar ›</button></td>'
            f'</tr>'
        )
    
    tfoot_tbl = (
        f'<tfoot><tr>'
        f'<td colspan="3"><b>TOTAL CONSOLIDADO DO COMANDO (9 OMDS)</b></td>'
        f'<td class="num"><b>{esc(brl(cmd_prov))}</b></td>'
        f'<td class="num"><b>{esc(brl(cmd_emp))}</b></td>'
        f'<td class="num"><b>{cmd_exec_pct:.1f}%</b></td>'
        f'<td class="num anchor col-pos"><b>{esc(brl(cmd_cred))}</b></td>'
        f'<td class="num"><b>{esc(brl(cmd_liq))}</b></td>'
        f'<td class="num"><b>{cmd_liq_pct:.1f}%</b></td>'
        f'<td class="num"><b>{esc(brl(cmd_pag))}</b></td>'
        f'<td class="num"><b>{cmd_n_cel}</b></td>'
        f'<td>—</td>'
        f'</tr></tfoot>'
    )
    
    tabela_ranking_html = (
        f'<div class="tbl-tools">'
        f'<label class="visually-hidden" for="q-tab-ranking-det">Buscar</label>'
        f'<input type="search" id="q-tab-ranking-det" class="tbl-search" placeholder="Buscar no comparativo por OMDS, UASG..." oninput="bcmsSearch(this,\'tab-ranking-det\')">'
        f'<button type="button" class="btn-excel btn-excel-lg" onclick="bcmsExportTable(this,\'tab-ranking-det\',\'ranking_comparativo_omds\')" title="Baixar comparativo completo das OMDS em planilha formatada para Excel"><span class="btn-excel-ic">📊</span> Exportar Planilha Excel</button>'
        f'<span class="tbl-count" id="cnt-tab-ranking-det" data-unit="unidades" aria-live="polite">9 unidades</span>'
        f'</div>'
        f'<div class="tbl-scroll" id="tab-ranking-det"><table class="det"><thead><tr>{ths}</tr></thead><tbody>{"".join(body_rows)}</tbody>{tfoot_tbl}</table></div>'
    )
    
    hero_eq_cmd = (
        f'<div class="hero-eq-box"><span class="eq-tag">PROVISÃO TOTAL</span><span class="eq-val num">{esc(brl(cmd_prov))}</span></div>'
        f'<span class="hero-eq-sign">−</span>'
        f'<div class="hero-eq-box"><span class="eq-tag">EMPENHADO TOTAL</span><span class="eq-val num eq-emp">{esc(brl(cmd_emp))}</span></div>'
        f'<span class="hero-eq-sign">=</span>'
        f'<div class="hero-eq-box eq-highlight"><span class="eq-tag">DISPONÍVEL COMANDO</span><span class="eq-val num eq-disp">{esc(brl(cmd_cred))}</span></div>'
    )
    
    frag = f"""<section class="unidade unidade-ranking" data-key="RANKING" data-sigla="Comando" style="display:none">
  <div class="ranking-header-card">
    <div class="rh-tag">🏆 BENCHMARKING ORÇAMENTÁRIO & FINANCEIRO</div>
    <h2 class="rh-title">Ranking & Comparativo Consolidado das OMDS</h2>
    <p class="rh-desc">Visão executiva integrada das 9 Organizações Militares Diretamente Subordinadas Diretamente Subordinadas da Base de Apoio Logístico do Exército. Acompanhe os indicadores de desempenho, taxa de execução orçamentária (% Empenhado) e créditos em tela.</p>
  </div>

  <section class="hero hero-cmd">
    <div class="hero-l">
      <div class="eyebrow">Crédito Disponível · Consolidado do Comando (9 OMDS)</div>
      <div class="hero-num num">{esc(brl(cmd_cred))}</div>
      <div class="hero-eq">{hero_eq_cmd}</div>
    </div>
    <div class="hero-r">
      <div class="delta flat">Consolidado das 18 UASGs (OGU + FEx)</div>
      {svg_util(cmd_prov, cmd_emp, cmd_cred)}
    </div>
  </section>

  <section class="sec">
    <div class="eyebrow">Indicadores Globais do Comando</div>
    <div class="kpis">{kpis_cmd}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">🏆 Pódio de Eficiência Orçamentária (% Empenhado / Recebido)</div>
    <p class="sec-nota">Destaque para as organizações com maior percentual de execução das dotações orçamentárias recebidas no exercício. Clique em uma unidade para detalhar.</p>
    <div class="podium-wrap">{"".join(podio_cards)}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Comparativos Visuais entre as Unidades</div>
    <div class="grid2">{ch_exec}{ch_cred}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Tabela Comparativa e Benchmarking Completo</div>
    <p class="sec-nota">Relação completa de todas as OMDS com dados consolidados de OGU e FEx. Ordene por qualquer coluna ou clique no botão para exportar para Excel (.xlsx).</p>
    {tabela_ranking_html}
  </section>
</section>"""
    return frag

UG_NOMES_PADRAO = {
    "160238": "Base de Apoio Logístico do Exército",
    "167238": "Base de Apoio Logístico do Exército (FEx)",
    "160329": "Batalhão Central de Manutenção e Suprimento",
    "167329": "Batalhão Central de Manutenção e Suprimento (FEx)",
    "160246": "Depósito Central de Munição",
    "167246": "Depósito Central de Munição (FEx)",
    "160304": "BMSA",
    "167304": "BMSA (FEx)",
    "160307": "1º Depósito de Suprimento",
    "167307": "1º Depósito de Suprimento (FEx)",
    "160321": "ECT",
    "167321": "ECT (FEx)",
    "160070": "Comando Logístico (COLOG)",
    "160069": "D Abst - Diretoria de Abastecimento",
    "160068": "D Mat - Diretoria de Material",
    "160035": "Dep Prom / D Abst",
    "160525": "DCT - Depto de Ciência e Tecnologia",
    "160065": "EME - Estado-Maior do Exército",
}

# ---------------- nova aba: Histórico de Notas de Crédito ----------------
def secao_historico_ncs(res, hist, data_str, periodo):
    """Gera a aba consolidada 'Histórico de Notas de Crédito' do exercício corrente.
    
    Traz todas as NCs emitidas e recebidas pelas OMDS no ano com 4 KPIs analíticos no topo,
    segmentadores multicritério (Período, Fonte, PTRES, Faixa de Saldo, Busca) e grid interativo
    ordenável com detalhamento completo via modal flutuante.
    """
    posicao = periodo if periodo else (data_str[8:10] + "/" + data_str[5:7] + "/" + data_str[0:4])
    hoje = datetime.date.today()

    # 1. Agrupa recebimentos por célula orçamentária para atribuir saldo às NCs
    cel_recs = {}
    todas_ncs = {}

    for cod, d in res.items():
        om_sigla = d["nome"].split("·")[0].strip()
        fonte_tipo = "OGU" if cod.startswith("160") else "FEx"
        for L in d["linhas"]:
            nc = L.get("nc")
            if not nc:
                continue
            key = (nc, cod)
            op_str = (L.get("op") or "").upper()
            emit_str = L.get("emit") or ""
            prov_liq = L.get("prov", 0.0) - L.get("conc", 0.0)
            crd_val = L.get("cred", 0.0)
            is_det = ("DETALHAMENTO" in op_str) or (emit_str == cod and abs(prov_liq) < 0.01)

            if key not in todas_ncs:
                todas_ncs[key] = {
                    "nc": nc, "uasg": cod, "fav_nome": d["nome"], "om_sigla": om_sigla,
                    "fonte": fonte_tipo,
                    "emit": L.get("emit") or "",
                    "emit_nome": L.get("emit_nome") or "",
                    "dia": L.get("dia") or "",
                    "acao": L.get("acao") or "",
                    "pi": L.get("pi") or "",
                    "pi_nome": L.get("pi_nome") or "",
                    "nd": L.get("nd") or "",
                    "nd_desc": L.get("nd_desc") or "",
                    "nd_de": L.get("nd_de") or "",
                    "op": L.get("op") or "",
                    "obj": L.get("obj") or "",
                    "prov": 0.0, "cred_calc": 0.0, "emp_calc": 0.0,
                    "liq": 0.0, "pag": 0.0, "conc": 0.0,
                    "is_det": is_det,
                    "linhas": []
                }
            it = todas_ncs[key]
            it["linhas"].append(L)
            it["liq"]  += L.get("liq", 0.0)
            it["pag"]  += L.get("pag", 0.0)
            it["conc"] += L.get("conc", 0.0)
            if not it["dia"] and L.get("dia"): it["dia"] = L["dia"]
            if not it["obj"] and L.get("obj"): it["obj"] = L["obj"]
            if not it["emit"] and L.get("emit"): it["emit"] = L["emit"]
            if not it["emit_nome"] and L.get("emit_nome"): it["emit_nome"] = L["emit_nome"]
            if L.get("nd_de"): it["nd_de"] = L.get("nd_de")

            if not is_det:
                it["prov"] += prov_liq
                val_inflow = prov_liq
            else:
                it["is_det"] = True
                val_inflow = max(0.0, crd_val)

            if val_inflow > 0.005:
                cel_key = (cod, L.get("acao"), L.get("pi"), L.get("nd"))
                cel_recs.setdefault(cel_key, []).append((key, val_inflow, L.get("dia")))

    # 2. Atribui o saldo restante em tela das células para as NCs (ordem cronológica decrescente)
    for cod, d in res.items():
        for (acao, pi, nd), cel in d["celulas"].items():
            saldo_cel = cel["cred"]
            cel_key = (cod, acao, pi, nd)
            recs = cel_recs.get(cel_key, [])
            recs_sorted = sorted(recs, key=lambda x: _dt_br(x[2]) or datetime.date.min, reverse=True)
            restante = saldo_cel
            for nc_key, val_rec, _ in recs_sorted:
                if restante <= 0.005:
                    break
                atribuido = min(val_rec, restante)
                todas_ncs[nc_key]["cred_calc"] += atribuido
                restante -= atribuido

    # 3. Formata itens e constrói dicionário HISTDATA
    hist_list = []
    histdata = {}
    ptres_set = set()
    meses_set = set()

    for idx, ((nc, cod), it) in enumerate(todas_ncs.items(), 1):
        hid = f"h_{idx}"
        dt = _dt_br(it["dia"])
        dt_iso = dt.isoformat() if dt else ""
        mes = f"{dt.month:02d}/{dt.year}" if dt else ""
        tri = f"T{(dt.month-1)//3 + 1}" if dt else ""
        if mes: meses_set.add(mes)
        if it["acao"]: ptres_set.add(it["acao"])

        dias = (hoje - dt).days if dt else None
        prov = it["prov"]
        cred = it["cred_calc"]
        is_det = it.get("is_det", False)
        emp = max(0.0, prov - cred) if not is_det else 0.0

        op_up = (it["op"] or "").upper()
        if "ANULA" in op_up or "CANCEL" in op_up or prov < -0.01:
            status = "Cancelada / Anulada"
            status_slug = "canc"
            faixa_saldo = "canc"
        elif is_det:
            status = "Detalhamento de ND"
            status_slug = "detalhada"
            faixa_saldo = "detalhada" if cred <= 0.01 else "saldo_pos"
        elif cred <= 0.01 and emp > 0:
            status = "Totalmente Executada"
            status_slug = "exec"
            faixa_saldo = "zerada"
        elif emp > 0 and cred > 0.01:
            status = "Parcialmente Executada"
            status_slug = "parcial"
            faixa_saldo = "saldo_pos"
        elif emp == 0 and cred > 0.01:
            status = "Disponível"
            status_slug = "disp"
            faixa_saldo = "saldo_pos"
        else:
            status = "Sem Saldo / Zerada"
            status_slug = "zerada"
            faixa_saldo = "zerada"

        emit_cod = it["emit"]
        emit_nome = it["emit_nome"]
        if not emit_nome or emit_nome in ("-9", "NAO SE APLICA"):
            emit_nome = UG_NOMES_PADRAO.get(emit_cod, f"UG {emit_cod}" if emit_cod else "Não informado")

        item_dict = {
            "hid": hid,
            "nc": nc,
            "dia": it["dia"],
            "dt": dt_iso,
            "mes": mes,
            "tri": tri,
            "dias": dias,
            "emit_cod": emit_cod,
            "emit_nome": emit_nome,
            "fav_cod": cod,
            "fav_nome": it["fav_nome"],
            "om_sigla": it["om_sigla"],
            "ptres": it["acao"],
            "acao": it["acao"],
            "acao_desc": f"Ação Governamental {it['acao']}" if it["acao"] else "",
            "fonte": it["fonte"],
            "nd": it["nd"],
            "nd_desc": it["nd_desc"],
            "nd_de": it.get("nd_de", ""),
            "is_det": is_det,
            "pi": it["pi"],
            "pi_desc": it["pi_nome"],
            "op": it["op"] or "Descentralização de Crédito",
            "obj": it["obj"] or "(Sem descrição detalhada)",
            "prov": round(prov, 2),
            "bloq": 0.0,
            "emp": round(emp, 2),
            "liq": round(it["liq"], 2),
            "pag": round(it["pag"], 2),
            "cred": round(cred, 2),
            "status": status,
            "status_slug": status_slug,
            "faixa_saldo": faixa_saldo,
            "ano": str(dt.year) if dt else "2026",
            "itens": [{
                "acao": L.get("acao", ""),
                "pi": L.get("pi", ""),
                "nd": L.get("nd", ""),
                "op": L.get("op", ""),
                "prov": round(L.get("prov", 0.0), 2),
                "emp": round(L.get("emp", 0.0), 2),
                "cred": round(L.get("cred", 0.0), 2),
            } for L in it["linhas"]] if len(it["linhas"]) > 1 else []
        }
        histdata[hid] = item_dict
        hist_list.append(item_dict)

    # Ordenação padrão: mais recentes primeiro (Data Decrescente)
    hist_list.sort(key=lambda x: (x["dt"] or "", x["cred"]), reverse=True)

    # Totais dos KPIs sem duplicidades
    tot_distintas = len(set(x["nc"] for x in hist_list))
    tot_prov = sum(x["prov"] for x in hist_list)
    tot_cred = sum(x["cred"] for x in hist_list)
    tot_emp  = max(0.0, tot_prov - tot_cred)
    p_emp = (tot_emp / tot_prov * 100) if tot_prov else 0.0
    p_cred = (tot_cred / tot_prov * 100) if tot_prov else 0.0

    # Opções para os filtros
    opt_ptres = "".join(f'<option value="{esc(p)}">{esc(p)}</option>' for p in sorted(ptres_set))
    opt_meses = "".join(f'<option value="{esc(m)}">{esc(m)}</option>' for m in sorted(meses_set, reverse=True))

    kpis_html = (
        kpi_tile("Total de NCs Distintas", str(tot_distintas), f"{len(hist_list)} lançamentos no ano", "total", id_v="kpi-hist-total") +
        kpi_tile("Total Descentralizado", brl(tot_prov), "100% da provisão recebida", "prov", id_v="kpi-hist-prov") +
        kpi_tile("Total Empenhado", brl(tot_emp), f"{p_emp:.1f}% executado", "emp", id_v="kpi-hist-emp") +
        kpi_tile("Saldo Disponível", brl(tot_cred), f"{p_cred:.1f}% remanescente", "pag", id_v="kpi-hist-cred")
    )

    opt_ptres = '<option value="">Ação: todas</option>' + "".join(f'<option value="{esc(p)}">{esc(p)}</option>' for p in sorted(ptres_set))
    opt_meses = (
        '<option value="">Período: todos</option>'
        '<optgroup label="Trimestres">'
        '<option value="T1">1º Trimestre (Jan–Mar)</option>'
        '<option value="T2">2º Trimestre (Abr–Jun)</option>'
        '<option value="T3">3º Trimestre (Jul–Set)</option>'
        '<option value="T4">4º Trimestre (Out–Dez)</option>'
        '</optgroup>'
        '<optgroup label="Meses">'
        + "".join(f'<option value="{esc(m)}">{esc(m)}</option>' for m in sorted(meses_set, reverse=True))
        + '</optgroup>'
    )

    om_unicas = sorted(list(set(it["om_sigla"] for it in hist_list if it.get("om_sigla"))))
    opt_oms = "".join(f'<option value="{esc(om)}">{esc(om)}</option>' for om in om_unicas)

    tools_hist_html = (
        f'<div class="tbl-tools">'
        f'<label class="visually-hidden" for="flt-hist-busca">Buscar</label>'
        f'<input type="search" id="flt-hist-busca" class="tbl-search" placeholder="Buscar por NC, Justificativa, UG, PTRES, ND ou PI…" oninput="bcmsFiltraHistorico()">'
        f'<button type="button" class="btn-excel btn-excel-lg" onclick="bcmsExportHistoricoExcel()" title="Baixar histórico consolidado em planilha formatada para Excel"><span class="btn-excel-ic">📊</span> Exportar Histórico Completo (Excel)</button>'
        f'<span class="tbl-count" id="cnt-hist-ncs" data-unit="Notas de Crédito" aria-live="polite">Exibindo {len(hist_list)} de {len(hist_list)} Notas de Crédito ({tot_distintas} distintas)</span>'
        f'</div>'
        f'<div class="tbl-filtros" id="flt-hist" role="group" aria-label="Filtros do histórico consolidado">'
        f'<span class="flt-lbl">Filtrar:</span>'
        f'<select class="flt" id="flt-hist-om" aria-label="Filtrar por organização militar" onchange="bcmsFiltraHistorico()">'
        f'<option value="">Unidade: todas</option>'
        f'{opt_oms}'
        f'</select>'
        f'<select class="flt" id="flt-hist-fonte" aria-label="Filtrar por fonte" onchange="bcmsFiltraHistorico()">'
        f'<option value="">Fonte: todas</option>'
        f'<option value="OGU">OGU (160)</option>'
        f'<option value="FEx">FEx (167)</option>'
        f'</select>'
        f'<select class="flt" id="flt-hist-ptres" aria-label="Filtrar por ação de governo" onchange="bcmsFiltraHistorico()">{opt_ptres}</select>'
        f'<select class="flt" id="flt-hist-periodo" aria-label="Filtrar por período" onchange="bcmsFiltraHistorico()">{opt_meses}</select>'
        f'<select class="flt" id="flt-hist-faixa" aria-label="Filtrar por status do saldo" onchange="bcmsFiltraHistorico()">'
        f'<option value="">Status: todos</option>'
        f'<option value="saldo_pos">🟢 Com Saldo (&gt; R$ 0)</option>'
        f'<option value="parcial">🟡 Parcialmente Executadas</option>'
        f'<option value="zerada">⚪ Executadas / Zeradas</option>'
        f'<option value="detalhada">🔄 Detalhamentos de ND</option>'
        f'<option value="canc">🔴 Anuladas / Canceladas</option>'
        f'</select>'
        f'<button type="button" class="flt-limpa" onclick="bcmsLimpaFiltrosHistorico()" title="Limpar todos os filtros">✕ Limpar</button>'
        f'<span class="flt-resumo" id="flt-hist-res" aria-live="polite"></span>'
        f'</div>'
    )

    initial_rows = []
    for item in hist_list[:50]:
        status_cls = f"status-{item['status_slug']}"
        emit_nome_curto = (item["emit_nome"][:22] + "…") if len(item["emit_nome"]) > 22 else item["emit_nome"]
        desc_completa = item.get("obj", "")
        desc_resumo = desc_completa[:118] + ("…" if len(desc_completa) > 118 else "")
        nc_full = str(item["nc"] or "")
        m_nc = re.search(r"NC(\d+)$", nc_full)
        if m_nc:
            nc_lbl = f'<span class="nc-num">NC {esc(m_nc.group(1))}</span> <span class="nc-ug">· {esc(nc_full[:6])}</span>'
        else:
            nc_lbl = f'<span class="nc-num">{esc(nc_full)}</span>'
        acao_nd = f'{item["ptres"]} · {item["nd"]}' if item.get("ptres") and item.get("nd") else (item.get("ptres") or item.get("nd") or "—")

        initial_rows.append(
            f'<tr class="cel-row" data-hid="{item["hid"]}" tabindex="0" role="button" onclick="bcmsDetalheNC(\'{item["hid"]}\')" '
            f'title="Clique para abrir o detalhamento completo da NC {esc(nc_full)}" onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsDetalheNC(\'{item["hid"]}\')}}">'
            f'<td><span class="pill-fonte">{esc(item["fonte"])}</span></td>'
            f'<td class="mono2" title="{esc(nc_full)}">{nc_lbl}</td>'
            f'<td title="{esc(item["fav_nome"])}"><span class="ug-pill fav">{esc(item["fav_cod"])}</span> <b>{esc(item["om_sigla"])}</b></td>'
            f'<td class="mono2">{esc(acao_nd)}</td>'
            f'<td class="obj" title="{esc(desc_completa)}" data-full-desc="{esc(desc_completa)}">{esc(desc_resumo)}</td>'
            f'<td title="{esc(item["emit_nome"])}"><span class="ug-pill emit">{esc(item["emit_cod"])}</span> <small>{esc(emit_nome_curto)}</small></td>'
            f'<td class="mono2">{esc(item["dia"] or "—")}</td>'
            f'<td class="num" data-sort="{item["prov"]:.2f}">{esc(brl(item["prov"]))}</td>'
            f'<td class="num anchor" data-sort="{item["cred"]:.2f}">{esc(brl(item["cred"]))}</td>'
            f'<td class="num"><span class="pill-status {status_cls}">{esc(item["status"])}</span><i class="chev" aria-hidden="true">›</i></td>'
            f'</tr>'
        )

    tot_pages = math.ceil(len(hist_list)/50) if hist_list else 1
    table_hist_html = (
        f'<div class="tbl-scroll" id="scroll-hist-ncs">'
        f'<table class="tbl tbl-hist" id="tbl-hist-ncs">'
        f'<thead>'
        f'<tr>'
        f'<th>Fonte</th>'
        f'<th>Nota de Crédito</th>'
        f'<th>Favorecido</th>'
        f'<th>Ação · ND</th>'
        f'<th class="obj-col">Justificativa / Objeto</th>'
        f'<th>Emitente</th>'
        f'<th tabindex="0" role="button" aria-sort="descending" onclick="bcmsSortHistorico(\'dt\', this)" title="Ordenar por Data">Recebido em <span class="sort">▼</span></th>'
        f'<th class="num" tabindex="0" role="button" aria-sort="none" onclick="bcmsSortHistorico(\'prov\', this)" title="Ordenar por Valor Original">Provisão <span class="sort"></span></th>'
        f'<th class="num" tabindex="0" role="button" aria-sort="none" onclick="bcmsSortHistorico(\'cred\', this)" title="Ordenar por Saldo Disponível">Crédito Disp. <span class="sort"></span></th>'
        f'<th tabindex="0" role="button" aria-sort="none" onclick="bcmsSortHistorico(\'status\', this)" title="Ordenar por Status">Status <span class="sort"></span></th>'
        f'</tr>'
        f'</thead>'
        f'<tbody id="tbody-hist-ncs">'
        f'{"".join(initial_rows)}'
        f'</tbody>'
        f'<tfoot>'
        f'<tr>'
        f'<td colspan="7" id="tf-hist-label">TOTAL · {len(hist_list)} Nota(s) de Crédito no histórico</td>'
        f'<td class="num" id="tf-hist-prov">{esc(brl(tot_prov))}</td>'
        f'<td class="num anchor" id="tf-hist-cred">{esc(brl(tot_cred))}</td>'
        f'<td>—</td>'
        f'</tr>'
        f'</tfoot>'
        f'</table>'
        f'</div>'
        f'<div class="tbl-tools" style="margin-top:10px;justify-content:space-between;" id="paginacao-hist">'
        f'<span class="pag-info" id="pag-info-txt" style="font-size:0.8125rem;color:var(--ink-muted);font-weight:600;">Página 1 de {tot_pages} (Exibindo 1–{min(50, len(hist_list))} de {len(hist_list)})</span>'
        f'<div style="display:flex;gap:8px;align-items:center;">'
        f'<button type="button" class="flt-limpa" id="btn-pag-ant" onclick="bcmsPaginaHistorico(-1)" disabled>‹ Anterior</button>'
        f'<button type="button" class="flt-limpa" id="btn-pag-prox" onclick="bcmsPaginaHistorico(1)"' + (' disabled' if tot_pages <= 1 else '') + '>Próxima ›</button>'
        f'</div>'
        f'</div>'
    )

    frag = f"""<section class="unidade unidade-hist" id="secao-HISTORICO" data-key="HISTORICO" style="display:none">
  <div class="ranking-header-card" style="border-top-color:var(--primary-600);">
    <div class="rh-tag" style="color:var(--primary);">EXERCÍCIO FINANCEIRO CORRENTE · PAINEL ANALÍTICO CONSOLIDADO</div>
    <h2 class="rh-title">📜 Histórico Consolidado de Notas de Crédito (NC)</h2>
    <p class="rh-desc">Repositório analítico de todas as Notas de Crédito recebidas no exercício pelo 6º Batalhão de Engenharia de Construção (UASGs 160353 OGU e 167353 FEx). Pesquise por termos da justificativa, combine filtros por fonte, ação e período, ou clique em qualquer linha para inspecionar o detalhamento cadastral e financeiro completo no modal.</p>
  </div>

  <section class="sec">
    <div class="eyebrow">Indicadores Globais de Notas de Crédito no Exercício</div>
    <div class="kpis">{kpis_html}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Base Completa de Notas de Crédito do Comando</div>
    <p class="sec-nota">Relação consolidada de <b>todas as Notas de Crédito</b> das 6 OMDS (12 UASGs de OGU e FEx). Utilize a busca rápida e os filtros seletivos para auditar lançamentos. <b>Clique em uma linha</b> para abrir a ficha completa.</p>
    {tools_hist_html}
    {table_hist_html}
  </section>
</section>"""

    histdata_payload = {
        "summary": {
            "total_distintas": tot_distintas,
            "tot_prov": round(tot_prov, 2),
            "tot_emp": round(tot_emp, 2),
            "tot_cred": round(tot_cred, 2),
            "p_emp": round(p_emp, 1),
            "p_cred": round(p_cred, 1),
        },
        "items": hist_list,
        "by_id": histdata,
    }
    return frag, histdata_payload

# ---------------- shell da página (multi-OMDS) ----------------

# ---------------- módulo Operação Catrimani ----------------
def svg_catrimani_nds(nds_list, titulo="Despesas por Natureza de Despesa — Operação Catrimani (R$)"):
    itens = [(x["nome"][:24], x["emp"], "#15803D") for x in nds_list if x["emp"] > 1.0][:8]
    if not itens:
        return f'<div class="card chart"><div class="eyebrow">{esc(titulo)}</div><p class="vazio">Sem dados de empenho por ND</p></div>'
    vmax = max(x[1] for x in itens) or 1
    labW, rh, pad = 160, 36, 12
    zx = labW + 10
    plotW = 400
    W = zx + plotW + 140
    H = pad * 2 + rh * len(itens)
    el = []
    for i, (nome, val, accent) in enumerate(itens):
        y = pad + i * rh
        w = max(4, (val / vmax) * plotW)
        el.append(f'<text x="{labW-8}" y="{y+rh/2+4:.0f}" text-anchor="end" class="s-cat" style="font-weight:600;font-size:11px;">{esc(nome)}</text>')
        el.append(f'<rect x="{zx}" y="{y+4}" width="{w:.1f}" height="{rh-8}" rx="4" style="fill:{accent}"><title>{esc(nome)}: {esc(brl(val))}</title></rect>')
        el.append(f'<text x="{zx+w+8:.1f}" y="{y+rh/2+4:.0f}" class="s-num s-on2" style="font-weight:700">{esc(brl(val))}</text>')
    svg = f'<svg viewBox="0 0 {W} {H}" class="svg" role="img" aria-label="{esc(titulo)}">{"".join(el)}</svg>'
    return f'<div class="card chart"><div class="eyebrow">{esc(titulo)}</div>{svg}</div>'

def secao_operacao_catrimani(catr, data_str, periodo):
    tot = catr["totais"]
    por_ug = catr["por_ug"]
    por_nd = catr["por_nd"]
    linhas = catr["linhas"]

    kpis_html = (
        kpi_tile("Dotação Descentralizada", brl(tot["prov"]), "100% dos recursos provisionados", "prov", id_v="catr-kpi-prov") +
        kpi_tile("Despesas Empenhadas", brl(tot["emp"]), f"{tot['pct_emp']:.1f}% executado", "emp", id_v="catr-kpi-emp") +
        kpi_tile("Crédito Disponível", brl(tot["cred"]), f"{(100-tot['pct_emp']):.1f}% livre p/ empenho", "pag", id_v="catr-kpi-cred") +
        kpi_tile("Despesas Liquidadas", brl(tot["liq"]), f"{tot['pct_liq']:.1f}% do empenho liquidado", "conc", id_v="catr-kpi-liq")
    )

    hero_eq = (
        f'<div class="eq-chips" role="group" aria-label="Equação do Crédito Operação Catrimani">'
        f'<span class="eq-chip eq-prov" title="Provisão recebida no exercício"><b>{esc(brl(tot["prov"]))}</b> <span>Recebido</span></span>'
        f'<span class="eq-op" aria-hidden="true">−</span>'
        f'<span class="eq-chip eq-emp" title="Despesas já empenhadas"><b>{esc(brl(tot["emp"]))}</b> <span>Empenhado</span></span>'
        f'<span class="eq-op" aria-hidden="true">=</span>'
        f'<span class="eq-chip eq-cred" title="Crédito disponível para novos empenhos"><b>{esc(brl(tot["cred"]))}</b> <span>Disponível em Tela</span></span>'
        f'</div>'
    )

    # Tabela comparativa das 10 UGs
    distinct_catr_ncs = catr.get("ncs_distintas", len(set(x["nc"] for x in linhas if x.get("nc"))))

    ug_rows = []
    for u in por_ug:
        sem_cor = "var(--ok, #10B981)" if u["pct_emp"] >= 80.0 else ("var(--gold, #F59E0B)" if u["pct_emp"] >= 60.0 else "var(--bad, #EF4444)")
        sig_ug = u.get("sigla", u["cod"])
        ug_rows.append(
            f'<tr class="tr-click" tabindex="0" role="button" onclick="bcmsDetalheUG(\'{esc(u["cod"])}\')" '
            f'onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsDetalheUG(\'{esc(u["cod"])}\');}}" '
            f'title="Clique para ver o detalhamento completo de dotações, empenhos e NCs de {esc(u["nome"])}">'
            f'<td class="col-ug"><b style="font-size:0.8125rem;color:var(--primary);">{esc(sig_ug)}</b><span style="display:block;font-size:0.6875rem;color:var(--ink-muted);font-weight:normal;">UG {esc(u["cod"])}</span></td>'
            f'<td class="col-nome"><b>{esc(u["nome"])}</b></td>'
            f'<td class="num col-moeda">{esc(brl(u["prov"]))}</td>'
            f'<td class="num col-moeda" style="font-weight:700;">{esc(brl(u["emp"]))}</td>'
            f'<td class="num col-moeda anchor" style="font-weight:700;" title="Crédito Disponível Líquido: {esc(brl(u["cred"]))} (Clique para detalhar)">{esc(brl(u["cred"]))}</td>'
            f'<td class="num col-moeda">{esc(brl(u["liq"]))}</td>'
            f'<td class="num col-moeda">{esc(brl(u["pag"]))}</td>'
            f'<td class="num col-pct">'
            f'  <div class="bar-pct-wrap">'
            f'    <div class="bar-pct-track">'
            f'      <div class="bar-pct-fill" style="width:{min(100.0, max(2.0, u["pct_emp"])):.1f}%;background:{sem_cor};"></div>'
            f'    </div>'
            f'    <span class="bar-pct-val" style="color:{sem_cor};">{u["pct_emp"]:.1f}%</span>'
            f'  </div>'
            f'</td>'
            f'</tr>'
        )

    tabela_ugs_html = (
        f'<div class="tbl-wrap tbl-scroll" style="margin-top:16px;">'
        f'<table class="tbl tbl-ugs" aria-label="Comparativo de UGs na Operação Catrimani">'
        f'<thead><tr>'
        f'<th class="col-ug">Unidade</th><th class="col-nome">Descrição da OM Executora</th><th class="num col-moeda">Recebido</th>'
        f'<th class="num col-moeda">Empenhado</th><th class="num col-moeda">Crédito Disponível</th>'
        f'<th class="num col-moeda">Liquidado</th><th class="num col-moeda">Pago</th><th class="num col-pct">% Execução</th>'
        f'</tr></thead>'
        f'<tbody>{"".join(ug_rows)}</tbody>'
        f'<tfoot><tr>'
        f'<td colspan="2" class="tfoot-label"><b>TOTAL CONSOLIDADO · {len(por_ug)} UGs EXECUTORAS</b></td>'
        f'<td class="num col-moeda"><b>{esc(brl(tot["prov"]))}</b></td>'
        f'<td class="num col-moeda"><b>{esc(brl(tot["emp"]))}</b></td>'
        f'<td class="num col-moeda anchor"><b>{esc(brl(tot["cred"]))}</b></td>'
        f'<td class="num col-moeda"><b>{esc(brl(tot["liq"]))}</b></td>'
        f'<td class="num col-moeda"><b>{esc(brl(tot["pag"]))}</b></td>'
        f'<td class="num col-pct"><b>{tot["pct_emp"]:.1f}%</b></td>'
        f'</tr></tfoot>'
        f'</table></div>'
    )

    # Select de UGs para o filtro
    opt_ugs = '<option value="">UG: todas as 10 UGs</option>' + "".join(
        f'<option value="{esc(u["cod"])}">{esc(u["cod"])} · {esc(u["nome"])}</option>' for u in por_ug
    )

    initial_catr_rows = []
    for item in linhas[:25]:
        st_cor = "var(--ok, #10B981)" if item["cred"] > 0.01 else "var(--ink-muted)"
        st_txt = "Com Saldo" if item["cred"] > 0.01 else "Empenhada"
        nc_cod = esc(item["nc"])
        initial_catr_rows.append(
            f'<tr class="tr-click" tabindex="0" role="button" onclick="bcmsOpenNCModalManual(\'{nc_cod}\')" '
            f'onkeydown="if(event.key===\'Enter\'||event.key===\' \'){{event.preventDefault();bcmsOpenNCModalManual(\'{nc_cod}\');}}" '
            f'title="Clique para abrir a ficha cadastral completa desta NC">'
            f'<td>{esc(item["dia"])}</td>'
            f'<td><b class="nc-mono">{esc(item["nc"])}</b></td>'
            f'<td><b>{esc(item["ug"])}</b> <span class="tbl-om-sub">{esc(item["ug_nome"])}</span></td>'
            f'<td><span class="pill-ptres">{esc(item["acao"])}</span> · <span class="pill-pi">{esc(item["pi"])}</span></td>'
            f'<td>{esc(item["nd"])} <span class="tbl-om-sub">{esc(item["nd_desc"][:20])}</span></td>'
            f'<td class="wrap-txt" title="{esc(item["obj"])}">{esc(item["obj"][:55])}...</td>'
            f'<td class="num">{esc(brl(item["prov"]))}</td>'
            f'<td class="num">{esc(brl(item["emp"]))}</td>'
            f'<td class="num anchor" style="font-weight:700;">{esc(brl(item["cred"]))}</td>'
            f'<td><span class="pill-nd" style="background:var(--track);color:{st_cor};font-weight:700;">{st_txt}</span></td>'
            f'</tr>'
        )

    tot_paginas_catr = math.ceil(len(linhas) / 25) or 1

    frag = f"""<section class="unidade unidade-catrimani" id="secao-CATRIMANI" data-key="CATRIMANI" style="display:none">
  <div class="ranking-header-card" style="border-top-color:#15803D;background:linear-gradient(180deg, rgba(21,128,61,0.08) 0%, transparent 100%);">
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:8px;">
      <img src="assets/logos/CATRIMANI.png" alt="Operação Catrimani" style="width:48px;height:48px;border-radius:50%;box-shadow:0 4px 12px rgba(0,0,0,0.15);">
      <div>
        <div class="rh-tag" style="color:#15803D;font-weight:800;">OPERAÇÃO CONJUNTA · TERRITÓRIO INDÍGENA YANOMAMI / RORAIMA</div>
        <h2 class="rh-title" style="margin:0;">🎖️ Operação Catrimani II — Execução Orçamentária Multi-UGs</h2>
      </div>
    </div>
    <p class="rh-desc">Painel executivo oficial de acompanhamento das descentralizações de crédito orçamentário da <b>Ação Governamental 21EM</b> e Planos Internos correlatos. Consolida todas as movimentações financeiras das 10 Unidades Gestoras Executoras da Amazônia, com extrato analítico de empenhos, liquidações e saldos livres disponíveis em tela.</p>
  </div>

  <section class="hero" style="margin-top:20px;">
    <div class="hero-l">
      <div class="eyebrow">Crédito Disponível Consolidado da Operação</div>
      <div class="hero-num num" style="color:#15803D;">{esc(brl(tot["cred"]))}</div>
      <div class="hero-eq">{hero_eq}</div>
    </div>
    <div class="hero-r">
      {svg_util(tot["prov"], tot["emp"], tot["cred"])}
    </div>
  </section>

  <section class="sec">
    <div class="eyebrow">Indicadores Globais da Operação Catrimani</div>
    <div class="kpis">{kpis_html}</div>
  </section>

  <section class="sec">
    <div class="eyebrow">Composição Visual por Natureza de Despesa</div>
    <div class="grid2">
      {svg_catrimani_nds(por_nd)}
      <div class="card chart">
        <div class="eyebrow">Desdobramento da Execução por Estágio</div>
        {svg_waterfall(tot["prov"], tot["emp"], tot["cred"])}
      </div>
    </div>
  </section>

  <section class="sec">
    <div class="eyebrow">Acompanhamento Orçamentário por Unidade Gestora Executora</div>
    <p class="sec-nota">Relação consolidada das <b>10 Organizações Militares</b> executoras da Operação Catrimani II. <b>Clique em qualquer linha ou crédito</b> para abrir a ficha completa com balanço, despesas por ND e extrato de NCs daquela UG.</p>
    {tabela_ugs_html}
  </section>

  <section class="sec">
    <div class="eyebrow">Extrato Completo de Notas de Crédito da Operação Catrimani</div>
    <p class="sec-nota">Relação auditada das <b>{distinct_catr_ncs} Notas de Crédito distintas</b> ({len(linhas)} lançamentos de dotação) recebidas pelas 10 UGs executoras no âmbito da Operação Catrimani II. Utilize os filtros interativos de UG, Fonte e Saldo, pesquise em tempo real ou exporte a relação completa para Excel. <b>Clique em qualquer linha</b> para abrir a ficha cadastral no modal.</p>
    <div class="tbl-tools">
      <input type="search" id="flt-catr-busca" class="tbl-search" placeholder="Buscar por NC, Favorecido, Objeto, ND ou PI (multi-termos)…" oninput="bcmsFiltraCatrimani()">
      <button type="button" class="flt-limpa" onclick="bcmsLimparFiltrosCatrimani()" title="Limpar todos os termos e filtros de pesquisa">↺ Limpar Filtros</button>
      <button type="button" class="btn-excel btn-excel-lg" onclick="bcmsExportCatrimaniExcel()" title="Baixar relatório completo da Operação Catrimani em planilha Excel formatada">
        <span class="btn-excel-ic">📊</span> Exportar Catrimani em Excel
      </button>
      <span class="tbl-count" id="cnt-catr-ncs" aria-live="polite">Exibindo {min(25, len(linhas))} de {len(linhas)} lançamentos ({distinct_catr_ncs} NCs distintas)</span>
    </div>
    <div class="tbl-filtros" role="group" aria-label="Filtros da Operação Catrimani">
      <span class="flt-lbl">Filtrar:</span>
      <select class="flt" id="flt-catr-ug" aria-label="Filtrar por UG executora" onchange="bcmsFiltraCatrimani()">{opt_ugs}</select>
      <select class="flt" id="flt-catr-fonte" aria-label="Filtrar por fonte" onchange="bcmsFiltraCatrimani()">
        <option value="">Fonte: todas</option>
        <option value="160">OGU (160)</option>
        <option value="167">FEx (167)</option>
      </select>
      <select class="flt" id="flt-catr-saldo" aria-label="Filtrar por saldo" onchange="bcmsFiltraCatrimani()">
        <option value="">Saldo: todos</option>
        <option value="com_saldo">🟢 Apenas com Saldo Livre (&gt; R$ 0)</option>
        <option value="zerada">⚪ Empenhadas / Zeradas</option>
      </select>
    </div>
    <div class="tbl-wrap">
      <table class="tbl" id="tab-catrimani-ncs" aria-label="Notas de Crédito da Operação Catrimani">
        <thead>
          <tr>
            <th>Emissão</th>
            <th>NC</th>
            <th>UG Executora</th>
            <th>Ação · PI</th>
            <th>Natureza Despesa</th>
            <th>Objeto / Finalidade</th>
            <th class="num">Recebido</th>
            <th class="num">Empenhado</th>
            <th class="num">Crédito Disp.</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody id="tbody-catr-ncs">
          {"".join(initial_catr_rows)}
        </tbody>
        <tfoot>
          <tr>
            <td colspan="6"><b>TOTAL CONSOLIDADO · {distinct_catr_ncs} NOTAS DE CRÉDITO ({len(linhas)} ITENS)</b></td>
            <td class="num"><b>{esc(brl(tot["prov"]))}</b></td>
            <td class="num"><b>{esc(brl(tot["emp"]))}</b></td>
            <td class="num anchor"><b>{esc(brl(tot["cred"]))}</b></td>
            <td>—</td>
          </tr>
        </tfoot>
      </table>
    </div>
    <div class="tbl-tools" style="margin-top:10px;justify-content:space-between;" id="paginacao-catr">
      <span class="pag-info" id="pag-catr-txt" style="font-size:0.8125rem;color:var(--ink-muted);font-weight:600;">Página 1 de {tot_paginas_catr} (Exibindo 1–{min(25, len(linhas))} de {len(linhas)})</span>
      <div style="display:flex;gap:8px;align-items:center;">
        <button type="button" class="flt-limpa" id="btn-catr-ant" onclick="bcmsPaginaCatrimani(-1)" disabled>‹ Anterior</button>
        <button type="button" class="flt-limpa" id="btn-catr-prox" onclick="bcmsPaginaCatrimani(1)"{" disabled" if tot_paginas_catr <= 1 else ""}>Próxima ›</button>
      </div>
    </div>
  </section>
</section>"""

    return frag


def montar_pagina(res, hist, data_str, periodo=None, alertas=None, catrimani_data=None, omds_totais=None):
    hist_frag, histdata = secao_historico_ncs(res, hist, data_str, periodo)

    frags, CEL, NCD, DAY, TELA = [], {}, {}, {}, {}
    for u in UNIDADES:
        hist_u = [{"data": h.get("data"),
                   "total": {"cred": round(h.get(u["ogu"], {}).get("cred", 0.0)
                                           + h.get(u["fex"], {}).get("cred", 0.0), 2)}}
                  for h in hist]
        u_hist_items = [it for it in histdata["items"] if it["fav_cod"] in (u["ogu"], u["fex"])]
        frag, cel, ncd, day, tela = conteudo_unidade(res, hist_u, data_str, periodo, u, u_hist_items)
        frags.append(frag); CEL.update(cel); NCD.update(ncd); DAY.update(day); TELA.update(tela)
    
    ranking_frag = secao_comparativo_omds(omds_totais, hist, data_str, periodo) if omds_totais else ""
    if ranking_frag:
        frags.append(ranking_frag)

    catrimani_frag = secao_operacao_catrimani(catrimani_data, data_str, periodo) if catrimani_data else ""
    if catrimani_frag:
        frags.append(catrimani_frag)
    frags.append(hist_frag)

    banner = ""
    if alertas:
        itens = "".join(f"<li>{esc(a)}</li>" for a in alertas)
        banner = f'<div class="banner" role="alert"><b>⚠ Verificação de Consistência:</b><ul>{itens}</ul></div>'
    omds = "".join(
        f'<button class="omds{" on" if i == 0 else ""}" data-key="{u["key"]}" aria-current="{"true" if i == 0 else "false"}" '
        f'title="{esc(u["nome"])}" onclick="trocaOMDS(this)">'
        f'<img src="assets/logos/{u["logo"]}" alt="" loading="lazy" onerror="this.style.display=\'none\'">'
        f'<span>{esc(u["sigla"])}</span></button>'
        for i, u in enumerate(UNIDADES))
    omds += (
        '<button class="omds omds-ranking" data-key="RANKING" aria-current="false" '
        'title="Ranking & Benchmarking de Execução Orçamentária das 9 OMDS da Amazônia" onclick="trocaOMDS(this)">'
        '<span class="hist-icon" aria-hidden="true" style="margin-right:6px;">🏆</span>'
        '<span>Ranking OMDS</span></button>'
    )
    omds += (
        '<button class="omds omds-catr" data-key="CATRIMANI" aria-current="false" '
        'title="Acompanhamento Orçamentário da Operação Catrimani II" onclick="trocaOMDS(this)">'
        '<img src="assets/logos/CATRIMANI.png" alt="" loading="lazy" style="width:20px;height:20px;border-radius:50%;margin-right:6px;" onerror="this.style.display=\'none\'">'
        '<span>🎖️ Operação Catrimani</span></button>'
    )
    omds += (
        '<button class="omds omds-hist" data-key="HISTORICO" aria-current="false" '
        'title="Histórico de Notas de Crédito emitidas e recebidas no Exercício" onclick="trocaOMDS(this)">'
        '<span class="hist-icon" aria-hidden="true">📜</span>'
        '<span>Histórico de NCs</span></button>'
    )
    ujs = json.dumps({u["key"]: {"sigla": u["sigla"], "nome": u["nome"], "ogu": u["ogu"], "fex": u["fex"],
                                 "logo": u["logo"], "accent": u["accent"]} for u in UNIDADES}, ensure_ascii=False)
    u0 = UNIDADES[0]
    posicao = periodo if periodo else (data_str[8:10] + "/" + data_str[5:7] + "/" + data_str[0:4])
    ger = datetime.datetime.now().strftime("%d/%m/%Y às %H:%M")
    celdata_json = json.dumps(CEL, ensure_ascii=False).replace("</", "<\\/")
    ncdata_json = json.dumps(NCD, ensure_ascii=False).replace("</", "<\\/")
    daydata_json = json.dumps(DAY, ensure_ascii=False).replace("</", "<\\/")
    teladata_json = json.dumps(TELA, ensure_ascii=False).replace("</", "<\\/")
    histdata_json = json.dumps(histdata, ensure_ascii=False).replace("</", "<\\/")
    catrimani_json = json.dumps(catrimani_data or {}, ensure_ascii=False).replace("</", "<\\/")
    omds_json = json.dumps(omds_totais or {}, ensure_ascii=False).replace("</", "<\\/")
    return f"""<!doctype html><html lang="pt-BR" style="--accent:{u0['accent']}"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light dark">
<meta name="description" content="Painel de Execução Orçamentária e Crédito Disponível do 6º Batalhão de Engenharia de Construção (6º BEC) & Operação Catrimani II — SIAFI / Tesouro Gerencial">
<title>Crédito Disponível — 6º BEC & Operação Catrimani II</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&family=Newsreader:ital,opsz,wght@0,6..72,500;0,6..72,600;0,6..72,700;1,6..72,400&display=swap" rel="stylesheet">
<style>{CSS}</style></head>
<body>
<div class="rm12-bar" aria-hidden="true"></div>
<header class="topbar">
  <div class="brand">
    <div class="brasaos-group">
      <img class="brasao brasao-rm" src="assets/logos/12RM.png" alt="Brasão 12ª Região Militar" title="12ª Região Militar — Região Mendonça Furtado" loading="eager">
      <img class="brasao" id="emblema" src="assets/logos/{u0['logo']}" alt="Brasão {esc(u0['sigla'])}" loading="eager">
    </div>
    <div class="brand-text">
      <div class="eschelon-tag" id="uEscalao">12ª REGIÃO MILITAR · COMANDO MILITAR DA AMAZÔNIA</div>
      <h1 id="uTitulo">Crédito Disponível — {esc(u0['sigla'])}</h1>
      <p class="subtitle"><span id="uNome">{esc(u0['nome'])}</span> · Tesouro Gerencial / SIAFI · <span id="uUasg">UASGs {u0['ogu']} (OGU) e {u0['fex']} (FEx)</span></p>
    </div>
  </div>
  <div class="topbar-r">
    <div class="selo-wrap"><span class="selo"><span class="live-dot" aria-hidden="true"></span> Posição {esc(posicao)}</span><span class="selo-delay">⏱ dados com ~24h de defasagem</span></div>
    <button class="theme" id="themeBtn" aria-pressed="false" aria-label="Alternar tema claro/escuro" onclick="bcmsTheme()" title="Alternar tema">
      <svg viewBox="0 0 24 24" class="ic-sun" aria-hidden="true"><circle cx="12" cy="12" r="4.5" style="fill:currentColor"/><g style="stroke:currentColor;stroke-width:1.8;stroke-linecap:round"><path d="M12 2v2.5M12 19.5v2.5M2 12h2.5M19.5 12h2.5M4.93 4.93l1.77 1.77M17.3 17.3l1.77 1.77M19.07 4.93l-1.77 1.77M6.7 17.3l-1.77 1.77"/></g></svg>
      <svg viewBox="0 0 24 24" class="ic-moon" aria-hidden="true"><path d="M20 14.5A8 8 0 019.5 4 8 8 0 1020 14.5z" style="fill:currentColor"/></svg>
    </button>
  </div>
</header>
<nav class="omds-nav" aria-label="Trocar de organização militar"><div class="omds-nav-in">{omds}</div></nav>
<main class="wrap">
  {banner}
  {"".join(frags)}
</main>
<div class="modal" id="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title" aria-hidden="true" onclick="if(event.target===this)bcmsCelClose()">
  <div class="modal-panel">
    <button class="modal-x" aria-label="Fechar" onclick="bcmsCelClose()">✕</button>
    <div id="modal-body"></div>
  </div>
</div>
<footer class="rodape">
  <p class="rodape-brand">⚙ 6º Batalhão de Engenharia de Construção · Operação Catrimani II · Comando Militar da Amazônia</p>
  <p><b>Metodologia:</b> Crédito Disponível = Provisão Recebida − Provisão Concedida − Despesas Empenhadas (saldo líquido não empenhado no Tesouro Gerencial / SIAFI). O detalhe é o saldo real por célula orçamentária (Ação · PI · ND). A aba Catrimani consolida o acompanhamento inter-unidades de todas as UGs executoras da Ação 21EM.</p>
  <p>Fonte: CRÉDITO DISP 160353.xlsx (Tesouro Gerencial / SIAFI) · <b>⏱ Dados com defasagem de aproximadamente 24 horas.</b> · Painel atualizado em {esc(ger)}</p>
  <p style="margin-top:8px;font-size:12px;opacity:0.85;">💻 <b>Desenvolvido por:</b> 2º Sgt De Campos (BCMS / 6º BEC) &nbsp;·&nbsp; 🔍 <b>Auditado por:</b> Seção de Execução Orçamentária &amp; Fiscalização Administrativa (SIAFI / Tesouro Gerencial)</p>
</footer>
<script>var CELDATA={celdata_json};var NCDATA={ncdata_json};var DAYDATA={daydata_json};var TELADATA={teladata_json};var UNIDADES={ujs};var HISTDATA={histdata_json};var CATRDATA={catrimani_json};var OMDSDATA={omds_json};</script>
<script>{JS}</script>
</body></html>"""

# ============ CSS / JS (Constantes Sênior UI/UX) ============
CSS = r"""
/* ==========================================================================
   DESIGN TOKENS & MASTER UI/UX ARCHITECTURE (60fps GPU + WCAG 2.2 AAA)
   ========================================================================== */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  /* Escala Neutra HSL */
  --neutral-0:   #FFFFFF;
  --neutral-50:  #F8FAFC;
  --neutral-100: #F1F5F9;
  --neutral-200: #E2E8F0;
  --neutral-300: #CBD5E1;
  --neutral-400: #94A3B8;
  --neutral-500: #64748B;
  --neutral-600: #475569;
  --neutral-700: #334155;
  --neutral-800: #1E293B;
  --neutral-900: #0F172A;
  --neutral-950: #020617;

  /* Superfícies & Fundo */
  --bg:          var(--neutral-50);
  --bg-surface:  var(--neutral-0);
  --bg-elevated: var(--neutral-0);
  --bg-subtle:   var(--neutral-100);
  --surface:     var(--neutral-0);
  --surface-2:   var(--neutral-100);

  /* Tipografia & Textos */
  --ink:         var(--neutral-900);
  --ink-muted:   var(--neutral-500);
  --ink-soft:    var(--neutral-400);

  /* Bordas */
  --border:        var(--neutral-200);
  --border-strong: var(--neutral-300);
  --border-focus:  #2563EB;

  /* ==========================================
     HERÁLDICA & IDENTIDADE 12ª REGIÃO MILITAR
     ========================================== */
  --rm12-crimson:       #991B1B;
  --rm12-crimson-dark:  #7F1D1D;
  --rm12-crimson-soft:  rgba(153, 27, 27, 0.08);
  --rm12-blue:          #1E3A8A;
  --rm12-blue-dark:     #172554;
  --rm12-blue-soft:     rgba(30, 58, 138, 0.08);
  --rm12-gold:          #D97706;
  --rm12-gold-dark:     #B45309;
  --rm12-gold-soft:     rgba(217, 119, 6, 0.12);
  --rm12-green:         #15803D;
  --rm12-green-dark:    #166534;
  --rm12-green-soft:    rgba(21, 128, 61, 0.10);

  /* Cores de Marca & Primárias */
  --primary:        var(--rm12-blue);
  --primary-strong: var(--rm12-blue-dark);
  --primary-50:     #EFF6FF;
  --primary-600:    #2563EB;

  /* Cores Semânticas de Estado (WCAG AAA) */
  --success:        var(--rm12-green);
  --success-strong: var(--rm12-green-dark);
  --success-main:   #15803D;
  --success-bg:     #ECFDF5;
  --success-border: #A7F3D0;
  --hero-soft:      #F0FDF4;

  --warning:        var(--rm12-gold);
  --warning-ink:    #92400E;
  --warning-main:   #D97706;
  --warning-bg:     #FFFBEB;
  --warning-border: #FDE68A;

  --danger:         var(--rm12-crimson);
  --danger-main:    #DC2626;
  --danger-bg:      #FEF2F2;
  --danger-border:  #FECACA;

  --gold:           var(--rm12-gold);
  --gold-bg:        #FFFDF5;
  --gold-border:    #FDE68A;
  --gold-dark:      #92400E;

  --focus:          #059669;
  --track:          #E2E8F0;
  --ok:             #10B981;
  --bad:            #EF4444;

  /* Estágios Funil */
  --stg1: #1C4A73;
  --stg2: #3B82F6;
  --stg3: #10B981;

  /* Sombras Físicas em Camadas */
  --shadow-xs: 0 1px 2px rgba(15, 23, 42, 0.04);
  --shadow:    0 1px 3px rgba(15, 23, 42, 0.06), 0 1px 2px rgba(15, 23, 42, 0.04);
  --shadow-md: 0 4px 8px -2px rgba(15, 23, 42, 0.08), 0 2px 4px -2px rgba(15, 23, 42, 0.04);
  --shadow-h:  0 10px 20px -3px rgba(15, 23, 42, 0.10), 0 4px 6px -4px rgba(15, 23, 42, 0.05);
  --shadow-lg: 0 20px 25px -5px rgba(15, 23, 42, 0.10), 0 8px 10px -6px rgba(15, 23, 42, 0.04);

  /* Tipografia */
  --sans:  'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  --mono:  'JetBrains Mono', Consolas, monospace;
  --serif: 'Newsreader', Georgia, serif;

  /* Transições Spring */
  --ease-spring: cubic-bezier(0.16, 1, 0.3, 1);
  --ease-smooth: cubic-bezier(0.25, 1, 0.5, 1);
}

/* --- Dark Mode Elegante (OLED + Baixa Fadiga Ocular) --- */
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    color-scheme: dark;
    --bg:          #090E17;
    --bg-surface:  #101926;
    --bg-elevated: #162234;
    --bg-subtle:   #1C2B40;
    --surface:     #101926;
    --surface-2:   #162234;

    --ink:         #F8FAFC;
    --ink-muted:   #94A3B8;
    --ink-soft:    #64748B;

    --border:        rgba(255, 255, 255, 0.08);
    --border-strong: rgba(255, 255, 255, 0.16);
    --border-focus:  #3B82F6;

    --primary:        #60A5FA;
    --primary-strong: #93C5FD;
    --primary-50:     #1E293B;
    --primary-600:    #3B82F6;

    --success:        #34D399;
    --success-strong: #34D399;
    --success-bg:     rgba(5, 150, 105, 0.15);
    --success-border: rgba(5, 150, 105, 0.35);
    --hero-soft:      rgba(5, 150, 105, 0.08);

    --warning:        #FBBF24;
    --warning-ink:    #FBBF24;
    --warning-main:   #F59E0B;
    --warning-bg:     rgba(217, 119, 6, 0.15);
    --warning-border: rgba(217, 119, 6, 0.35);

    --danger:         #F87171;
    --danger-main:    #EF4444;
    --danger-bg:      rgba(220, 38, 38, 0.15);
    --danger-border:  rgba(220, 38, 38, 0.35);

    --gold:           #F59E0B;
    --gold-bg:        rgba(217, 155, 38, 0.15);
    --gold-border:    rgba(217, 155, 38, 0.35);
    --gold-dark:      #FDE68A;

    --focus:          #34D399;
    --track:          #1E293B;
    --ok:             #34D399;
    --bad:            #F87171;

    --stg1: #60A5FA;
    --stg2: #93C5FD;
    --stg3: #34D399;

    --shadow-xs: none;
    --shadow:    0 2px 6px rgba(0, 0, 0, 0.4);
    --shadow-md: 0 4px 12px rgba(0, 0, 0, 0.5);
    --shadow-h:  0 10px 24px rgba(0, 0, 0, 0.6);
  }
}

:root[data-theme="dark"] {
  color-scheme: dark;
  --bg:          #090E17;
  --bg-surface:  #101926;
  --bg-elevated: #162234;
  --bg-subtle:   #1C2B40;
  --surface:     #101926;
  --surface-2:   #162234;

  --ink:         #F8FAFC;
  --ink-muted:   #94A3B8;
  --ink-soft:    #64748B;

  --border:        rgba(255, 255, 255, 0.08);
  --border-strong: rgba(255, 255, 255, 0.16);
  --border-focus:  #3B82F6;

  --primary:        #60A5FA;
  --primary-strong: #93C5FD;
  --primary-50:     #1E293B;
  --primary-600:    #3B82F6;

  --success:        #34D399;
  --success-strong: #34D399;
  --success-main:   #10B981;
  --success-bg:     rgba(5, 150, 105, 0.15);
  --success-border: rgba(5, 150, 105, 0.35);
  --hero-soft:      rgba(5, 150, 105, 0.08);

  --warning:        #FBBF24;
  --warning-ink:    #FBBF24;
  --warning-main:   #F59E0B;
  --warning-bg:     rgba(217, 119, 6, 0.15);
  --warning-border: rgba(217, 119, 6, 0.35);

  --danger:         #F87171;
  --danger-main:    #EF4444;
  --danger-bg:      rgba(220, 38, 38, 0.15);
  --danger-border:  rgba(220, 38, 38, 0.35);

  --gold:           #F59E0B;
  --gold-bg:        rgba(217, 155, 38, 0.15);
  --gold-border:    rgba(217, 155, 38, 0.35);
  --gold-dark:      #FDE68A;

  --focus:          #34D399;
  --track:          #1E293B;
  --ok:             #34D399;
  --bad:            #F87171;

  --stg1: #60A5FA;
  --stg2: #93C5FD;
  --stg3: #34D399;

  --shadow-xs: none;
  --shadow:    0 2px 6px rgba(0, 0, 0, 0.4);
  --shadow-md: 0 4px 12px rgba(0, 0, 0, 0.5);
  --shadow-h:  0 10px 24px rgba(0, 0, 0, 0.6);
}

:root[data-theme="light"] {
  color-scheme: light;
  --bg:          var(--neutral-50);
  --bg-surface:  var(--neutral-0);
  --bg-elevated: var(--neutral-0);
  --bg-subtle:   var(--neutral-100);
  --surface:     var(--neutral-0);
  --surface-2:   var(--neutral-100);
  --ink:         var(--neutral-900);
  --ink-muted:   var(--neutral-500);
  --border:        var(--neutral-200);
  --border-strong: var(--neutral-300);
  --hero-soft:      #F0FDF4;
}

html {
  transition: background-color .2s var(--ease-smooth), color .2s var(--ease-smooth);
  font-size: 16px;
  scroll-behavior: smooth;
}

body {
  background: var(--bg);
  color: var(--ink);
  font-family: var(--sans);
  font-size: 0.9375rem;
  line-height: 1.55;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}

.num {
  font-family: var(--mono);
  font-variant-numeric: tabular-nums lining-nums;
  font-feature-settings: "tnum" 1, "lnum" 1;
}

.visually-hidden {
  position: absolute;
  width: 1px; height: 1px;
  overflow: hidden;
  clip: rect(0 0 0 0);
}

.wrap {
  max-width: 1280px;
  margin: 0 auto;
  padding: 0 24px 64px;
}

.sec { margin-top: 36px; }
.eyebrow {
  font-size: 0.75rem;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.12em;
  color: var(--ink-muted);
  margin-bottom: 12px;
}

.sec-nota {
  font-size: 0.8125rem;
  color: var(--ink-muted);
  margin: -4px 0 16px;
  max-width: 960px;
  line-height: 1.6;
}
.sec-nota b { color: var(--ink); font-weight: 600; }

/* Botões Excel Profissionais */
.btn-excel {
  position: relative;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  background: linear-gradient(135deg, #107C41 0%, #0B5E31 100%);
  color: #FFFFFF !important;
  border: 1px solid #0E6B38;
  border-radius: 8px;
  padding: 0 16px;
  min-height: 48px;
  font-family: var(--sans);
  font-size: 0.8125rem;
  font-weight: 600;
  cursor: pointer;
  box-shadow: 0 2px 4px rgba(16, 124, 65, 0.2);
  transition: transform .2s var(--ease-spring);
  white-space: nowrap;
}
.btn-excel::after {
  content: '';
  position: absolute; inset: -1px;
  border-radius: inherit;
  background: linear-gradient(135deg, #148C4A 0%, #0E733D 100%);
  box-shadow: 0 6px 14px rgba(16, 124, 65, 0.32);
  opacity: 0;
  pointer-events: none;
  transition: opacity .2s var(--ease-smooth);
}
.btn-excel:hover {
  transform: translateY(-2px);
}
.btn-excel:hover::after {
  opacity: 1;
}
.btn-excel:active {
  transform: translateY(0) scale(0.97);
}
.btn-excel:active::after {
  opacity: 0;
}
.btn-excel-ic { font-size: 1rem; line-height: 1; }
.btn-excel-lg { padding: 10px 20px; font-size: 0.875rem; border-radius: 10px; }

/* Abas de Topo (Segmented Control Fluido) */
.toptabs {
  display: flex;
  gap: 6px;
  margin: 24px 0 8px;
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 5px;
}
.toptab {
  flex: 1;
  background: transparent;
  border: none;
  color: var(--ink-muted);
  font-family: var(--sans);
  font-size: 0.875rem;
  font-weight: 600;
  padding: 10px 16px;
  min-height: 48px;
  border-radius: 8px;
  cursor: pointer;
  transition: color .2s var(--ease-out-expo), background-color .2s var(--ease-out-expo), box-shadow .2s var(--ease-spring);
}
.toptab.on {
  background: var(--bg-surface);
  color: var(--ink);
  box-shadow: var(--shadow-md);
  font-weight: 700;
}
.toptab:hover:not(.on) { color: var(--ink); background: rgba(0,0,0,0.03); }
:root[data-theme="dark"] .toptab:hover:not(.on) { background: rgba(255,255,255,0.04); }

/* Topbar & Header 12ª RM */
.rm12-bar {
  height: 5px;
  background: linear-gradient(
    to right,
    var(--rm12-crimson) 0%,
    var(--rm12-crimson) 24%,
    var(--rm12-gold) 24%,
    var(--rm12-gold) 28%,
    var(--rm12-blue) 28%,
    var(--rm12-blue) 66%,
    var(--rm12-gold) 66%,
    var(--rm12-gold) 68%,
    var(--rm12-green) 68%,
    var(--rm12-green) 100%
  );
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.2);
}
.bcms-bar { display: none; }
.topbar {
  position: sticky;
  top: 0;
  z-index: 40;
  background: color-mix(in srgb, var(--bg-surface) 90%, transparent);
  backdrop-filter: blur(14px);
  -webkit-backdrop-filter: blur(14px);
  border-bottom: 1px solid var(--border);
  padding: 14px 24px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 16px;
}
.brand { display: flex; align-items: center; gap: 14px; }
.brasaos-group {
  display: flex;
  align-items: center;
  gap: 8px;
  flex: none;
}
.brasao {
  width: 48px;
  height: 48px;
  object-fit: contain;
  flex: none;
  filter: drop-shadow(0 2px 5px rgba(0, 0, 0, 0.16));
  transition: transform 0.2s var(--ease-spring);
}
.brasao:hover {
  transform: scale(1.06);
}
.brasao-rm {
  border-radius: 4px;
}
.brand-text { display: flex; flex-direction: column; }
.eschelon-tag {
  font-size: 0.6875rem;
  font-weight: 800;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: var(--rm12-gold);
  margin-bottom: 2px;
  display: inline-flex;
  align-items: center;
  gap: 6px;
}
.eschelon-tag::before {
  content: "";
  display: inline-block;
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: var(--rm12-crimson);
}
h1 {
  font-family: var(--serif);
  font-size: clamp(1.25rem, 1.1rem + 0.8vw, 1.75rem);
  font-weight: 700;
  letter-spacing: -0.01em;
  line-height: 1.15;
  color: var(--ink);
}
.subtitle {
  font-size: 0.78125rem;
  color: var(--ink-muted);
  margin-top: 3px;
}
.topbar-r { display: flex; align-items: center; gap: 12px; }
.selo-wrap { display: flex; flex-direction: column; align-items: flex-end; gap: 3px; }
.selo {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  background: var(--bg-subtle);
  color: var(--ink-muted);
  border: 1px solid var(--border);
  border-radius: 999px;
  padding: 5px 12px;
  font-size: 0.78125rem;
  font-weight: 600;
  white-space: nowrap;
}
.live-dot {
  width: 7px;
  height: 7px;
  border-radius: 50%;
  background: var(--success);
  box-shadow: 0 0 8px var(--success);
}
.selo-delay { font-size: 0.6875rem; color: var(--warning-ink); font-weight: 600; white-space: nowrap; }

/* Botão Tema */
.theme {
  width: 48px;
  height: 48px;
  border: 1px solid var(--border);
  background: var(--bg-surface);
  border-radius: 9px;
  color: var(--ink-muted);
  cursor: pointer;
  display: grid;
  place-items: center;
  transition: border-color .2s var(--ease-out-expo), color .2s var(--ease-out-expo), transform .2s var(--ease-spring);
}
.theme:hover { border-color: var(--border-strong); color: var(--ink); transform: scale(1.05); }
.theme svg { width: 18px; height: 18px; }
.ic-moon { display: none; }
:root[data-theme="dark"] .ic-sun, html:not([data-theme]) .ic-sun { display: block; }
:root[data-theme="dark"] .ic-moon { display: block; }
:root[data-theme="dark"] .ic-sun { display: none; }
@media (prefers-color-scheme: dark) {
  html:not([data-theme]) .ic-sun { display: none; }
  html:not([data-theme]) .ic-moon { display: block; }
}

/* Barra de Navegação OMDS */
.omds-nav {
  background: var(--bg-surface);
  border-bottom: 1px solid var(--border);
  position: relative;
  z-index: 30;
}
.omds-nav-in {
  max-width: 1280px;
  margin: 0 auto;
  padding: 10px 24px;
  display: flex;
  gap: 10px;
  overflow-x: auto;
  -webkit-overflow-scrolling: touch;
  scrollbar-width: thin;
}
.omds {
  flex: none;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 7px 15px 7px 9px;
  min-height: 48px;
  border: 1px solid var(--border);
  border-radius: 999px;
  background: var(--bg-subtle);
  color: var(--ink-muted);
  font-family: var(--sans);
  font-size: 0.8125rem;
  font-weight: 700;
  cursor: pointer;
  white-space: nowrap;
  transition: transform .2s var(--ease-spring), border-color .2s var(--ease-out-expo), color .2s var(--ease-out-expo), background-color .2s var(--ease-out-expo);
}
.omds img { width: 22px; height: 22px; object-fit: contain; flex: none; }
.omds:hover { border-color: var(--border-strong); color: var(--ink); transform: translateY(-1px); }
.omds.on {
  color: #FFFFFF;
  background: var(--accent, #CE2B2B);
  border-color: var(--accent, #CE2B2B);
  box-shadow: 0 4px 12px color-mix(in srgb, var(--accent, #CE2B2B) 40%, transparent);
}
.omds.on img { filter: drop-shadow(0 0 2px rgba(255,255,255,0.7)); }

.omds-rank {
  background: linear-gradient(135deg, var(--bg-subtle) 0%, var(--gold-bg) 100%);
  border-color: var(--gold-border);
  color: var(--gold-dark);
}
.omds-rank.on {
  background: linear-gradient(135deg, #D99B26 0%, #B47810 100%) !important;
  border-color: #B47810 !important;
  color: #FFFFFF !important;
  box-shadow: 0 4px 14px rgba(217, 155, 38, 0.45);
}
.rank-icon { font-size: 0.9375rem; }

/* Cards & Superfícies */
.card {
  position: relative;
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 20px;
  box-shadow: var(--shadow);
  transition: border-color .2s var(--ease-out-expo);
}
.card:hover { border-color: var(--border-strong); }

/* Hero Card Moderno */
.hero {
  margin-top: 24px;
  background: var(--hero-soft);
  border: 1px solid var(--border);
  border-left: 5px solid var(--success);
  border-radius: 16px;
  padding: 24px 28px;
  display: grid;
  grid-template-columns: 1.35fr 1fr;
  gap: 28px;
  align-items: center;
  box-shadow: var(--shadow-md);
}
.hero-cmd {
  border-left-color: var(--gold);
  background: var(--gold-bg);
}
.hero-num {
  font-size: clamp(2.2rem, 1.8rem + 2vw, 3.25rem);
  font-weight: 800;
  letter-spacing: -0.025em;
  color: var(--success-strong);
  line-height: 1.05;
  margin: 6px 0 16px;
}
.hero-cmd .hero-num { color: var(--gold-dark); }

.hero-eq {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px;
}
.hero-eq-box {
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 6px 12px;
  display: flex;
  flex-direction: column;
  gap: 2px;
  box-shadow: var(--shadow-xs);
}
.hero-eq-box .eq-tag {
  font-size: 0.625rem;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--ink-muted);
}
.hero-eq-box .eq-val {
  font-size: 0.9375rem;
  font-weight: 700;
  color: var(--ink);
}
.hero-eq-box.eq-highlight {
  border-color: var(--success-border);
  background: var(--success-bg);
}
.hero-eq-box.eq-highlight .eq-val { color: var(--success-strong); }
.hero-eq-sign {
  font-size: 1.25rem;
  font-weight: 300;
  color: var(--ink-muted);
  padding: 0 2px;
}

.delta {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 0.8125rem;
  font-weight: 600;
  margin-bottom: 12px;
  padding: 4px 12px;
  border-radius: 999px;
  background: var(--bg-subtle);
  border: 1px solid var(--border);
}
.delta.up { color: var(--success-strong); }
.delta.down { color: var(--danger); }
.delta.flat { color: var(--ink-muted); }
.delta small { font-weight: 400; color: var(--ink-muted); }

/* KPIs Grid */
.kpis {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 16px;
}
.kpi {
  position: relative;
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-left: 4px solid var(--border);
  border-radius: 12px;
  padding: 16px 18px;
  box-shadow: var(--shadow);
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  transition: transform .2s var(--ease-spring);
}
.kpi::after {
  content: '';
  position: absolute;
  inset: -1px;
  border-radius: inherit;
  box-shadow: var(--shadow-md);
  opacity: 0;
  pointer-events: none;
  transition: opacity .2s var(--ease-out-expo);
}
.kpi:hover { transform: translateY(-2px); }
.kpi:hover::after { opacity: 1; }
.kpi-total { border-left-color: #6366F1; }
.kpi-prov { border-left-color: var(--primary-600); }
.kpi-emp  { border-left-color: var(--warning-main); }
.kpi-liq  { border-left-color: var(--stg2); }
.kpi-pag  { border-left-color: var(--success); }
.kpi-l {
  font-size: 0.75rem;
  font-weight: 700;
  letter-spacing: 0.05em;
  color: var(--ink-muted);
  text-transform: uppercase;
}
.kpi-v {
  font-size: 1.5rem;
  font-weight: 700;
  margin: 6px 0;
  color: var(--ink);
}
.chip {
  display: inline-block;
  font-size: 0.6875rem;
  font-weight: 600;
  color: var(--ink-muted);
  background: var(--bg-subtle);
  border-radius: 999px;
  padding: 2px 10px;
  width: fit-content;
}

/* Grids */
.grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }

/* Cards de UASG */
.uasg { display: flex; flex-direction: column; gap: 10px; }
.uasg-h { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.uasg-cod { font-size: 1rem; font-weight: 700; }
.uasg-nome { font-size: 0.8125rem; color: var(--ink-muted); }
.uasg-disp { display: flex; justify-content: space-between; align-items: baseline; margin-top: 4px; }
.uasg-disp-l { font-size: 0.75rem; font-weight: 600; text-transform: uppercase; color: var(--ink-muted); }
.uasg-disp-v { font-size: 1.375rem; font-weight: 800; color: var(--success-strong); }
.uasg-eq { font-size: 0.78125rem; color: var(--ink-muted); }
.uasg-eq i { font-style: normal; color: var(--warning-ink); font-weight: 600; }
.uasg-exec { margin-top: 4px; }
.exec-l { display: flex; justify-content: space-between; font-size: 0.75rem; color: var(--ink-muted); margin-bottom: 5px; font-weight: 600; }
.exec-track { height: 8px; background: var(--track); border-radius: 999px; overflow: hidden; }
.exec-fill { height: 100%; background: var(--primary-600); border-radius: 999px; }

/* Módulo Créditos em Tela por NC */
.et-head {
  display: flex;
  flex-wrap: wrap;
  gap: 14px;
  align-items: stretch;
  margin-bottom: 16px;
}
.et-kpi {
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 14px 18px;
  display: flex;
  flex-direction: column;
  gap: 3px;
  min-width: 140px;
  box-shadow: var(--shadow);
}
.et-kpi span { font-size: 0.6875rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-muted); }
.et-kpi b { font-size: 1.25rem; font-weight: 700; color: var(--ink); }
.et-hero { background: var(--hero-soft); border-left: 4px solid var(--success); }
.et-hero b { font-size: 1.5rem; color: var(--success-strong); }
.et-action { display: flex; align-items: center; }
.et-meta { margin-left: auto; align-self: center; font-size: 0.75rem; color: var(--ink-muted); text-align: right; line-height: 1.5; }

/* Badges de Idade de Crédito (Semáforo) */
.badge-age {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 6px;
  font-size: 0.75rem;
  font-weight: 700;
  font-family: var(--mono);
}
.age-green { background: var(--success-bg); color: var(--success-strong); border: 1px solid var(--success-border); }
.age-amber { background: var(--warning-bg); color: var(--warning-ink); border: 1px solid var(--warning-border); }
.age-red   { background: var(--danger-bg);  color: var(--danger);       border: 1px solid var(--danger-border); }
.age-none  { color: var(--ink-muted); }

/* Pílula de Fonte (160 / 167) */
.pill-fonte {
  display: inline-block;
  font-size: 0.6875rem;
  font-weight: 700;
  font-family: var(--mono);
  color: var(--primary);
  background: var(--bg-subtle);
  border: 1px solid var(--border-strong);
  border-radius: 5px;
  padding: 1px 7px;
}

/* Tabelas e Ferramentas */
.tabs {
  display: flex;
  gap: 6px;
  border-bottom: 1px solid var(--border);
  margin-bottom: 16px;
  overflow-x: auto;
}
.tab {
  background: none;
  border: none;
  border-bottom: 2px solid transparent;
  color: var(--ink-muted);
  font-family: var(--sans);
  font-size: 0.8125rem;
  font-weight: 600;
  padding: 10px 16px;
  cursor: pointer;
  white-space: nowrap;
  transition: transform .2s var(--ease-out-expo), color .2s var(--ease-out-expo), background-color .2s var(--ease-out-expo), border-color .2s var(--ease-out-expo);
}
.tab.on { color: var(--success-strong); border-bottom-color: var(--success); font-weight: 700; }
.tab:hover:not(.on) { color: var(--ink); }

.tbl-tools {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 12px;
  flex-wrap: wrap;
}
.tbl-search {
  flex: 1;
  min-width: 240px;
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 10px 14px;
  color: var(--ink);
  font-family: var(--sans);
  font-size: 0.875rem;
  box-shadow: var(--shadow-xs);
  transition: border-color .2s var(--ease-out-expo), box-shadow .2s var(--ease-out-expo);
}
.tbl-search:focus {
  outline: 2px solid var(--border-focus);
  outline-offset: 1px;
  border-color: var(--border-focus);
}
.tbl-count { font-size: 0.78125rem; color: var(--ink-muted); margin-left: auto; font-weight: 500; }
/* ==========================================================================
   SISTEMA DE TABELAS & ALINHAMENTO COGNITIVO (WCAG 2.2 AAA + TUFTE DATA-VIZ)
   ========================================================================== */
.tbl-wrap, .tbl-scroll {
  overflow-x: auto;
  -webkit-overflow-scrolling: touch;
  border: 1px solid var(--border);
  border-radius: 12px;
  background: var(--bg-surface);
  box-shadow: var(--shadow);
  margin: 16px 0;
  max-width: 100%;
}

table.tbl, table.det, table.tbl-hist {
  border-collapse: separate;
  border-spacing: 0;
  width: 100%;
  min-width: 680px;
  font-size: 0.75rem;
  line-height: 1.2;
}

table.tbl th, table.det th, table.tbl-hist th {
  position: sticky;
  top: 0;
  z-index: 10;
  background: var(--bg-subtle);
  color: var(--ink-muted);
  font-size: 0.65rem;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  text-align: left;
  padding: 4px 6px;
  white-space: nowrap;
  border-bottom: 2px solid var(--border);
  user-select: none;
}

table.tbl th.num, table.det th.num, table.tbl-hist th.num {
  text-align: right;
}

table.tbl th .sort, table.det th .sort, table.tbl-hist th .sort {
  display: inline-block;
  width: 12px;
  color: var(--rm12-gold);
}

table.tbl td, table.det td, table.tbl-hist td {
  padding: 4px 6px;
  border-bottom: 1px solid var(--border);
  color: var(--ink);
  vertical-align: middle;
  font-size: 0.75rem;
  line-height: 1.2;
}

table.tbl td.num, table.det td.num, table.tbl-hist td.num {
  text-align: right;
  font-family: var(--mono);
  font-variant-numeric: tabular-nums;
  white-space: nowrap;
  font-size: 0.75rem;
  letter-spacing: -0.015em;
  padding-left: 6px;
  padding-right: 6px;
  min-width: 85px;
}

.kpi {
  cursor: pointer;
  transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.kpi:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(0,0,0,0.12);
}
.tr-click, .cel-row, table.det tr:not(.det-hdr) {
  cursor: pointer;
  transition: background 0.12s ease;
}
.tr-click:hover, .cel-row:hover, table.det tr:not(.det-hdr):hover {
  background: rgba(21, 128, 61, 0.08) !important;
}
td.num.anchor, td.col-moeda.anchor {
  cursor: pointer;
  color: var(--primary);
  font-weight: 700;
  text-decoration: underline dotted;
}
td.num.anchor:hover, td.col-moeda.anchor:hover {
  color: #15803D;
  background: rgba(21, 128, 61, 0.12);
}

table.tbl td.col-ug, table.tbl th.col-ug {
  min-width: 75px;
  white-space: nowrap;
}

table.tbl td.col-nome, table.tbl th.col-nome {
  min-width: 180px;
}

table.tbl td.col-moeda, table.tbl th.col-moeda,
table.det td.col-moeda, table.det th.col-moeda {
  min-width: 115px;
}

table.tbl td.col-pct, table.tbl th.col-pct {
  min-width: 135px;
  padding-right: 12px;
}

.bar-pct-wrap {
  display: flex;
  align-items: center;
  gap: 8px;
  justify-content: flex-end;
}

.bar-pct-track {
  flex: 1;
  height: 8px;
  background: var(--neutral-200, #E2E8F0);
  border-radius: 4px;
  overflow: hidden;
  min-width: 60px;
  border: 1px solid rgba(0, 0, 0, 0.06);
}

:root[data-theme="dark"] .bar-pct-track {
  background: rgba(255, 255, 255, 0.12);
  border: 1px solid rgba(255, 255, 255, 0.08);
}

.bar-pct-fill {
  height: 100%;
  border-radius: 4px;
  transition: width 0.3s ease;
  min-width: 2px;
}

.bar-pct-val {
  font-weight: 700;
  font-size: 0.8125rem;
  font-family: var(--mono);
  min-width: 44px;
  text-align: right;
}

table.tbl td.anchor, table.det td.anchor, table.tbl-hist td.anchor,
table.tbl th.anchor, table.det th.anchor, table.tbl-hist th.anchor,
table.tbl tfoot td.anchor, table.det tfoot td.anchor, table.tbl-hist tfoot td.anchor {
  font-weight: 700;
  color: var(--rm12-green, #15803D);
  cursor: pointer;
}

:root[data-theme="dark"] table.tbl td.anchor,
:root[data-theme="dark"] table.det td.anchor,
:root[data-theme="dark"] table.tbl-hist td.anchor,
:root[data-theme="dark"] table.tbl th.anchor,
:root[data-theme="dark"] table.det th.anchor,
:root[data-theme="dark"] table.tbl-hist th.anchor,
:root[data-theme="dark"] table.tbl tfoot td.anchor,
:root[data-theme="dark"] table.det tfoot td.anchor,
:root[data-theme="dark"] table.tbl-hist tfoot td.anchor {
  color: #34D399;
}

table.tbl tbody tr:hover, table.det tbody tr:hover, table.tbl-hist tbody tr:hover,
tr.tr-click:hover, tr.cel-row:hover {
  background: rgba(37, 99, 235, 0.07) !important;
}

:root[data-theme="dark"] table.tbl tbody tr:hover,
:root[data-theme="dark"] table.det tbody tr:hover,
:root[data-theme="dark"] table.tbl-hist tbody tr:hover,
:root[data-theme="dark"] tr.tr-click:hover,
:root[data-theme="dark"] tr.cel-row:hover {
  background: rgba(59, 130, 246, 0.12) !important;
}

tr.tr-click, tr.cel-row {
  cursor: pointer;
}

table.tbl tfoot td, table.det tfoot td, table.tbl-hist tfoot td {
  padding: 8px 12px;
  font-weight: 800;
  background: var(--bg-subtle);
  border-top: 2px solid var(--border-strong);
  color: var(--ink);
  font-size: 0.8125rem;
}

table.tbl tfoot td.tfoot-label {
  padding-right: 20px;
  white-space: nowrap;
}

table.tbl tfoot td.num, table.det tfoot td.num, table.tbl-hist tfoot td.num {
  font-family: var(--mono);
  font-variant-numeric: tabular-nums;
  white-space: nowrap;
  text-align: right;
  min-width: 110px;
}
.det .obj { max-width: 450px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink-muted); }
.det-compact td { padding: 4px 8px; line-height: 1.25; }
.det-compact th { padding: 6px 8px; }
.det-compact tfoot td { padding: 6px 8px; }
.det-compact .obj { max-width: 500px; color: var(--ink); }
.det-compact .mono2 { font-size: 0.78125rem; white-space: nowrap; }
.det-compact .badge-age { padding: 1px 6px; font-size: 0.6875rem; white-space: nowrap; }
.det-compact .pill-fonte { padding: 1px 6px; font-size: 0.6875rem; white-space: nowrap; }
.nc-num { font-weight: 700; letter-spacing: .01em; white-space: nowrap; }
.nc-ug { color: var(--ink-muted); font-weight: 500; white-space: nowrap; }
.nc-lbl-wrap { display: inline-flex; align-items: center; gap: 4px; white-space: nowrap; }
/* selo de crédito vindo de mudança de ND (objeto herdado da NC de origem) */
.tag-nd {
  display: inline-block; padding: 0 6px; margin-right: 5px; border-radius: 5px;
  background: var(--warning-bg, rgba(181,130,43,.14)); color: var(--warning-ink, #8A631C);
  font-size: 0.6875rem; font-weight: 700; white-space: nowrap; vertical-align: 1px;
}
/* ---- barra de filtros da lista em tela ---- */
.tbl-filtros {
  display: flex; flex-wrap: wrap; align-items: center; gap: 8px;
  padding: 10px 12px; margin-bottom: 10px; border-radius: 10px;
  background: var(--surface-2); border: 1px solid var(--border);
}
.flt-lbl { font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-muted); }
.flt {
  font-family: inherit; font-size: 0.8125rem; color: var(--ink);
  background: var(--surface); border: 1px solid var(--border-strong);
  border-radius: 8px; padding: 6px 10px; max-width: 240px; cursor: pointer;
}
.flt:focus-visible { outline: 2px solid var(--focus); outline-offset: 1px; }

/* Contraste estrito e compatibilidade para dropdowns/selects em tema escuro e claro */
select, .flt, .hist-select {
  color-scheme: light dark;
}
.flt option, .flt optgroup,
.hist-select option, .hist-select optgroup,
select option, select optgroup {
  background-color: var(--bg-surface);
  color: var(--ink);
}
:root[data-theme="dark"] select,
:root[data-theme="dark"] select option,
:root[data-theme="dark"] select optgroup,
:root[data-theme="dark"] .flt,
:root[data-theme="dark"] .flt option,
:root[data-theme="dark"] .flt optgroup,
:root[data-theme="dark"] .hist-select,
:root[data-theme="dark"] .hist-select option,
:root[data-theme="dark"] .hist-select optgroup,
:root:not([data-theme="light"]) select,
:root:not([data-theme="light"]) select option,
:root:not([data-theme="light"]) select optgroup,
:root:not([data-theme="light"]) .flt,
:root:not([data-theme="light"]) .flt option,
:root:not([data-theme="light"]) .flt optgroup,
:root:not([data-theme="light"]) .hist-select,
:root:not([data-theme="light"]) .hist-select option,
:root:not([data-theme="light"]) .hist-select optgroup {
  color-scheme: dark !important;
  background-color: #101926 !important;
  color: #F8FAFC !important;
}
:root[data-theme="light"] select,
:root[data-theme="light"] select option,
:root[data-theme="light"] select optgroup,
:root[data-theme="light"] .flt,
:root[data-theme="light"] .flt option,
:root[data-theme="light"] .flt optgroup,
:root[data-theme="light"] .hist-select,
:root[data-theme="light"] .hist-select option,
:root[data-theme="light"] .hist-select optgroup {
  color-scheme: light !important;
  background-color: #FFFFFF !important;
  color: #0F172A !important;
}
select option:checked, .flt option:checked, .hist-select option:checked {
  background-color: #2563EB !important;
  color: #FFFFFF !important;
}
.flt-limpa {
  font-family: inherit; font-size: 0.75rem; font-weight: 600; color: var(--ink-muted);
  background: transparent; border: 1px solid var(--border-strong); border-radius: 8px;
  padding: 6px 10px; cursor: pointer;
}
.flt-limpa:hover { color: var(--danger); border-color: var(--danger); }
.flt-limpa:disabled, .btn-pag:disabled { opacity: 0.4; cursor: not-allowed; pointer-events: none; }
.flt-resumo { font-size: 0.8125rem; color: var(--ink-muted); font-weight: 600; }
.flt-resumo.on { color: var(--primary); }
@media (max-width: 640px) { .flt { flex: 1 1 100%; max-width: none; } }
/* ---- modal por etapas do crédito ---- */
.m-etapas { display: flex; flex-wrap: wrap; gap: 6px; margin: 14px 0 4px; border-bottom: 1px solid var(--border); padding-bottom: 10px; }
.m-etapa {
  border: 1px solid var(--border); background: var(--bg-subtle); color: var(--ink-muted);
  border-radius: 999px; padding: 6px 14px; font-size: 0.8125rem; font-weight: 600;
  cursor: pointer; transition: transform .2s var(--ease-out-expo), color .2s var(--ease-out-expo), background-color .2s var(--ease-out-expo), border-color .2s var(--ease-out-expo); font-family: inherit;
}
.m-etapa:hover { color: var(--ink); border-color: var(--border-strong); }
.m-etapa.on { background: var(--primary); border-color: var(--primary); color: #fff; }
.m-etapa-body { padding-top: 14px; }
.m-tela-card {
  display: flex; flex-direction: column; gap: 4px; padding: 18px 20px; border-radius: 14px;
  background: var(--hero-soft, rgba(15,122,90,.10)); border: 1px solid var(--success, #0F7A5A);
}
.m-tela-lbl { font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: var(--success-strong, #0B5E45); }
.m-tela-val { font-size: 2rem; font-weight: 800; line-height: 1.1; font-variant-numeric: tabular-nums; color: var(--ink); }
.m-tela-meta { font-size: 0.8125rem; color: var(--ink-muted); }
.m-org { margin-top: 14px; border: 1px dashed var(--warning, #B5822B); border-radius: 12px; overflow: hidden; }
.m-org-h { padding: 8px 14px; background: var(--warning-bg, rgba(181,130,43,.12)); font-size: 0.8125rem; color: var(--warning-ink, #8A631C); }
.m-org-b { padding: 10px 14px; }
.m-org-b span { font-size: 0.75rem; color: var(--ink-muted); }
.m-org-b p { margin-top: 4px; font-size: 0.875rem; color: var(--ink); }
.m-nc-desc.big { font-size: 0.9375rem; line-height: 1.5; color: var(--ink); }
.m-nc-esta { border-color: var(--primary); box-shadow: 0 0 0 1px var(--primary) inset; }
.m-nc-tag { font-style: normal; font-size: 0.625rem; font-weight: 700; text-transform: uppercase;
  background: var(--primary); color: #fff; padding: 1px 6px; border-radius: 4px; margin-left: 6px; }
.m-barra { margin: 14px 0 4px; }
.m-barra-t { display: flex; justify-content: space-between; font-size: 0.8125rem; margin-bottom: 5px; color: var(--ink-muted); }
.m-barra-t b { color: var(--ink); font-variant-numeric: tabular-nums; }
.m-barra-track { height: 10px; border-radius: 999px; background: var(--track, #E4E8EF); overflow: hidden; }
.m-barra-fill { height: 100%; border-radius: 999px; background: var(--success, #0F7A5A); }
.det .mono2 { font-size: 0.8125rem; color: var(--ink-muted); }
.cell-neg { color: var(--danger); background: var(--danger-bg); }
.det tfoot td { padding: 12px 14px; font-weight: 700; background: var(--bg-subtle); border-top: 2px solid var(--border-strong); }
.cel-row { cursor: pointer; }
.cel-row .chev { float: right; margin-left: 8px; color: var(--ink-soft); font-weight: 400; transition: transform .15s ease, color .15s ease; }
.cel-row:hover .chev { color: var(--success); transform: translateX(3px); }

/* Pódio & Benchmarking (Ranking OMDS) */
.ranking-header-card {
  margin-top: 24px;
  padding: 26px 30px;
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-top: 5px solid var(--gold);
  border-radius: 16px;
  box-shadow: var(--shadow-md);
}
.rh-tag { font-size: 0.6875rem; font-weight: 800; letter-spacing: 0.14em; color: var(--gold-dark); margin-bottom: 6px; }
.rh-title { font-family: var(--serif); font-size: clamp(1.4rem, 1.2rem + 1vw, 2rem); font-weight: 700; margin-bottom: 8px; line-height: 1.2; }
.rh-desc { font-size: 0.875rem; color: var(--ink-muted); max-width: 900px; line-height: 1.6; }

.podium-wrap {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 20px;
  margin-top: 18px;
  align-items: end;
}
.podium-step {
  position: relative;
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 24px 20px 20px;
  text-align: center;
  cursor: pointer;
  box-shadow: var(--shadow);
  transition: opacity .25s var(--ease-out-expo), transform .25s var(--ease-spring);
  display: flex;
  flex-direction: column;
  align-items: center;
}
.podium-step:hover {
  transform: translateY(-6px);
  box-shadow: var(--shadow-h);
  border-color: var(--border-strong);
}
.podium-gold {
  border-top: 6px solid #F59E0B;
  background: linear-gradient(to bottom, var(--gold-bg), var(--bg-surface));
  order: 2;
  padding-top: 32px;
  margin-bottom: 16px;
}
.podium-silver {
  border-top: 6px solid #94A3B8;
  background: linear-gradient(to bottom, var(--bg-subtle), var(--bg-surface));
  order: 1;
}
.podium-bronze {
  border-top: 6px solid #CD7F32;
  background: linear-gradient(to bottom, color-mix(in srgb, #CD7F32 8%, var(--bg-surface)), var(--bg-surface));
  order: 3;
}
.podium-badge {
  display: inline-block;
  font-size: 0.75rem;
  font-weight: 800;
  padding: 4px 14px;
  border-radius: 999px;
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  margin-bottom: 14px;
}
.podium-gold .podium-badge { background: #FFF5D6; color: #855A00; border-color: #F0D070; }
:root[data-theme="dark"] .podium-gold .podium-badge { background: #3B2C08; color: #F0D070; border-color: #6E5110; }

.podium-avatar-wrap {
  width: 68px;
  height: 68px;
  border-radius: 50%;
  background: var(--bg-surface);
  border: 2px solid var(--border);
  display: flex;
  align-items: center;
  justify-content: center;
  margin-bottom: 12px;
  overflow: hidden;
  padding: 6px;
  box-shadow: var(--shadow-xs);
}
.podium-gold .podium-avatar-wrap {
  width: 80px;
  height: 80px;
  border-color: #F59E0B;
  box-shadow: 0 0 20px rgba(245, 158, 11, 0.35);
}
.podium-logo { width: 100%; height: 100%; object-fit: contain; }
.podium-sigla { font-size: 1.1875rem; font-weight: 800; letter-spacing: -0.01em; margin-bottom: 2px; }
.podium-nome { font-size: 0.75rem; color: var(--ink-muted); margin-bottom: 16px; max-width: 210px; line-height: 1.35; height: 32px; overflow: hidden; }
.podium-stat-pill {
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 10px 16px;
  width: 100%;
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 10px;
}
.podium-stat-pill .stat-l { font-size: 0.6875rem; text-transform: uppercase; color: var(--ink-muted); font-weight: 700; }
.podium-stat-pill .stat-v { font-size: 1.125rem; font-weight: 800; color: var(--success-strong); }
.podium-substat { font-size: 0.75rem; color: var(--ink-muted); margin-bottom: 16px; }
.podium-btn {
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 8px 14px;
  font-family: var(--sans);
  font-size: 0.75rem;
  font-weight: 700;
  color: var(--primary);
  cursor: pointer;
  transition: transform .2s var(--ease-out-expo), background-color .2s var(--ease-out-expo);
  width: 100%;
}
.podium-btn:hover { background: var(--primary); color: #FFFFFF; border-color: var(--primary); }

/* Tabela Comparativa */
.tbl-om-cell { display: flex; align-items: center; gap: 10px; }
.tbl-om-logo { width: 28px; height: 28px; object-fit: contain; flex: none; }
.tbl-om-sub { font-size: 0.75rem; color: var(--ink-muted); font-weight: 400; display: block; }
.tbl-pct-cell { display: flex; align-items: center; gap: 8px; justify-content: flex-end; }
.mini-track { width: 64px; height: 6px; background: var(--track); border-radius: 3px; overflow: hidden; }
.mini-fill { height: 100%; border-radius: 3px; }
.tbl-action-btn {
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 5px 10px;
  font-size: 0.75rem;
  font-weight: 700;
  color: var(--primary);
  cursor: pointer;
}
.tbl-action-btn:hover { background: var(--primary); color: #FFFFFF; }

/* SVG Classes */
.svg { width: 100%; height: auto; display: block; }
.s-lbl { font-size: 11px; font-weight: 700; letter-spacing: 0.06em; fill: var(--ink-muted); }
.s-seg { font-size: 11px; fill: var(--warning-ink); font-weight: 600; font-family: var(--mono); }
.s-seg-ok { fill: var(--success-strong); }
.s-brk { stroke: var(--border-strong); stroke-width: 1; }
.s-cat { font-size: 12px; fill: var(--ink-muted); }
.s-cat-ok { fill: var(--success-strong); font-weight: 600; }
.s-val { font-size: 13px; font-weight: 700; font-family: var(--mono); }
.s-on { fill: #FFFFFF; }
.s-ok { fill: var(--success-strong); }
.s-warn { fill: var(--warning-ink); }
.s-on2 { fill: var(--ink); }
.s-num { font-size: 11px; font-weight: 600; font-family: var(--mono); }
.s-neg { fill: var(--danger); }
.s-conn { stroke: var(--border-strong); stroke-width: 1; stroke-dasharray: 4 3; }
.s-zero { stroke: var(--border-strong); stroke-width: 1.5; }
.s-grid { stroke: var(--border); stroke-width: 1; }
.s-ax { font-size: 11px; fill: var(--ink-muted); font-family: var(--mono); }
.s-line { fill: none; stroke: var(--success); stroke-width: 2.5; }
.s-area { fill: var(--success); opacity: .12; }
.s-dot { fill: var(--success); }
.s-conv { font-size: 10.5px; fill: var(--ink-muted); }

/* Modal Drill-Down (Física Spring + Backdrop Blur) */
.modal {
  position: fixed;
  inset: 0;
  z-index: 999;
  display: none;
  align-items: center;
  justify-content: center;
  padding: 24px 16px;
  background: rgba(2, 6, 23, 0.65);
  backdrop-filter: blur(10px);
  -webkit-backdrop-filter: blur(10px);
  overflow-y: auto;
}
.modal.open { display: flex; }
.modal-panel {
  position: relative;
  background: var(--bg-surface);
  border: 1px solid var(--border-strong);
  border-radius: 18px;
  box-shadow: var(--shadow-lg);
  max-width: 780px;
  width: 100%;
  padding: 28px 30px;
  animation: modalIn .25s var(--ease-spring);
  /* [FIX janela fora da tela] o painel NUNCA excede a viewport: limita a altura e
     rola o conteúdo por dentro (o ✕ fica sempre visível). Antes crescia sem limite
     e, com align-items:center, o topo era cortado e ficava inalcançável. */
  max-height: calc(100vh - 48px);
  max-height: calc(100dvh - 48px);
  margin: auto;
  display: flex;
  flex-direction: column;
  min-height: 0;
}
#modal-body {
  overflow-y: auto;
  overscroll-behavior: contain;
  min-height: 0;
  margin-right: -10px;
  padding-right: 10px;
}
@keyframes modalIn {
  from { opacity: 0; transform: scale(0.96) translateY(10px); }
  to { opacity: 1; transform: scale(1) translateY(0); }
}
.modal-x {
  position: absolute;
  top: 16px;
  right: 16px;
  width: 36px;
  height: 36px;
  border: 1px solid var(--border);
  background: var(--bg-subtle);
  border-radius: 10px;
  color: var(--ink-muted);
  cursor: pointer;
  font-size: 16px;
  line-height: 1;
  display: grid;
  place-items: center;
  transition: transform .2s var(--ease-out-expo), color .2s var(--ease-out-expo), background-color .2s var(--ease-out-expo), border-color .2s var(--ease-out-expo);
}
.modal-x:hover { color: var(--ink); border-color: var(--border-strong); transform: scale(1.05); }
#modal-body h3 { font-size: 1.25rem; font-weight: 700; margin-bottom: 4px; padding-right: 44px; }
.m-sub { font-size: 0.8125rem; color: var(--ink-muted); margin-bottom: 16px; }
.m-ficha {
  display: flex;
  flex-wrap: wrap;
  gap: 10px 20px;
  margin: 4px 0 16px;
  padding: 14px 16px;
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 12px;
}
.m-ficha span { display: flex; flex-direction: column; gap: 2px; font-size: 0.6875rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--ink-muted); min-width: 120px; }
.m-ficha span.wide { flex-basis: 100%; }
.m-ficha b { font-size: 0.875rem; font-weight: 600; color: var(--ink); text-transform: none; }
.m-kpis { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; }
.m-kpis span {
  flex: 1;
  min-width: 120px;
  display: flex;
  flex-direction: column;
  gap: 2px;
  font-size: 0.6875rem;
  color: var(--ink-muted);
  text-transform: uppercase;
  letter-spacing: 0.04em;
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 10px 12px;
}
.m-kpis b { font-size: 1.125rem; font-weight: 700; color: var(--ink); font-family: var(--mono); }
.m-kpis .ok b { color: var(--success-strong); }
.m-kpis .ok { border-left: 3px solid var(--success); }
.m-formula { font-size: 0.75rem; color: var(--ink-muted); margin: -8px 0 16px; font-style: italic; }
.m-ncs-h { font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: var(--ink-muted); margin-bottom: 10px; }
.m-ncs { display: flex; flex-direction: column; gap: 10px; max-height: 46vh; overflow-y: auto; }
.m-nc { border: 1px solid var(--border); border-radius: 12px; padding: 12px 14px; background: var(--bg-surface); }
.m-nc-h { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; }
.m-nc-num { font-weight: 700; font-size: 0.875rem; font-family: var(--mono); }
.m-nc-val { font-weight: 700; font-size: 0.875rem; color: var(--success-strong); white-space: nowrap; font-family: var(--mono); }
.m-nc-val.neg { color: var(--danger); }
.m-nc-op { font-size: 0.6875rem; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-muted); margin-top: 4px; }
.m-nc-desc { font-size: 0.8125rem; color: var(--ink); margin-top: 6px; line-height: 1.55; white-space: pre-wrap; word-break: break-word; }

/* Toast Feedback */
.toast {
  position: fixed;
  bottom: 28px;
  right: 28px;
  background: var(--neutral-900);
  color: #FFFFFF;
  padding: 14px 22px;
  border-radius: 12px;
  box-shadow: var(--shadow-lg);
  font-size: 0.875rem;
  font-weight: 600;
  display: flex;
  align-items: center;
  gap: 12px;
  z-index: 9999;
  opacity: 0;
  transform: translateY(16px);
  transition: opacity .25s var(--ease-out-expo), transform .25s var(--ease-spring);
  pointer-events: none;
  border: 1px solid rgba(255, 255, 255, 0.15);
}
:root[data-theme="dark"] .toast { background: var(--bg-elevated); color: var(--ink); border-color: var(--border-strong); }
.toast.show { opacity: 1; transform: translateY(0); }
.toast-ic { font-size: 1.125rem; }

/* Banner de Alerta */
.banner {
  background: var(--danger-bg);
  border: 1px solid var(--danger-border);
  border-radius: 12px;
  padding: 14px 18px;
  margin-top: 20px;
  font-size: 0.8125rem;
  color: var(--danger);
}
.banner ul { margin: 6px 0 0 20px; }

/* Rodapé */
.rodape {
  max-width: 1280px;
  margin: 48px auto 0;
  padding: 24px;
  border-top: 1px solid var(--border);
  color: var(--ink-muted);
  font-size: 0.75rem;
  line-height: 1.7;
}
.rodape b { color: var(--ink); }
.rodape-brand { color: var(--gold-dark); font-weight: 700; font-size: 0.8125rem; letter-spacing: 0.02em; margin-bottom: 8px; }

/* Acessibilidade & Estados de Foco */
:focus-visible { outline: 2px solid var(--border-focus); outline-offset: 2px; border-radius: 6px; }
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after { transition: none !important; animation: none !important; }
}

/* Responsividade Mobile-First */
@media (max-width: 1023px) {
  .grid2 { grid-template-columns: 1fr; }
  .hero { grid-template-columns: 1fr; }
  .podium-wrap { grid-template-columns: 1fr; gap: 14px; }
  .podium-gold { order: 1; margin-bottom: 0; padding-top: 24px; }
  .podium-silver { order: 2; }
  .podium-bronze { order: 3; }
  .kpis { grid-template-columns: repeat(2, 1fr); }
}

@media (max-width: 640px) {
  .wrap { padding: 0 16px 48px; }
  .topbar { padding: 12px 16px; flex-wrap: wrap; }
  h1 { font-size: 1.25rem; }
  .subtitle { display: none; }
  .hero-num { font-size: 2.2rem; }
  .kpis { grid-template-columns: 1fr; }
  .det td.mono2, .det th:first-child, .det td:first-child { position: sticky; left: 0; background: var(--bg-surface); }
}

/* Indicador e contenedor aprimorado para tabelas mobile */
.tbl-scroll-hint {
  display: none;
}
@media (max-width: 768px) {
  .tbl-scroll-hint {
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 6px;
    font-size: 0.6875rem;
    font-weight: 600;
    color: var(--ink-muted);
    padding: 6px 12px;
    background: var(--bg-subtle);
    border-bottom: 1px solid var(--border);
  }
}
/* ==========================================================================
   ESTILOS PARA A ABA: HISTÓRICO DE NOTAS DE CRÉDITO E MODAL EXPANDIDO
   ========================================================================== */
.omds-hist {
  background: var(--bg-subtle);
  border-color: var(--primary-600);
}
.omds-hist.on {
  background: linear-gradient(135deg, #1C4A73 0%, #2563EB 100%);
  color: #fff;
  border-color: #2563EB;
}
.omds-hist.on span { color: #fff; }
.hist-icon { font-size: 1.25rem; }

.unidade-hist { display: none; }
.hist-header-card {
  margin-top: 24px;
  padding: 26px 30px;
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-top: 5px solid var(--primary-600);
  border-radius: 16px;
  box-shadow: var(--shadow-md);
}

.hist-kpis {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
  gap: 16px;
  margin-top: 22px;
}
.hist-kpi {
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 18px 20px;
  box-shadow: var(--shadow-sm);
  transition: transform .2s ease, box-shadow .2s ease;
  display: flex;
  flex-direction: column;
  gap: 6px;
  border-left: 4px solid var(--primary-600);
}
.hist-kpi:hover {
  transform: translateY(-2px);
  box-shadow: var(--shadow-md);
}
.hist-kpi.kpi-total { border-left-color: #6366F1; }
.hist-kpi.kpi-prov  { border-left-color: #2563EB; }
.hist-kpi.kpi-emp   { border-left-color: #D97706; }
.hist-kpi.kpi-saldo { border-left-color: #059669; }
.hist-kpi-lbl { font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; color: var(--ink-muted); }
.hist-kpi-val { font-size: 1.625rem; font-weight: 800; color: var(--ink); font-family: var(--serif); line-height: 1.2; }
.hist-kpi-sub { font-size: 0.8125rem; color: var(--ink-muted); display: flex; align-items: center; gap: 6px; }
.kpi-dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
.dot-indigo { background: #6366F1; }
.dot-blue   { background: #2563EB; }
.dot-amber  { background: #D97706; }
.dot-green  { background: #059669; }

.hist-filters-card {
  margin-top: 14px;
  margin-bottom: 20px;
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 20px;
  box-shadow: var(--shadow-sm);
}
.hist-filters-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 14px;
  align-items: end;
}
.hist-filter-group { display: flex; flex-direction: column; gap: 6px; }
.hist-filter-group label { font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; color: var(--ink-muted); }
.hist-select, .hist-input {
  padding: 9px 12px;
  border-radius: 8px;
  border: 1px solid var(--border-strong);
  background: var(--bg);
  color: var(--ink);
  font-family: var(--sans);
  font-size: 0.875rem;
  width: 100%;
  transition: border-color .15s ease, box-shadow .15s ease;
}
.hist-select:focus, .hist-input:focus {
  outline: none;
  border-color: var(--border-focus);
  box-shadow: 0 0 0 3px rgba(37,99,235,0.15);
}
.hist-toolbar-actions {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  align-items: center;
  margin-top: 16px;
  padding-top: 16px;
  border-top: 1px solid var(--border);
  justify-content: space-between;
}
.hist-actions-l { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
.hist-actions-r { display: flex; align-items: center; gap: 10px; }
.btn-clear-flt {
  background: var(--bg-subtle);
  border: 1px solid var(--border-strong);
  color: var(--ink-muted);
  border-radius: 8px;
  padding: 7px 14px;
  font-size: 0.8125rem;
  font-weight: 600;
  cursor: pointer;
  transition: all .15s ease;
}
.btn-clear-flt:hover {
  background: var(--bg-surface);
  color: var(--danger);
  border-color: var(--danger-border);
}
.hist-count-badge { font-size: 0.8125rem; font-weight: 600; color: var(--ink-muted); }

/* Tabela e Badges */
.det-hist .nc-tag { color: var(--primary-600); font-family: var(--mono); letter-spacing: 0.02em; }
.ug-pill {
  display: inline-block;
  padding: 1px 6px;
  border-radius: 4px;
  font-family: var(--mono);
  font-size: 0.75rem;
  font-weight: 600;
  margin-right: 4px;
}
.ug-pill.emit { background: var(--bg-subtle); border: 1px solid var(--border); color: var(--ink); }
.ug-pill.fav  { background: var(--primary-50); border: 1px solid rgba(37,99,235,0.2); color: var(--primary-600); }

.pill-status {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 3px 10px;
  border-radius: 999px;
  font-size: 0.75rem;
  font-weight: 700;
  white-space: nowrap;
}
.status-disp      { background: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; }
.status-parcial   { background: #FFFBEB; color: #92400E; border: 1px solid #FDE68A; }
.status-exec      { background: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; }
.status-detalhada { background: #EFF6FF; color: #1E40AF; border: 1px solid #BFDBFE; }
.status-canc      { background: #FEF2F2; color: #991B1B; border: 1px solid #FECACA; }
.status-zerada    { background: #F8FAFC; color: #64748B; border: 1px solid #E2E8F0; }

:root[data-theme="dark"] .status-disp      { background: #064E3B; color: #A7F3D0; border-color: #047857; }
:root[data-theme="dark"] .status-parcial   { background: #78350F; color: #FDE68A; border-color: #B45309; }
:root[data-theme="dark"] .status-exec      { background: #1E293B; color: #94A3B8; border-color: #334155; }
:root[data-theme="dark"] .status-detalhada { background: #1E3A8A; color: #BFDBFE; border-color: #2563EB; }
:root[data-theme="dark"] .status-canc      { background: #7F1D1D; color: #FECACA; border-color: #B91C1C; }
:root[data-theme="dark"] .status-zerada    { background: #0F172A; color: #64748B; border-color: #1E293B; }

/* Paginação */
.hist-pagination {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 14px 18px;
  background: var(--bg-surface);
  border: 1px solid var(--border);
  border-top: none;
  border-radius: 0 0 14px 14px;
  flex-wrap: wrap;
  gap: 10px;
}
.pag-info { font-size: 0.8125rem; color: var(--ink-muted); }
.pag-btns { display: flex; gap: 8px; }
.btn-pag {
  background: var(--bg-surface);
  border: 1px solid var(--border-strong);
  border-radius: 6px;
  padding: 6px 14px;
  font-size: 0.8125rem;
  font-weight: 600;
  cursor: pointer;
  color: var(--ink);
  transition: all .15s ease;
}
.btn-pag:hover:not(:disabled) {
  background: var(--primary-50);
  border-color: var(--primary-600);
  color: var(--primary-600);
}
.btn-pag:disabled {
  opacity: 0.45;
  cursor: not-allowed;
}

/* Modal Detalhe da NC */
/* ==========================================================================
   EXECUTIVE MILITARY MODAL DESIGN (High-Fidelity FinTech & Defense Analytics)
   ========================================================================== */
.modal {
  position: fixed;
  inset: 0;
  z-index: 9999;
  background: rgba(3, 7, 18, 0.78);
  backdrop-filter: blur(14px);
  -webkit-backdrop-filter: blur(14px);
  display: none;
  align-items: center;
  justify-content: center;
  padding: 20px;
  overflow-y: auto;
}
.modal.open { display: flex; }
.modal-panel {
  position: relative;
  background: var(--bg-surface);
  border: 1px solid var(--border-strong);
  border-radius: 20px;
  box-shadow: 0 30px 70px -15px rgba(0, 0, 0, 0.7), 0 0 0 1px rgba(255, 255, 255, 0.08);
  max-width: 860px;
  width: 100%;
  max-height: calc(100vh - 40px);
  max-height: calc(100dvh - 40px);
  margin: auto;
  display: flex;
  flex-direction: column;
  min-height: 0;
  overflow: hidden;
  animation: modalPopIn .24s cubic-bezier(0.16, 1, 0.3, 1);
}
@keyframes modalPopIn {
  from { opacity: 0; transform: scale(0.96) translateY(12px); }
  to   { opacity: 1; transform: scale(1) translateY(0); }
}
.modal-x {
  position: absolute;
  top: 16px;
  right: 18px;
  width: 36px;
  height: 36px;
  border: 1px solid var(--border);
  background: var(--bg-subtle);
  border-radius: 10px;
  color: var(--ink-muted);
  cursor: pointer;
  font-size: 15px;
  line-height: 1;
  display: grid;
  place-items: center;
  transition: all .2s ease;
  z-index: 20;
}
.modal-x:hover {
  background: var(--danger-bg);
  color: var(--danger);
  border-color: var(--danger-border);
  transform: scale(1.08);
}
#modal-body {
  overflow-y: auto;
  overscroll-behavior: contain;
  min-height: 0;
  padding: 0;
}
.m-accent-bar {
  height: 5px;
  width: 100%;
  background: linear-gradient(90deg, #10B981 0%, #059669 35%, #2563EB 75%, #8B5CF6 100%);
}
.m-content-wrap {
  padding: 24px 28px 28px 28px;
}
.m-header-v2 {
  margin-bottom: 20px;
  padding-bottom: 18px;
  border-bottom: 1px solid var(--border);
}
.m-header-meta-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 10px;
  flex-wrap: wrap;
}
.m-badge-op {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 0.6875rem;
  font-weight: 800;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  background: rgba(16, 185, 129, 0.12);
  color: #34D399;
  border: 1px solid rgba(16, 185, 129, 0.3);
}
.m-badge-status-lg {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 5px 12px;
  border-radius: 999px;
  font-size: 0.75rem;
  font-weight: 700;
}
.m-badge-status-lg.status-ok {
  background: rgba(16, 185, 129, 0.15);
  color: #34D399;
  border: 1px solid rgba(16, 185, 129, 0.35);
  box-shadow: 0 0 12px rgba(16, 185, 129, 0.2);
}
.m-badge-status-lg.status-warn {
  background: rgba(245, 158, 11, 0.15);
  color: #FBBF24;
  border: 1px solid rgba(245, 158, 11, 0.35);
}
.m-badge-status-lg.status-info {
  background: rgba(59, 130, 246, 0.15);
  color: #93C5FD;
  border: 1px solid rgba(59, 130, 246, 0.35);
}
.m-title-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  flex-wrap: wrap;
}
.m-nc-code-block h3 {
  font-family: 'JetBrains Mono', monospace;
  font-size: 1.35rem;
  font-weight: 800;
  letter-spacing: -0.02em;
  color: var(--ink);
  margin: 0;
  word-break: break-all;
}
.m-sub-meta {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-top: 8px;
  flex-wrap: wrap;
}
.m-meta-chip {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 3px 9px;
  border-radius: 6px;
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  font-size: 0.75rem;
  color: var(--ink-muted);
  font-weight: 600;
}
.m-fin-section {
  margin-bottom: 22px;
}
.m-fin-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(175px, 1fr));
  gap: 12px;
  margin-bottom: 12px;
}
.m-fin-card {
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 14px 16px;
  display: flex;
  flex-direction: column;
  gap: 4px;
  position: relative;
  transition: transform .15s ease, border-color .15s ease;
}
.m-fin-card:hover {
  transform: translateY(-2px);
  border-color: var(--border-strong);
}
.m-fin-card.hero-saldo {
  background: linear-gradient(135deg, rgba(16, 185, 129, 0.14) 0%, rgba(5, 150, 105, 0.05) 100%);
  border: 1.5px solid rgba(16, 185, 129, 0.45);
  box-shadow: 0 4px 20px rgba(16, 185, 129, 0.15);
}
.m-fin-label {
  font-size: 0.6875rem;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  color: var(--ink-muted);
}
.m-fin-val {
  font-family: 'Inter', sans-serif;
  font-size: 1.15rem;
  font-weight: 700;
  color: var(--ink);
  font-variant-numeric: tabular-nums;
}
.m-fin-val.col-prov { color: var(--primary-600); }
.m-fin-val.col-emp  { color: var(--warning-ink); }
.m-fin-val-hero {
  font-family: 'Newsreader', Georgia, serif;
  font-size: 1.45rem;
  font-weight: 800;
  color: #34D399;
  font-variant-numeric: tabular-nums;
}
.m-fin-sub {
  font-size: 0.725rem;
  color: var(--ink-muted);
}
.m-fin-tag-hero {
  display: inline-flex;
  align-items: center;
  font-size: 0.725rem;
  font-weight: 700;
  color: #34D399;
}
.m-pipeline-card {
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 14px 18px;
}
.m-pipeline-header {
  font-size: 0.75rem;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  color: var(--ink-muted);
  margin-bottom: 12px;
}
.m-stages-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 16px;
}
.m-stage-item {
  display: flex;
  flex-direction: column;
  gap: 5px;
}
.m-stage-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 0.75rem;
  color: var(--ink-muted);
}
.m-stage-header b {
  color: var(--ink);
  font-weight: 700;
  font-variant-numeric: tabular-nums;
}
.m-progress-track {
  height: 7px;
  background: var(--track);
  border-radius: 999px;
  overflow: hidden;
}
.m-progress-bar {
  height: 100%;
  border-radius: 999px;
  transition: width .3s ease;
}
.m-progress-bar.bar-emp { background: linear-gradient(90deg, #F59E0B, #FBBF24); }
.m-progress-bar.bar-liq { background: linear-gradient(90deg, #3B82F6, #60A5FA); }
.m-progress-bar.bar-pag { background: linear-gradient(90deg, #10B981, #34D399); }
.m-stage-sub {
  font-size: 0.7rem;
  color: var(--ink-soft);
  font-variant-numeric: tabular-nums;
}
.m-route-grid {
  display: grid;
  grid-template-columns: 1fr 40px 1fr;
  align-items: center;
  gap: 12px;
  margin-bottom: 20px;
}
@media (max-width: 680px) {
  .m-route-grid { grid-template-columns: 1fr; }
  .m-route-arrow { transform: rotate(90deg); margin: 6px auto; }
}
.m-route-card {
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 14px 16px;
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.m-route-badge {
  font-size: 0.6875rem;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  color: var(--ink-muted);
  margin-bottom: 2px;
}
.m-route-title {
  font-size: 0.875rem;
  font-weight: 700;
  color: var(--ink);
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}
.m-route-desc {
  font-size: 0.75rem;
  color: var(--ink-muted);
}
.m-route-arrow {
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--primary-600);
  font-size: 1.25rem;
  font-weight: 700;
}
.m-class-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 12px;
  margin-bottom: 20px;
}
.m-class-card {
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 12px 14px;
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.m-class-label {
  font-size: 0.6875rem;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  color: var(--ink-muted);
}
.m-class-code {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.95rem;
  font-weight: 700;
  color: var(--ink);
}
.m-class-desc {
  font-size: 0.75rem;
  color: var(--ink-muted);
  line-height: 1.35;
}
.m-justif-card {
  background: var(--bg-subtle);
  border: 1px solid var(--border);
  border-left: 4px solid var(--success-main);
  border-radius: 12px;
  padding: 16px 18px;
  margin-bottom: 22px;
}
.m-justif-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 8px;
}
.m-justif-title {
  font-size: 0.75rem;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  color: var(--ink-muted);
}
.m-justif-body {
  font-size: 0.8125rem;
  line-height: 1.6;
  color: var(--ink);
  white-space: pre-wrap;
  word-break: break-word;
  max-height: 200px;
  overflow-y: auto;
  font-family: 'Inter', sans-serif;
}
.m-footer-actions-v2 {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding-top: 16px;
  border-top: 1px solid var(--border);
  flex-wrap: wrap;
}
.m-btn-pill {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  padding: 9px 18px;
  border-radius: 10px;
  font-size: 0.8125rem;
  font-weight: 600;
  cursor: pointer;
  transition: all .2s ease;
  border: 1px solid var(--border-strong);
  background: var(--bg-surface);
  color: var(--ink);
}
.m-btn-pill:hover {
  background: var(--bg-subtle);
  border-color: var(--primary-600);
  transform: translateY(-1px);
}
.m-btn-pill.primary {
  background: var(--primary-600);
  color: #FFFFFF;
  border-color: var(--primary-600);
  box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25);
}
.m-btn-pill.primary:hover {
  background: var(--primary-strong);
  box-shadow: 0 6px 16px rgba(37, 99, 235, 0.35);
}

"""

JS = r"""
(function(){
  var s=localStorage.getItem('bcms-theme');
  if(s){ document.documentElement.setAttribute('data-theme',s); }
})();

function bcmsTheme(){
  var h=document.documentElement;
  var cur=h.getAttribute('data-theme');
  var dark=cur?cur==='dark':window.matchMedia('(prefers-color-scheme:dark)').matches;
  var next=dark?'light':'dark';
  h.setAttribute('data-theme',next);
  localStorage.setItem('bcms-theme',next);
  var btn=document.getElementById('themeBtn');
  if(btn) btn.setAttribute('aria-pressed',next==='dark');
}

function bcmsTab(btn,id){
  var list=btn.parentNode.querySelectorAll('.tab');
  list.forEach(function(b){b.classList.remove('on');b.setAttribute('aria-selected','false');b.tabIndex=-1;});
  btn.classList.add('on');btn.setAttribute('aria-selected','true');btn.tabIndex=0;
  (btn.closest('.unidade')||document).querySelectorAll('.tabpanel').forEach(function(p){p.style.display='none';});
  var target=document.getElementById(id);
  if(target) target.style.display='block';
}

function bcmsTabKey(e,btn){
  var t=Array.prototype.slice.call(btn.parentNode.querySelectorAll('.tab'));
  var i=t.indexOf(btn);
  if(e.key==='ArrowRight'){e.preventDefault();t[(i+1)%t.length].focus();t[(i+1)%t.length].click();}
  else if(e.key==='ArrowLeft'){e.preventDefault();t[(i-1+t.length)%t.length].focus();t[(i-1+t.length)%t.length].click();}
}

function bcmsView(btn,which){
  var m=btn.closest('.unidade')||document;
  btn.parentNode.querySelectorAll('.toptab').forEach(function(b){b.classList.remove('on');b.setAttribute('aria-selected','false');});
  btn.classList.add('on');btn.setAttribute('aria-selected','true');
  var vr=m.querySelector('.view-resumo');if(vr)vr.style.display=which==='resumo'?'':'none';
  var vc=m.querySelector('.view-completo');if(vc)vc.style.display=which==='completo'?'':'none';
  var vh=m.querySelector('.view-historico');if(vh)vh.style.display=which==='historico'?'':'none';
  if(which==='historico'){
    var sfx=m.getAttribute('data-key');
    if(sfx && typeof bcmsInitUHist === 'function'){
      bcmsInitUHist(sfx);
    }
  }
  window.scrollTo({top:0,behavior:'smooth'});
}

function trocaOMDS(btn){
  var key=btn.getAttribute('data-key');
  document.querySelectorAll('.omds-nav .omds').forEach(function(b){b.classList.remove('on');b.setAttribute('aria-current','false');});
  btn.classList.add('on');btn.setAttribute('aria-current','true');
  document.querySelectorAll('.unidade').forEach(function(s){s.style.display=(s.getAttribute('data-key')===key)?'':'none';});
  var esc=document.getElementById('uEscalao');
  if(key==='CATRIMANI'){
    document.documentElement.style.setProperty('--accent','#15803D');
    var em=document.getElementById('emblema');if(em){em.src='assets/logos/CATRIMANI.png';em.alt='Brasão Operação Catrimani';}
    if(esc)esc.textContent='OPERAÇÃO CONJUNTA · TERRITÓRIO INDÍGENA YANOMAMI / RORAIMA';
    var t=document.getElementById('uTitulo');if(t)t.textContent='Operação Catrimani II — Execução Orçamentária Multi-UGs';
    var n=document.getElementById('uNome');if(n)n.textContent='Ação Governamental 21EM · Comando Militar da Amazônia (CMA)';
    var uu=document.getElementById('uUasg');if(uu)uu.textContent='Acompanhamento Orçamentário de 10 Unidades Gestoras Executoras';
    try{document.title='Operação Catrimani II — Execução Orçamentária';}catch(e){}
    if(!CATR_INITIALIZED){
      bcmsInitCatrimani();
    }
  } else if(key==='RANKING'){
    document.documentElement.style.setProperty('--accent','#D97706');
    var em=document.getElementById('emblema');if(em){em.src='assets/logos/12RM.png';em.alt='Brasão 12ª RM';}
    if(esc)esc.textContent='12ª REGIÃO MILITAR · COMANDO MILITAR DA AMAZÔNIA';
    var t=document.getElementById('uTitulo');if(t)t.textContent='Ranking & Benchmarking OMDS';
    var n=document.getElementById('uNome');if(n)n.textContent='Visão Comparativa de Execução Orçamentária das 9 Organizações Militares';
    var uu=document.getElementById('uUasg');if(uu)uu.textContent='18 UASGs (OGU + FEx) · Exercício 2026';
    try{document.title='Ranking & Comparativo OMDS';}catch(e){}
  } else if(key==='HISTORICO'){
    document.documentElement.style.setProperty('--accent','#15803D');
    var em=document.getElementById('emblema');if(em){em.src='assets/logos/6BEC.png';em.alt='Brasão 6º BEC';}
    if(esc)esc.textContent='6º BATALHÃO DE ENGENHARIA DE CONSTRUÇÃO · 12ª RM';
    var t=document.getElementById('uTitulo');if(t)t.textContent='Histórico de Notas de Crédito — 6º BEC';
    var n=document.getElementById('uNome');if(n)n.textContent='6º Batalhão de Engenharia de Construção — Registro Geral de NCs';
    var uu=document.getElementById('uUasg');if(uu)uu.textContent='UASGs 160353 (OGU) e 167353 (FEx) · Exercício 2026';
    try{document.title='Histórico de NCs — 6º BEC';}catch(e){}
    if(!HIST_INITIALIZED){
      bcmsInitHistorico();
    }
  } else {
    var u=UNIDADES[key];if(!u)return;
    document.documentElement.style.setProperty('--accent',u.accent);
    var em=document.getElementById('emblema');if(em){em.src='assets/logos/'+u.logo;em.alt='Brasão '+u.sigla;}
    if(esc)esc.textContent='12ª REGIÃO MILITAR · COMANDO MILITAR DA AMAZÔNIA';
    var t=document.getElementById('uTitulo');if(t)t.textContent='Crédito Disponível — '+u.sigla;
    var n=document.getElementById('uNome');if(n)n.textContent=u.nome;
    var uu=document.getElementById('uUasg');if(uu)uu.textContent='UASGs '+u.ogu+' (OGU) e '+u.fex+' (FEx)';
    try{document.title='Crédito Disponível — '+u.sigla;}catch(e){}
  }
  try{localStorage.setItem('bcms-omds',key);}catch(e){}
  window.scrollTo({top:0,behavior:'smooth'});
}

function trocaOMDSPorKey(key){
  var btn=document.querySelector('.omds-nav .omds[data-key="'+key+'"]');
  if(btn){btn.click();window.scrollTo({top:0,behavior:'smooth'});}
}

(function(){
  try{
    var k=localStorage.getItem('bcms-omds');
    if(k){
      var b=document.querySelector('.omds-nav .omds[data-key="'+k+'"]');
      if(b&&!b.classList.contains('on')) b.click();
    }
  }catch(e){}
  setTimeout(function(){
    if(typeof HISTDATA !== 'undefined' && HISTDATA && HISTDATA.items && !HIST_INITIALIZED){
      bcmsInitHistorico();
    }
  }, 100);
})();

function bcmsSearch(inp,tid){
  var q=inp.value.toLowerCase();
  var container=document.getElementById(tid);
  if(!container)return;
  var tb=container.querySelector('tbody');
  if(!tb)return;
  var rows=tb.querySelectorAll('tr');
  var n=0;
  rows.forEach(function(r){
    var ok=r.textContent.toLowerCase().indexOf(q)>-1;
    r.style.display=ok?'':'none';
    if(ok) n++;
  });
  var cnt=document.getElementById('cnt-'+tid);
  if(cnt){
    var u=cnt.getAttribute('data-unit')||'linha(s)';
    cnt.textContent=n+' '+u;
  }
}

function bcmsNC(row){
  var ncKey=row.getAttribute('data-nc');
  var d=NCDATA[ncKey];
  if(!d)return;
  if(typeof HISTDATA !== 'undefined' && HISTDATA && HISTDATA.items){
    for(var i=0; i<HISTDATA.items.length; i++){
      if(HISTDATA.items[i].nc === d.nc || HISTDATA.items[i].hid === ncKey){
        bcmsDetalheNC(HISTDATA.items[i].hid);
        return;
      }
    }
  }
  var neg=d.val<0;
  var h='<h3 id="modal-title">NC '+bcmsEsc(d.nc)+'</h3>';
  h+='<p class="m-sub">'+bcmsEsc((d.u?d.u+' · ':'')+'Ação '+d.acao+' · PI '+d.pi+' · ND '+d.nd+(d.ndn?' — '+d.ndn:''))+'</p>';
  h+='<div class="m-kpis"><span>Data<b>'+bcmsEsc(d.dia||'—')+'</b></span><span>Operação<b class="op">'+bcmsEsc(d.op||'—')+'</b></span><span class="'+(neg?'':'ok')+'">Valor<b class="'+(neg?'neg':'')+'">'+bcmsBRL(d.val)+'</b></span></div>';
  h+='<div class="m-ncs-h">Descrição completa do lançamento</div>';
  h+='<div class="m-nc"><div class="m-nc-desc">'+bcmsEsc(d.obj||'(sem descrição)')+'</div></div>';
  document.getElementById('modal-body').innerHTML=h;
  var m=document.getElementById('modal');
  m.classList.add('open');
  m.setAttribute('aria-hidden','false');
  var x=document.querySelector('.modal-x');
  if(x) x.focus();
}

function bcmsSort(th){
  var table=th.closest('table');
  var idx=Array.prototype.indexOf.call(th.parentNode.children,th);
  var dir=th.getAttribute('aria-sort')==='ascending'?'descending':'ascending';
  th.parentNode.querySelectorAll('th').forEach(function(h){h.setAttribute('aria-sort','none');h.querySelector('.sort').textContent='';});
  th.setAttribute('aria-sort',dir);
  th.querySelector('.sort').textContent=dir==='ascending'?' ▲':' ▼';
  var tb=table.querySelector('tbody');
  var rows=Array.prototype.slice.call(tb.querySelectorAll('tr'));
  rows.sort(function(a,b){
    var ca=a.children[idx], cb=b.children[idx];
    var da=ca.getAttribute('data-sort'), db=cb.getAttribute('data-sort');
    var va,vb;
    if(da!==null&&db!==null){va=parseFloat(da);vb=parseFloat(db);}
    else{va=ca.textContent.trim().toLowerCase();vb=cb.textContent.trim().toLowerCase();}
    if(va<vb)return dir==='ascending'?-1:1;
    if(va>vb)return dir==='ascending'?1:-1;
    return 0;
  });
  rows.forEach(function(r){tb.appendChild(r);});
}

function bcmsEsc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function bcmsEscXml(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c];});}
function bcmsBRL(v){var neg=v<0,s=Math.abs(v).toFixed(2).split('.');var i=s[0].replace(/\B(?=(\d{3})+(?!\d))/g,'.');return (neg?'−R$ ':'R$ ')+i+','+s[1];}

function bcmsToast(msg){
  var old=document.getElementById('bcms-toast');if(old)old.remove();
  var t=document.createElement('div');t.id='bcms-toast';t.className='toast';
  t.innerHTML='<span class="toast-ic">✨</span><span>'+bcmsEsc(msg)+'</span>';
  document.body.appendChild(t);
  setTimeout(function(){t.classList.add('show');},10);
  setTimeout(function(){t.classList.remove('show');setTimeout(function(){if(t.parentNode)t.remove();},300);},3500);
}

/* ---- Filtros combinados da lista "em tela" (Fonte + Ação + ND + Idade + busca).
   O botão Exportar Excel usa exatamente o que sobra visível aqui. ---- */
function bcmsFiltra(sfx){
  var tid='tab-emtela-'+sfx;
  var cont=document.getElementById(tid); if(!cont) return;
  var tb=cont.querySelector('tbody'); if(!tb) return;
  var v=function(id){var e=document.getElementById(id+'-'+sfx); return e?e.value:'';};
  var q=(v('q-tab-emtela')||'').toLowerCase();
  var fFonte=v('f-fonte'), fAcao=v('f-acao'), fNd=v('f-nd'), fId=v('f-idade');
  var n=0, soma=0;
  tb.querySelectorAll('tr').forEach(function(r){
    var ok=true;
    if(fFonte && r.getAttribute('data-fonte')!==fFonte) ok=false;
    if(ok&&fAcao && r.getAttribute('data-acao')!==fAcao) ok=false;
    if(ok&&fNd && r.getAttribute('data-nd')!==fNd) ok=false;
    if(ok&&fId && r.getAttribute('data-faixa')!==fId) ok=false;
    if(ok&&q && r.textContent.toLowerCase().indexOf(q)===-1) ok=false;
    r.style.display=ok?'':'none';
    if(ok){n++; soma+=parseFloat(r.getAttribute('data-val')||'0')||0;}
  });
  var cnt=document.getElementById('cnt-'+tid);
  if(cnt){var u=cnt.getAttribute('data-unit')||'linha(s)'; cnt.textContent=n+' '+u;}
  /* rodapé passa a refletir o subconjunto filtrado */
  var tf=cont.querySelector('tfoot tr');
  if(tf){
    var tds=tf.querySelectorAll('td');
    var ativo=!!(fFonte||fAcao||fNd||fId||q);
    if(tds.length){tds[0].textContent=(ativo?'FILTRADO · ':'TOTAL · ')+n+' Nota(s) de Crédito em tela';}
    if(tds.length>1){tds[tds.length-2].textContent=bcmsBRL(soma);}
  }
  var res=document.getElementById('flt-res-'+sfx);
  if(res){
    var pk=[];
    if(fFonte)pk.push('Fonte '+fFonte);
    if(fAcao)pk.push('Ação '+fAcao);
    if(fNd)pk.push('ND '+fNd);
    if(fId)pk.push({v:'≤30d',a:'31–60d',r:'>60d'}[fId]||fId);
    if(q)pk.push('“'+q+'”');
    res.textContent=pk.length?(pk.join(' · ')+' — '+bcmsBRL(soma)):'';
    res.className='flt-resumo'+(pk.length?' on':'');
  }
}
function bcmsLimpaFiltros(sfx){
  ['f-fonte','f-acao','f-nd','f-idade'].forEach(function(id){
    var e=document.getElementById(id+'-'+sfx); if(e)e.value='';});
  var b=document.getElementById('q-tab-emtela-'+sfx); if(b)b.value='';
  bcmsFiltra(sfx);
}
/* Descreve os filtros ativos, p/ registrar no cabeçalho da planilha exportada. */
function bcmsFiltrosAtivos(tid){
  var m=/^tab-emtela-(.+)$/.exec(tid||''); if(!m) return '';
  var sfx=m[1];
  var v=function(id){var e=document.getElementById(id+'-'+sfx); return e?e.value:'';};
  var pk=[];
  if(v('f-fonte'))pk.push('Fonte: '+v('f-fonte'));
  if(v('f-acao'))pk.push('Ação: '+v('f-acao'));
  if(v('f-nd'))pk.push('ND: '+v('f-nd'));
  var fi=v('f-idade'); if(fi)pk.push('Idade: '+({v:'até 30 dias',a:'31 a 60 dias',r:'mais de 60 dias'}[fi]||fi));
  var q=v('q-tab-emtela'); if(q)pk.push('Busca: "'+q+'"');
  return pk.length?pk.join(' · '):'';
}

function bcmsExportTable(btn,tid,filename){
  var container=document.getElementById(tid);if(!container)return;
  var table=container.tagName==='TABLE'?container:container.querySelector('table');if(!table)return;
  filename=(filename||'creditos_em_tela')+'_'+(new Date().toISOString().slice(0,10));
  var ths=table.querySelectorAll('thead th');var headers=[];var colTypes=[];
  ths.forEach(function(th){
    var txt=th.textContent.replace('▲','').replace('▼','').trim();
    if(txt&&txt!=='Ação'&&txt!=='AÇÃO'){
      headers.push(txt);
      var u=txt.toUpperCase();
      if(u.indexOf('NOTA')>-1||u.indexOf('NC')>-1||u.indexOf('FONTE')>-1||u.indexOf('UASG')>-1||
         u.indexOf('RECEBIDO EM')>-1||u.indexOf('DATA')>-1||u.indexOf('EMISSÃO')>-1||u.indexOf('EMISSAO')>-1||
         u.indexOf('DESCRIÇÃO')>-1||u.indexOf('DESCRICAO')>-1||u.indexOf('OBJETO')>-1||u.indexOf('AÇÃO')>-1||
         u.indexOf('ACAO')>-1||u.indexOf('ND')>-1||u.indexOf('PI')>-1||u.indexOf('OPERAÇÃO')>-1||
         u.indexOf('OPERACAO')>-1||u.indexOf('EMITENTE')>-1||u.indexOf('ORGANIZAÇÃO')>-1||u.indexOf('OMDS')>-1||
         u.indexOf('DIA ANTERIOR')>-1){
        colTypes.push('String');
      } else if(u.indexOf('DIA')>-1||u.indexOf('IDADE')>-1||u.indexOf('CÉLULA')>-1||u.indexOf('CELULA')>-1||u.indexOf('Nº')>-1||u.indexOf('POS')>-1||u.indexOf('RANK')>-1){
        colTypes.push('Integer');
      } else if(u.indexOf('%')>-1||u.indexOf('TAXA')>-1){
        colTypes.push('Percent');
      } else if(u.indexOf('R$')>-1||u.indexOf('VALOR')>-1||u.indexOf('CRÉDITO')>-1||u.indexOf('CREDITO')>-1||
                 u.indexOf('EMPENHADO')>-1||u.indexOf('LIQUIDADO')>-1||u.indexOf('PAGO')>-1||
                 u.indexOf('PROVISÃO')>-1||u.indexOf('PROVISAO')>-1||u.indexOf('SALDO')>-1||
                 u.indexOf('REDUÇ')>-1||u.indexOf('LÍQUIDO')>-1||u.indexOf('LIQUIDO')>-1||
                 u.indexOf('RECEBIDO')>-1||u.indexOf('DISP')>-1){
        colTypes.push('Currency');
      } else {
        colTypes.push('String');
      }
    }
  });
  var trs=table.querySelectorAll('tbody tr');var rows=[];
  trs.forEach(function(tr){if(tr.style.display==='none')return;
    var cells=tr.querySelectorAll('td');var rowData=[];
    cells.forEach(function(td,idx){if(idx>=headers.length)return;
      var sortVal=td.getAttribute('data-sort');var exp=colTypes[idx]||'String';
      /* guarda: conteúdo com cara de data (dd/mm/aa) é texto, nunca número —
         senão a coluna "Recebido" (data) era somada como se fosse moeda. */
      var _txtCel=td.textContent.trim();
      if(/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(_txtCel)) exp='String';
      var fullText=td.getAttribute('data-full-desc')||td.getAttribute('title')||td.textContent.replace('›','').trim();
      if(!td.getAttribute('data-full-desc')&&!td.getAttribute('title')){fullText=td.textContent.replace('›','').trim();}
      if(exp==='String'){
        rowData.push({v:fullText,t:'String',s:'Default'});
      } else if(exp==='Integer'){
        var numVal=sortVal!==null&&!isNaN(parseFloat(sortVal))?parseInt(sortVal,10):parseInt(fullText.replace(/\D/g,''),10);
        if(isNaN(numVal)||numVal<0){rowData.push({v:fullText||'—',t:'String',s:'Default'});}
        else{rowData.push({v:numVal,t:'Number',s:'Integer'});}
      } else if(exp==='Currency'){
        var numVal=sortVal!==null&&!isNaN(parseFloat(sortVal))?parseFloat(sortVal):parseFloat(fullText.replace(/[^\d,-]/g,'').replace(',','.'));
        rowData.push({v:isNaN(numVal)?0.0:numVal,t:'Number',s:'Currency'});
      } else if(exp==='Percent'){
        var numVal=sortVal!==null&&!isNaN(parseFloat(sortVal))?parseFloat(sortVal):parseFloat(fullText.replace(/[^\d,-]/g,'').replace(',','.'));
        rowData.push({v:isNaN(numVal)?0.0:numVal/100.0,t:'Number',s:'Percent'});
      } else {
        rowData.push({v:fullText,t:'String',s:'Default'});
      }
    });
    if(rowData.length)rows.push(rowData);
  });
  var xml='<?xml version="1.0" encoding="UTF-8"?>\n'+
   '<?mso-application progid="Excel.Sheet"?>\n'+
   '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n'+
   ' xmlns:o="urn:schemas-microsoft-com:office:office"\n'+
   ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n'+
   ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n'+
   '<Styles>\n'+
   ' <Style ss:ID="Default" ss:Name="Normal"><Font ss:FontName="Calibri" ss:Size="11" ss:Color="#000000"/><Alignment ss:Vertical="Center"/></Style>\n'+
   ' <Style ss:ID="Header"><Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1C4A73" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>\n'+
   ' <Style ss:ID="Title"><Font ss:FontName="Calibri" ss:Size="14" ss:Bold="1" ss:Color="#1C4A73"/><Alignment ss:Vertical="Center"/></Style>\n'+
   ' <Style ss:ID="Currency"><NumberFormat ss:Format="&quot;R$&quot; #,##0.00"/><Alignment ss:Horizontal="Right" ss:Vertical="Center"/></Style>\n'+
   ' <Style ss:ID="Integer"><NumberFormat ss:Format="#,##0"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>\n'+
   ' <Style ss:ID="Percent"><NumberFormat ss:Format="0.0%"/><Alignment ss:Horizontal="Right" ss:Vertical="Center"/></Style>\n'+
   ' <Style ss:ID="Bold"><Font ss:FontName="Calibri" ss:Bold="1"/></Style>\n'+
   '</Styles>\n'+
   '<Worksheet ss:Name="Créditos em Tela">\n'+
   '<Table ss:DefaultRowHeight="20">\n';
  headers.forEach(function(h){
    var u=h.toUpperCase();var w=120;
    if(u.indexOf('DESCRIÇÃO')>-1||u.indexOf('DESCRICAO')>-1||u.indexOf('OBJETO')>-1||u.indexOf('APLICAÇÃO')>-1){w=380;}
    else if(u.indexOf('ORGANIZAÇÃO')>-1||u.indexOf('NOME')>-1){w=220;}
    else if(u.indexOf('NC')>-1||u.indexOf('NOTA')>-1){w=140;}
    else if(u.indexOf('DIA')>-1||u.indexOf('IDADE')>-1){w=100;}
    xml+=' <Column ss:AutoFitWidth="1" ss:Width="'+w+'"/>\n';
  });
  xml+=' <Row ss:Height="26"><Cell ss:StyleID="Title" ss:MergeAcross="'+(headers.length-1)+'"><Data ss:Type="String">6º BATALHÃO DE ENGENHARIA DE CONSTRUÇÃO — RELATÓRIO DE CRÉDITOS DISPONÍVEIS</Data></Cell></Row>\n';
  var _flt=(typeof bcmsFiltrosAtivos==='function')?bcmsFiltrosAtivos(tid):'';
  xml+=' <Row ss:Height="18"><Cell ss:MergeAcross="'+(headers.length-1)+'"><Data ss:Type="String">Posição extraída do Tesouro Gerencial / SIAFI · '+new Date().toLocaleDateString('pt-BR')+' · '+rows.length+' registro(s)'+(_flt?' · FILTROS APLICADOS — '+_flt:' · sem filtros (lista completa)')+'</Data></Cell></Row>\n';
  xml+=' <Row ss:Height="10"/>\n';
  xml+=' <Row ss:Height="24">\n';
  headers.forEach(function(h){xml+='  <Cell ss:StyleID="Header"><Data ss:Type="String">'+bcmsEscXml(h)+'</Data></Cell>\n';});
  xml+=' </Row>\n';
  rows.forEach(function(r){xml+=' <Row ss:Height="20">\n';
    r.forEach(function(c){
      if(c.t==='Number'){xml+='  <Cell ss:StyleID="'+(c.s||'Currency')+'"><Data ss:Type="Number">'+c.v+'</Data></Cell>\n';}
      else{xml+='  <Cell><Data ss:Type="String">'+bcmsEscXml(c.v)+'</Data></Cell>\n';}
    });
    xml+=' </Row>\n';
  });
  xml+='</Table></Worksheet></Workbook>';
  var blob=new Blob([xml],{type:'application/vnd.ms-excel;charset=utf-8;'});
  var link=document.createElement('a');link.href=URL.createObjectURL(blob);
  link.download=filename+'.xls';document.body.appendChild(link);link.click();document.body.removeChild(link);
  bcmsToast('📊 Planilha Excel gerada com sucesso! Download iniciado.');
}

/* ---- Detalhamento por NC "em tela", dividido nas ETAPAS do crédito ----
   Card principal = SÓ o que está realmente disponível em tela desta NC.
   Abas: Em tela | Histórico do PI | Liquidação | Pagamento.
   (Antes o clique abria a célula inteira, misturando NCs de objetos diferentes
    que só compartilham o mesmo PI.) */
var BTELA=null;
function bcmsTela(row){
  var t=TELADATA[row.getAttribute('data-tela')];
  if(!t){return bcmsCel(row);}   /* fallback: comportamento antigo */
  BTELA=t;
  bcmsTelaAba('tela');
  var m=document.getElementById('modal');
  m.classList.add('open');m.setAttribute('aria-hidden','false');
  var x=document.querySelector('.modal-x');if(x)x.focus();
}
function bcmsTelaAba(aba){
  var t=BTELA;if(!t)return;
  var c=CELDATA[t.cid]||{};
  var fonte=t.u==='OGU'?'OGU (Orçamento Geral da União)':(t.u==='FEx'?'FEx (Fundo do Exército)':(t.u||'—'));
  var ncCurta=(String(t.nc).match(/NC(\d+)$/)||[])[1];
  var h='<h3 id="modal-title">'+(ncCurta?'NC '+bcmsEsc(ncCurta):bcmsEsc(t.nc))+'</h3>';
  h+='<p class="m-sub">'+bcmsEsc(t.uasg+' · '+fonte+' · Ação '+t.acao+' · PI '+t.pi+' · ND '+t.nd+(t.ndnome?' — '+t.ndnome:''))+'</p>';
  /* etapas */
  var abas=[['tela','🟢 Em tela'],['hist','📜 Histórico do PI'],['liq','🧾 Liquidação'],['pag','💰 Pagamento']];
  h+='<div class="m-etapas" role="tablist">';
  abas.forEach(function(a){
    h+='<button class="m-etapa'+(a[0]===aba?' on':'')+'" role="tab" aria-selected="'+(a[0]===aba)+'" onclick="bcmsTelaAba(\''+a[0]+'\')">'+a[1]+'</button>';
  });
  h+='</div><div class="m-etapa-body">';
  if(aba==='tela'){
    var idade=(t.dias===null||t.dias===undefined)?'—':(t.dias+' dia'+(t.dias===1?'':'s'));
    h+='<div class="m-tela-card"><span class="m-tela-lbl">Disponível em tela nesta NC</span>'
      +'<b class="m-tela-val">'+bcmsBRL(t.v)+'</b>'
      +'<span class="m-tela-meta">Recebido em '+bcmsEsc(t.dia||'—')+' · há '+idade+(t.emit?' · Emitente '+bcmsEsc(t.emit):'')+'</span></div>';
    if(t.org){
      h+='<div class="m-org"><div class="m-org-h">↪ Crédito recebido por <b>mudança de ND</b>'
        +(t.nd_de?' (de '+bcmsEsc(t.nd_de)+' para '+bcmsEsc(t.nd)+')':'')+'</div>'
        +'<div class="m-org-b"><span>Objeto original — NC '+bcmsEsc(t.org.nc)
        +(t.org.dia?' de '+bcmsEsc(t.org.dia):'')+(t.org.emit?' · emitente '+bcmsEsc(t.org.emit):'')+'</span>'
        +'<p>'+bcmsEsc(t.org.obj||'—')+'</p></div></div>';
    }
    h+='<div class="m-ncs-h">Descrição desta NC</div><div class="m-nc-desc big">'+bcmsEsc(t.obj||'—')+'</div>';
    if(t.op) h+='<p class="m-formula">Operação: '+bcmsEsc(t.op)+'</p>';
  } else if(aba==='hist'){
    h+='<p class="m-formula">Todas as movimentações da célula <b>'+bcmsEsc(t.acao+' · PI '+t.pi+' · ND '+t.nd)+'</b> — recebimentos, detalhamentos, mudanças de ND e anulações.</p>';
    h+='<div class="m-kpis"><span>Recebido (líq)<b>'+bcmsBRL(c.r||0)+'</b></span><span>Empenhado<b>'+bcmsBRL(c.e||0)+'</b></span><span class="ok">Disponível<b>'+bcmsBRL(c.d||0)+'</b></span></div>';
    var itens='';
    (c.ncs||[]).forEach(function(n){
      if(!n[0])return;var neg=n[2]<0;var meta=[];
      if(n[4])meta.push('Emitente '+bcmsEsc(n[4]));
      if(n[1])meta.push(bcmsEsc(n[1]));
      if(n[5])meta.push('em '+bcmsEsc(n[5]));
      var ehEsta=(n[0]===t.nc);
      itens+='<div class="m-nc'+(ehEsta?' m-nc-esta':'')+'"><div class="m-nc-h"><span class="m-nc-num">'+bcmsEsc(n[0])+(ehEsta?' <i class="m-nc-tag">esta NC</i>':'')+'</span><span class="m-nc-val'+(neg?' neg':'')+'">'+bcmsBRL(n[2])+'</span></div>';
      if(meta.length)itens+='<div class="m-nc-op">'+meta.join(' · ')+'</div>';
      if(n[3])itens+='<div class="m-nc-desc">'+bcmsEsc(n[3])+'</div>';
      itens+='</div>';
    });
    h+='<div class="m-ncs">'+(itens||'<p class="vazio">Sem movimentações.</p>')+'</div>';
  } else if(aba==='liq'){
    var emp=c.e||0,liq=c.l||0;
    h+='<div class="m-kpis"><span>Empenhado<b>'+bcmsBRL(emp)+'</b></span><span class="ok">Liquidado<b>'+bcmsBRL(liq)+'</b></span><span>A liquidar<b>'+bcmsBRL(Math.max(0,emp-liq))+'</b></span></div>';
    h+=bcmsBarra(liq,emp,'Liquidado sobre o empenhado');
    h+='<p class="m-formula">Liquidação é a etapa em que a despesa é atestada (bem/serviço entregue). Valores da célula <b>'+bcmsEsc(t.acao+' · PI '+t.pi+' · ND '+t.nd)+'</b> — o SIAFI não segrega liquidação por NC individual.</p>';
  } else {
    var liq2=c.l||0,pag=c.p||0;
    h+='<div class="m-kpis"><span>Liquidado<b>'+bcmsBRL(liq2)+'</b></span><span class="ok">Pago<b>'+bcmsBRL(pag)+'</b></span><span>A pagar<b>'+bcmsBRL(Math.max(0,liq2-pag))+'</b></span></div>';
    h+=bcmsBarra(pag,liq2,'Pago sobre o liquidado');
    h+='<p class="m-formula">Pagamento é a quitação efetiva da despesa liquidada. Valores da célula <b>'+bcmsEsc(t.acao+' · PI '+t.pi+' · ND '+t.nd)+'</b> — o SIAFI não segrega pagamento por NC individual.</p>';
  }
  h+='</div>';
  document.getElementById('modal-body').innerHTML=h;
}
function bcmsBarra(v,total,rot){
  var p=total>0?Math.min(100,Math.max(0,v/total*100)):0;
  return '<div class="m-barra"><div class="m-barra-t"><span>'+bcmsEsc(rot)+'</span><b>'+p.toFixed(1)+'%</b></div>'
    +'<div class="m-barra-track"><div class="m-barra-fill" style="width:'+p.toFixed(1)+'%"></div></div></div>';
}

function bcmsCel(row){
  var d=CELDATA[row.getAttribute('data-cel')];if(!d)return;
  var h='<h3 id="modal-title">'+bcmsEsc(d.t)+'</h3>';
  var fonte=d.u==='OGU'?'OGU (Orçamento Geral da União)':(d.u==='FEx'?'FEx (Fundo do Exército)':(d.u||'—'));
  h+='<div class="m-ficha">'
    +'<span>UASG (Executora)<b>'+bcmsEsc(d.uasg||'—')+'</b></span>'
    +'<span>Fonte<b>'+bcmsEsc(fonte)+'</b></span>'
    +'<span>Ação Governo<b>'+bcmsEsc(d.acao||'—')+'</b></span>'
    +'<span class="wide">PI (Plano Interno)<b>'+bcmsEsc(d.pi||'—')+(d.pinome?' — '+bcmsEsc(d.pinome):'')+'</b></span>'
    +'<span class="wide">ND (Natureza de Despesa)<b>'+bcmsEsc(d.nd||'—')+(d.ndnome?' — '+bcmsEsc(d.ndnome):'')+'</b></span>'
    +'</div>';
  h+='<div class="m-kpis"><span>Recebido (líq)<b>'+bcmsBRL(d.r)+'</b></span><span>Empenhado<b>'+bcmsBRL(d.e)+'</b></span><span>Liquidado<b>'+bcmsBRL(d.l||0)+'</b></span><span>Pago<b>'+bcmsBRL(d.p||0)+'</b></span><span class="ok">Crédito Disponível<b>'+bcmsBRL(d.d)+'</b></span></div>';
  h+='<p class="m-formula">Recebido (líq) − Empenhado = Crédito Disponível · Empenhado ≥ Liquidado ≥ Pago</p>';
  var itens='';
  d.ncs.forEach(function(n){
    if(!n[0])return;
    var neg=n[2]<0;
    var meta=[];
    if(n[4]) meta.push('Emitente '+bcmsEsc(n[4]));
    if(n[1]) meta.push(bcmsEsc(n[1]));
    if(n[5]) meta.push('em '+bcmsEsc(n[5]));
    itens+='<div class="m-nc"><div class="m-nc-h"><span class="m-nc-num">'+bcmsEsc(n[0])+'</span><span class="m-nc-val'+(neg?' neg':'')+'">'+bcmsBRL(n[2])+'</span></div>';
    if(meta.length) itens+='<div class="m-nc-op">'+meta.join(' · ')+'</div>';
    if(n[3]) itens+='<div class="m-nc-desc">'+bcmsEsc(n[3])+'</div>';
    itens+='</div>';
  });
  var nq=d.ncs.filter(function(n){return n[0];}).length;
  h+='<div class="m-ncs-h">Notas de crédito da célula ('+nq+')</div>';
  h+='<div class="m-ncs">'+(itens||'<p class="vazio">Sem notas de crédito para detalhar.</p>')+'</div>';
  document.getElementById('modal-body').innerHTML=h;
  var m=document.getElementById('modal');
  m.classList.add('open');
  m.setAttribute('aria-hidden','false');
  var x=document.querySelector('.modal-x');
  if(x) x.focus();
}

function bcmsDay(row){
  var d=DAYDATA[row.getAttribute('data-day')];if(!d)return;
  var h='<h3 id="modal-title">Movimentação de '+bcmsEsc(d.d)+'</h3>';
  h+='<div class="m-kpis"><span>Nº de NC<b>'+d.n+'</b></span><span>Recebido<b class="col-pos">'+bcmsBRL(d.rec)+'</b></span><span>Reduções<b class="col-neg">'+bcmsBRL(d.red)+'</b></span><span class="ok">Líquido<b>'+bcmsBRL(d.liq)+'</b></span></div>';
  h+='<div class="m-ncs-h">Notas de crédito do dia ('+d.ncs.length+')</div>';
  var itens='';
  d.ncs.forEach(function(n){
    var neg=n[3]<0;
    itens+='<div class="m-nc"><div class="m-nc-h"><span class="m-nc-num">'+bcmsEsc(n[0])+' <span class="pill-fonte">'+bcmsEsc(n[1])+'</span></span><span class="m-nc-val'+(neg?' neg':'')+'">'+bcmsBRL(n[3])+'</span></div>';
    if(n[2]) itens+='<div class="m-nc-op">'+bcmsEsc(n[2])+'</div>';
    if(n[4]) itens+='<div class="m-nc-desc">'+bcmsEsc(n[4])+'</div>';
    itens+='</div>';
  });
  h+='<div class="m-ncs">'+(itens||'<p class="vazio">Sem NC neste dia.</p>')+'</div>';
  document.getElementById('modal-body').innerHTML=h;
  var m=document.getElementById('modal');
  m.classList.add('open');
  m.setAttribute('aria-hidden','false');
  var x=document.querySelector('.modal-x');
  if(x) x.focus();
}

function bcmsCelClose(){
  var m=document.getElementById('modal');
  if(m){ m.classList.remove('open'); m.setAttribute('aria-hidden','true'); }
}

document.addEventListener('keydown', function(e){
  if(e.key==='Escape') bcmsCelClose();
});

/* ==========================================================================
   HISTÓRICO CONSOLIDADO DE NOTAS DE CRÉDITO & DRILL-DOWN MODAL
   ========================================================================== */
var HIST_PAGE = 1;
var HIST_PAGE_SIZE = 50;
var HIST_FILTERED = [];
var HIST_SORT_COL = 'dt';
var HIST_SORT_DIR = 'desc';
var HIST_INITIALIZED = false;

function bcmsGetHistItems(){
  if(typeof HISTDATA === 'undefined' || !HISTDATA) return [];
  if(Array.isArray(HISTDATA.items)) return HISTDATA.items;
  if(Array.isArray(HISTDATA)) return HISTDATA;
  if(HISTDATA.by_id) return Object.values(HISTDATA.by_id);
  return Object.values(HISTDATA);
}

function bcmsInitHistorico(){
  var raw = bcmsGetHistItems();
  if(!raw || !raw.length) return;
  HIST_FILTERED = raw.slice();
  bcmsSortFilteredArray();
  bcmsAtualizaKPIs(HIST_FILTERED);
  bcmsRenderHistorico(1);
  HIST_INITIALIZED = true;
}

function bcmsSortFilteredArray(){
  var col = HIST_SORT_COL;
  var dir = HIST_SORT_DIR;
  HIST_FILTERED.sort(function(a, b){
    var va = a[col], vb = b[col];
    if(col === 'prov' || col === 'cred' || col === 'emp' || col === 'liq' || col === 'pag'){
      va = parseFloat(va) || 0;
      vb = parseFloat(vb) || 0;
    } else if(col === 'fav'){
      va = a.om_sigla || a.fav_nome || '';
      vb = b.om_sigla || b.fav_nome || '';
    } else if(col === 'emit'){
      va = a.emit_nome || a.emit_cod || '';
      vb = b.emit_nome || b.emit_cod || '';
    } else if(col === 'ptres'){
      va = (a.ptres || '') + ' ' + (a.nd || '');
      vb = (b.ptres || '') + ' ' + (b.nd || '');
    } else {
      va = String(va || '').toLowerCase();
      vb = String(vb || '').toLowerCase();
    }
    if(va < vb) return dir === 'asc' ? -1 : 1;
    if(va > vb) return dir === 'asc' ? 1 : -1;
    return 0;
  });
}

function bcmsSortHistorico(col, th){
  if(HIST_SORT_COL === col){
    HIST_SORT_DIR = (HIST_SORT_DIR === 'asc') ? 'desc' : 'asc';
  } else {
    HIST_SORT_COL = col;
    HIST_SORT_DIR = (col === 'dt' || col === 'prov' || col === 'cred') ? 'desc' : 'asc';
  }
  var thead = th.closest('thead');
  if(thead){
    thead.querySelectorAll('th').forEach(function(h){
      h.setAttribute('aria-sort', 'none');
      var s = h.querySelector('.sort');
      if(s) s.textContent = '';
    });
    th.setAttribute('aria-sort', HIST_SORT_DIR === 'asc' ? 'ascending' : 'descending');
    var s = th.querySelector('.sort');
    if(s) s.textContent = HIST_SORT_DIR === 'asc' ? ' ▲' : ' ▼';
  }
  bcmsSortFilteredArray();
  bcmsRenderHistorico(HIST_PAGE);
}

function bcmsFiltraHistorico(){
  var raw = bcmsGetHistItems();
  if(!raw || !raw.length) return;
  var v = function(id){ var e = document.getElementById(id); return e ? e.value.trim() : ''; };
  var fOm = v('flt-hist-om');
  var fPer = v('flt-hist-periodo');
  var fFonte = v('flt-hist-fonte');
  var fPtres = v('flt-hist-ptres');
  var fFaixa = v('flt-hist-faixa');
  var q = (v('flt-hist-busca') || '').toLowerCase();

  HIST_FILTERED = raw.filter(function(it){
    if(fOm && it.om_sigla !== fOm) return false;
    if(fPer){
      if(fPer.startsWith('T')){
        if(it.tri !== fPer) return false;
      } else {
        if(it.mes !== fPer) return false;
      }
    }
    if(fFonte && it.fonte !== fFonte) return false;
    if(fPtres && it.ptres !== fPtres && it.acao !== fPtres) return false;
    if(fFaixa){
      if(fFaixa === 'saldo_pos' && it.cred <= 0.01) return false;
      if(fFaixa === 'parcial' && (it.status_slug !== 'parcial')) return false;
      if(fFaixa === 'zerada' && (it.cred > 0.01 || it.status_slug === 'canc' || it.status_slug === 'detalhada')) return false;
      if(fFaixa === 'detalhada' && it.status_slug !== 'detalhada') return false;
      if(fFaixa === 'canc' && it.status_slug !== 'canc') return false;
    }
    if(q){
      var match = (it.nc && it.nc.toLowerCase().indexOf(q) > -1) ||
                  (it.obj && it.obj.toLowerCase().indexOf(q) > -1) ||
                  (it.pi && it.pi.toLowerCase().indexOf(q) > -1) ||
                  (it.pi_desc && it.pi_desc.toLowerCase().indexOf(q) > -1) ||
                  (it.fav_nome && it.fav_nome.toLowerCase().indexOf(q) > -1) ||
                  (it.om_sigla && it.om_sigla.toLowerCase().indexOf(q) > -1) ||
                  (it.emit_nome && it.emit_nome.toLowerCase().indexOf(q) > -1) ||
                  (it.emit_cod && String(it.emit_cod).indexOf(q) > -1) ||
                  (it.fav_cod && String(it.fav_cod).indexOf(q) > -1) ||
                  (it.ptres && it.ptres.toLowerCase().indexOf(q) > -1) ||
                  (it.nd && it.nd.toLowerCase().indexOf(q) > -1);
      if(!match) return false;
    }
    return true;
  });

  bcmsSortFilteredArray();
  bcmsAtualizaKPIs(HIST_FILTERED);
  bcmsRenderHistorico(1);
}

function bcmsAtualizaKPIs(list){
  var total = list.length;
  var distinctSet = {};
  var totProv = 0, totEmp = 0, totCred = 0;
  for(var i = 0; i < list.length; i++){
    var it = list[i];
    distinctSet[it.nc] = true;
    totProv += it.prov;
    totEmp += it.emp;
    totCred += it.cred;
  }
  var distinctCount = Object.keys(distinctSet).length;

  var k1 = document.getElementById('kpi-hist-total');
  if(k1) k1.textContent = distinctCount.toLocaleString('pt-BR');

  var k2 = document.getElementById('kpi-hist-prov');
  if(k2) k2.textContent = bcmsBRL(totProv);

  var k3 = document.getElementById('kpi-hist-emp');
  if(k3) k3.textContent = bcmsBRL(totEmp);

  var k4 = document.getElementById('kpi-hist-cred');
  if(k4) k4.textContent = bcmsBRL(totCred);

  var badge = document.getElementById('cnt-hist-ncs');
  if(badge){
    var totalGlobal = (HISTDATA && HISTDATA.items) ? HISTDATA.items.length : total;
    var ativo = (total !== totalGlobal);
    badge.textContent = (ativo ? 'Filtrado: ' : 'Exibindo ') + total + ' de ' + totalGlobal + ' Notas de Crédito (' + distinctCount + ' distintas)';
  }

  /* tfoot */
  var tfLabel = document.getElementById('tf-hist-label');
  if(tfLabel) tfLabel.textContent = (total !== (HISTDATA && HISTDATA.items ? HISTDATA.items.length : total) ? 'FILTRADO · ' : 'TOTAL · ') + total + ' Nota(s) de Crédito no histórico';
  var tfProv = document.getElementById('tf-hist-prov');
  if(tfProv) tfProv.textContent = bcmsBRL(totProv);
  var tfCred = document.getElementById('tf-hist-cred');
  if(tfCred) tfCred.textContent = bcmsBRL(totCred);

  /* chip de resumo dos filtros */
  var v = function(id){ var e = document.getElementById(id); return e ? e.value.trim() : ''; };
  var fOm = v('flt-hist-om');
  var fPer = v('flt-hist-periodo');
  var fFonte = v('flt-hist-fonte');
  var fPtres = v('flt-hist-ptres');
  var fFaixa = v('flt-hist-faixa');
  var q = v('flt-hist-busca');

  var res = document.getElementById('flt-hist-res');
  if(res){
    var pk = [];
    if(fOm) pk.push(fOm);
    if(fFonte) pk.push('Fonte ' + fFonte);
    if(fPtres) pk.push('Ação ' + fPtres);
    if(fPer) pk.push(fPer);
    if(fFaixa) pk.push({saldo_pos:'Com Saldo',parcial:'Parcial',zerada:'Zerada',canc:'Cancelada'}[fFaixa] || fFaixa);
    if(q) pk.push('“' + q + '”');
    res.textContent = pk.length ? (pk.join(' · ') + ' — ' + bcmsBRL(totCred)) : '';
    res.className = 'flt-resumo' + (pk.length ? ' on' : '');
  }
}

function bcmsLimpaFiltrosHistorico(){
  ['flt-hist-om','flt-hist-periodo','flt-hist-fonte','flt-hist-ptres','flt-hist-faixa'].forEach(function(id){
    var el = document.getElementById(id);
    if(el) el.value = '';
  });
  var b = document.getElementById('flt-hist-busca');
  if(b) b.value = '';
  bcmsFiltraHistorico();
  bcmsToast('Filtros do histórico redefinidos.');
}

function bcmsRenderHistorico(page){
  HIST_PAGE = page;
  var total = HIST_FILTERED.length;
  var totalPages = Math.max(1, Math.ceil(total / HIST_PAGE_SIZE));
  if(HIST_PAGE > totalPages) HIST_PAGE = totalPages;
  if(HIST_PAGE < 1) HIST_PAGE = 1;

  var start = (HIST_PAGE - 1) * HIST_PAGE_SIZE;
  var end = Math.min(start + HIST_PAGE_SIZE, total);
  var pageItems = HIST_FILTERED.slice(start, end);

  var tbody = document.getElementById('tbody-hist-ncs');
  if(!tbody) return;

  if(pageItems.length === 0){
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:48px 16px;color:var(--ink-muted);"><span style="font-size:2rem;display:block;margin-bottom:8px;">🔍</span>Nenhuma Nota de Crédito encontrada para os filtros selecionados.<br><button type="button" class="flt-limpa" style="margin-top:12px;" onclick="bcmsLimpaFiltrosHistorico()">✕ Limpar Filtros</button></td></tr>';
  } else {
    var rowsHtml = '';
    for(var i = 0; i < pageItems.length; i++){
      var item = pageItems[i];
      var statusCls = 'status-' + item.status_slug;
      var emitNomeCurto = (item.emit_nome && item.emit_nome.length > 22) ? item.emit_nome.substring(0, 22) + '…' : (item.emit_nome || '—');
      var descCompleta = item.obj || '';
      var descResumo = descCompleta.length > 118 ? (descCompleta.substring(0, 118) + '…') : descCompleta;
      var ncFull = String(item.nc || '');
      var mNc = ncFull.match(/NC(\d+)$/);
      var ncLbl = mNc ? ('<span class="nc-num">NC ' + mNc[1] + '</span> <span class="nc-ug">· ' + bcmsEsc(ncFull.substring(0,6)) + '</span>') : ('<span class="nc-num">' + bcmsEsc(ncFull) + '</span>');
      var acaoNd = (item.ptres && item.nd) ? (item.ptres + ' · ' + item.nd) : (item.ptres || item.nd || '—');

      rowsHtml += '<tr class="cel-row" data-hid="' + bcmsEsc(item.hid) + '" tabindex="0" role="button" onclick="bcmsDetalheNC(\'' + bcmsEsc(item.hid) + '\')" '
        + 'title="Clique para abrir o detalhamento completo da NC ' + bcmsEsc(ncFull) + '" '
        + 'onkeydown="if(event.key===\'Enter\'||event.key===\' \'){event.preventDefault();bcmsDetalheNC(\'' + bcmsEsc(item.hid) + '\')}">'
        + '<td><span class="pill-fonte">' + bcmsEsc(item.fonte) + '</span></td>'
        + '<td class="mono2" title="' + bcmsEsc(ncFull) + '">' + ncLbl + '</td>'
        + '<td title="' + bcmsEsc(item.fav_nome) + '"><span class="ug-pill fav">' + bcmsEsc(item.fav_cod) + '</span> <b>' + bcmsEsc(item.om_sigla) + '</b></td>'
        + '<td class="mono2">' + bcmsEsc(acaoNd) + '</td>'
        + '<td class="obj" title="' + bcmsEsc(descCompleta) + '" data-full-desc="' + bcmsEsc(descCompleta) + '">' + bcmsEsc(descResumo) + '</td>'
        + '<td title="' + bcmsEsc(item.emit_nome) + '"><span class="ug-pill emit">' + bcmsEsc(item.emit_cod) + '</span> <small>' + bcmsEsc(emitNomeCurto) + '</small></td>'
        + '<td class="mono2">' + bcmsEsc(item.dia || '—') + '</td>'
        + '<td class="num" data-sort="' + item.prov.toFixed(2) + '">' + bcmsBRL(item.prov) + '</td>'
        + '<td class="num anchor" data-sort="' + item.cred.toFixed(2) + '">' + bcmsBRL(item.cred) + '</td>'
        + '<td class="num"><span class="pill-status ' + statusCls + '">' + bcmsEsc(item.status) + '</span><i class="chev" aria-hidden="true">›</i></td>'
        + '</tr>';
    }
    tbody.innerHTML = rowsHtml;
  }

  var pagInfo = document.getElementById('pag-info-txt');
  if(pagInfo){
    if(total === 0){
      pagInfo.textContent = 'Página 0 de 0 (0 registros)';
    } else {
      pagInfo.textContent = 'Página ' + HIST_PAGE + ' de ' + totalPages + ' (Exibindo ' + (start + 1) + '–' + end + ' de ' + total + ')';
    }
  }

  var btnAnt = document.getElementById('btn-pag-ant');
  if(btnAnt) btnAnt.disabled = (HIST_PAGE <= 1);
  var btnProx = document.getElementById('btn-pag-prox');
  if(btnProx) btnProx.disabled = (HIST_PAGE >= totalPages);
}

function bcmsPaginaHistorico(delta){
  var novaPagina = HIST_PAGE + delta;
  bcmsRenderHistorico(novaPagina);
  var scrollCont = document.getElementById('scroll-hist-ncs');
  if(scrollCont) scrollCont.scrollIntoView({behavior:'smooth', block:'nearest'});
}

function bcmsDetalheNC(hid){
  var item = null;
  if(typeof HISTDATA !== 'undefined' && HISTDATA){
    if(HISTDATA.by_id && HISTDATA.by_id[hid]){
      item = HISTDATA.by_id[hid];
    } else if(HISTDATA[hid]){
      item = HISTDATA[hid];
    } else {
      var all = bcmsGetHistItems();
      for(var i = 0; i < all.length; i++){
        if(all[i].hid === hid || all[i].nc === hid){
          item = all[i];
          break;
        }
      }
    }
  }
  if(!item && typeof CATRDATA !== 'undefined' && CATRDATA && CATRDATA.linhas){
    var matching = [];
    for(var c = 0; c < CATRDATA.linhas.length; c++){
      var cl = CATRDATA.linhas[c];
      if(cl.nc === hid || cl.hid === hid || (cl.nc && cl.nc.indexOf(hid) !== -1)){
        matching.push(cl);
      }
    }
    if(matching.length > 0){
      var cl = matching[0];
      var totProv = 0, totEmp = 0, totCred = 0, totLiq = 0, totPag = 0;
      var subItens = [];
      matching.forEach(function(m){
        totProv += (m.prov || 0);
        totEmp  += (m.emp || 0);
        totCred += (m.cred || 0);
        totLiq  += (m.liq || 0);
        totPag  += (m.pag || 0);
        subItens.push({acao: m.acao, pi: m.pi, nd: m.nd, val: m.prov || m.cred || 0});
      });
      item = {
        hid: cl.nc,
        nc: cl.nc,
        op: cl.op || 'DESCENTRALIZACAO DE CREDITO',
        dia: cl.dia || '—',
        dias: null,
        status: (totCred > 0.01 ? 'Disponível' : 'Executado Integral'),
        status_slug: (totCred > 0.01 ? 'ok' : 'danger'),
        emit_cod: cl.emit || '160073',
        emit_nome: cl.emit_nome || 'DIRETORIA DE GESTAO ORCAMENTARIA - GESTOR',
        fav_cod: cl.ug || '—',
        fav_nome: cl.ug_nome || ('UG ' + cl.ug),
        om_sigla: cl.ug_nome || cl.ug,
        ptres: cl.acao || '21EM',
        acao_desc: 'Ação Governamental ' + (cl.acao || '21EM') + ' (Operação Catrimani II)',
        fonte: (String(cl.ug).indexOf('167') === 0) ? 'FEx' : 'OGU',
        nd: cl.nd || '—',
        nd_desc: cl.nd_desc || 'Natureza de Despesa',
        pi: cl.pi || '—',
        pi_desc: cl.pi_nome || 'Plano Interno',
        prov: totProv,
        bloq: 0,
        emp: totEmp,
        liq: totLiq,
        pag: totPag,
        cred: totCred,
        obj: cl.obj || '',
        itens: subItens.length > 1 ? subItens : null
      };
    }
  }
  if(!item && typeof NCDATA !== 'undefined' && NCDATA && NCDATA[hid]){
    var ndo = NCDATA[hid];
    item = {
      hid: ndo.nc,
      nc: ndo.nc,
      op: ndo.op || 'DESCENTRALIZACAO DE CREDITO',
      dia: ndo.dia || '—',
      dias: null,
      status: (ndo.val > 0 ? 'Disponível' : 'Consumido'),
      status_slug: (ndo.val > 0 ? 'ok' : 'danger'),
      emit_cod: '160073',
      emit_nome: 'DIRETORIA DE GESTAO ORCAMENTARIA - GESTOR',
      fav_cod: '160353',
      fav_nome: '6º BEC · OGU',
      om_sigla: '6º BEC',
      ptres: ndo.acao || '—',
      acao_desc: 'Ação ' + (ndo.acao || '—'),
      fonte: ndo.u || 'OGU',
      nd: ndo.nd || '—',
      nd_desc: ndo.ndn || 'Natureza de Despesa',
      pi: ndo.pi || '—',
      pi_desc: 'Plano Interno',
      prov: ndo.val || 0,
      bloq: 0,
      emp: 0,
      liq: 0,
      pag: 0,
      cred: ndo.val || 0,
      obj: ndo.obj || ''
    };
  }
  if(!item) return;

  var prov = item.prov || 0;
  var bloq = item.bloq || 0;
  var emp = item.emp || 0;
  var liq = item.liq || 0;
  var pag = item.pag || 0;
  var cred = item.cred || 0;

  var pEmp = prov > 0 ? (emp / prov * 100) : 0;
  var pCred = prov > 0 ? (cred / prov * 100) : 0;
  var pLiq = emp > 0 ? (liq / emp * 100) : 0;
  var pPag = liq > 0 ? (pag / liq * 100) : 0;

  var isCatr = (item.ptres === '21EM') || (item.obj && item.obj.indexOf('CATRIMANI') !== -1) || (item.acao_desc && item.acao_desc.indexOf('21EM') !== -1);
  var badgeOp = isCatr ? '<span class="m-badge-op">🛡️ OPERAÇÃO CATRIMANI II · AÇÃO 21EM</span>' : '<span class="m-badge-op" style="color:#60A5FA;background:rgba(59,130,246,0.12);border-color:rgba(59,130,246,0.3);">🏛️ 6º BATALHÃO DE ENGENHARIA DE CONSTRUÇÃO</span>';
  var barGrad = isCatr ? 'linear-gradient(90deg, #10B981 0%, #059669 35%, #F59E0B 75%, #EAB308 100%)' : 'linear-gradient(90deg, #10B981 0%, #059669 35%, #2563EB 75%, #8B5CF6 100%)';
  var statusBadge = cred > 0 ? '<span class="m-badge-status-lg status-ok">● DISPONÍVEL INTEGRAL</span>' : '<span class="m-badge-status-lg status-warn">● EXECUTADO INTEGRAL</span>';

  var fonteExtenso = item.fonte === 'OGU' ? '160 - Orçamento Geral da União (OGU)' : (item.fonte === 'FEx' ? '167 - Fundo do Exército (FEx)' : item.fonte);

  var h = '';
  h += '<div class="m-accent-bar" style="background:' + barGrad + ';"></div>';
  h += '<div class="m-content-wrap">';

  /* Cabeçalho Executivo */
  h += '  <div class="m-header-v2">';
  h += '    <div class="m-header-meta-row">';
  h += '      ' + badgeOp;
  h += '      ' + statusBadge;
  h += '    </div>';
  h += '    <div class="m-title-row">';
  h += '      <div class="m-nc-code-block">';
  h += '        <h3 id="modal-title">Nota de Crédito ' + bcmsEsc(item.nc) + '</h3>';
  h += '      </div>';
  h += '      <button type="button" class="m-btn-pill" onclick="bcmsCopiarNC(\'' + bcmsEsc(item.nc) + '\', this)" title="Copiar número da NC">';
  h += '        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg> Copiar NC';
  h += '      </button>';
  h += '    </div>';
  h += '    <div class="m-sub-meta">';
  h += '      <span class="m-meta-chip">⚙️ ' + bcmsEsc(item.op || 'DESCENTRALIZACAO DE CREDITO') + '</span>';
  h += '      <span class="m-meta-chip">📅 Emissão: <b>' + bcmsEsc(item.dia || '—') + '</b>' + (item.dias !== null && item.dias !== undefined ? ' (' + item.dias + ' dias)' : '') + '</span>';
  h += '      <span class="m-meta-chip">🏛️ Exercício Financeiro 2026</span>';
  h += '    </div>';
  h += '  </div>';

  /* Seção 1: Balanço Financeiro */
  h += '  <div class="m-fin-section">';
  h += '    <div class="m-fin-grid">';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Valor Original (Recebido)</span>';
  h += '        <span class="m-fin-val col-prov">' + bcmsBRL(prov) + '</span>';
  h += '        <span class="m-fin-sub">Dotação integral descentralizada</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Bloqueado / Pré-Empenho</span>';
  h += '        <span class="m-fin-val">' + bcmsBRL(bloq) + '</span>';
  h += '        <span class="m-fin-sub">Reserva de crédito orçamentário</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Executado (Empenhado)</span>';
  h += '        <span class="m-fin-val col-emp">' + bcmsBRL(emp) + '</span>';
  h += '        <span class="m-fin-sub">' + pEmp.toFixed(1) + '% consumido da dotação</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card hero-saldo">';
  h += '        <span class="m-fin-label" style="color:#10B981;">Saldo Disponível em Tela</span>';
  h += '        <span class="m-fin-val-hero">' + bcmsBRL(cred) + '</span>';
  h += '        <span class="m-fin-tag-hero">✓ ' + pCred.toFixed(1) + '% remanescente líquido</span>';
  h += '      </div>';
  h += '    </div>';

  /* Pipeline 3 Estágios */
  h += '    <div class="m-pipeline-card">';
  h += '      <div class="m-pipeline-header">Pipeline de Execução Orçamentária e Financeira (SIAFI)</div>';
  h += '      <div class="m-stages-grid">';
  h += '        <div class="m-stage-item">';
  h += '          <div class="m-stage-header"><span>1. Empenho / Recebido</span><b>' + bcmsBRL(emp) + ' (' + pEmp.toFixed(1) + '%)</b></div>';
  h += '          <div class="m-progress-track"><div class="m-progress-bar bar-emp" style="width:' + Math.min(100, pEmp).toFixed(1) + '%"></div></div>';
  h += '          <span class="m-stage-sub">Comprometimento formal da dotação</span>';
  h += '        </div>';
  h += '        <div class="m-stage-item">';
  h += '          <div class="m-stage-header"><span>2. Liquidação / Empenho</span><b>' + bcmsBRL(liq) + ' (' + pLiq.toFixed(1) + '%)</b></div>';
  h += '          <div class="m-progress-track"><div class="m-progress-bar bar-liq" style="width:' + Math.min(100, pLiq).toFixed(1) + '%"></div></div>';
  h += '          <span class="m-stage-sub">Atesto e liquidação de serviços ou bens</span>';
  h += '        </div>';
  h += '        <div class="m-stage-item">';
  h += '          <div class="m-stage-header"><span>3. Pagamento / Liquidação</span><b>' + bcmsBRL(pag) + ' (' + pPag.toFixed(1) + '%)</b></div>';
  h += '          <div class="m-progress-track"><div class="m-progress-bar bar-pag" style="width:' + Math.min(100, pPag).toFixed(1) + '%"></div></div>';
  h += '          <span class="m-stage-sub">Ordem bancária efetivada no Tesouro</span>';
  h += '        </div>';
  h += '      </div>';
  h += '    </div>';
  h += '  </div>';

  /* Seção 2: Tramitação de UGs */
  h += '  <div class="m-route-grid">';
  h += '    <div class="m-route-card">';
  h += '      <span class="m-route-badge">UG Emitente (Origem)</span>';
  h += '      <div class="m-route-title"><span class="ug-pill emit">' + bcmsEsc(item.emit_cod || '160073') + '</span> ' + bcmsEsc(item.emit_nome || 'DGO') + '</div>';
  h += '      <span class="m-route-desc">Órgão Central Setorial / Diretoria de Gestão Orçamentária</span>';
  h += '    </div>';
  h += '    <div class="m-route-arrow">➔</div>';
  h += '    <div class="m-route-card">';
  h += '      <span class="m-route-badge">UG Favorecida (Executora)</span>';
  h += '      <div class="m-route-title"><span class="ug-pill fav">' + bcmsEsc(item.fav_cod || '160353') + '</span> ' + bcmsEsc(item.fav_nome || '6º BEC') + (item.om_sigla ? ' (' + bcmsEsc(item.om_sigla) + ')' : '') + '</div>';
  h += '      <span class="m-route-desc">Unidade Gestora Executora Operacional</span>';
  h += '    </div>';
  h += '  </div>';

  /* Seção 3: Classificação Orçamentária */
  h += '  <div class="m-class-grid">';
  h += '    <div class="m-class-card">';
  h += '      <span class="m-class-label">PTRES / Ação Orçamentária</span>';
  h += '      <span class="m-class-code">' + bcmsEsc(item.ptres || '—') + '</span>';
  h += '      <span class="m-class-desc">' + bcmsEsc(item.acao_desc || 'Ação Governamental') + '</span>';
  h += '    </div>';
  h += '    <div class="m-class-card">';
  h += '      <span class="m-class-label">Fonte de Recursos</span>';
  h += '      <span class="m-class-code"><span class="pill-fonte">' + bcmsEsc(item.fonte || 'OGU') + '</span></span>';
  h += '      <span class="m-class-desc">' + bcmsEsc(fonteExtenso) + '</span>';
  h += '    </div>';
  h += '    <div class="m-class-card">';
  h += '      <span class="m-class-label">Natureza de Despesa (ND)</span>';
  h += '      <span class="m-class-code">' + bcmsEsc(item.nd || '—') + '</span>';
  h += '      <span class="m-class-desc">' + bcmsEsc(item.nd_desc || 'Elemento de Despesa') + '</span>';
  h += '    </div>';
  h += '    <div class="m-class-card">';
  h += '      <span class="m-class-label">Plano Interno (PI)</span>';
  h += '      <span class="m-class-code">' + bcmsEsc(item.pi || '—') + '</span>';
  h += '      <span class="m-class-desc">' + bcmsEsc(item.pi_desc || 'Plano Interno') + '</span>';
  h += '    </div>';
  h += '  </div>';

  /* Seção 4: Finalidade / Observação */
  var rawObj = item.obj || '(Sem texto de observação cadastrado no SIAFI)';
  var escObjAttr = bcmsEsc(rawObj).replace(/"/g, '&quot;');
  h += '  <div class="m-justif-card">';
  h += '    <div class="m-justif-header">';
  h += '      <span class="m-justif-title">Finalidade / Objeto Cadastrado no SIAFI</span>';
  h += '      <button type="button" class="m-btn-pill" style="padding:4px 10px;font-size:0.75rem;" onclick="bcmsCopiarTexto(this)" data-copy="' + escObjAttr + '">📋 Copiar Texto</button>';
  h += '    </div>';
  h += '    <div class="m-justif-body">' + bcmsEsc(rawObj) + '</div>';
  h += '  </div>';

  /* Subtabela de desdobramentos se houver mais de 1 lançamento */
  if(item.itens && item.itens.length > 1){
    h += '  <div class="m-justif-card" style="border-left-color:var(--primary-600);">';
    h += '    <div class="m-justif-header"><span class="m-justif-title">Desdobramentos da Nota (' + item.itens.length + ' lançamentos)</span></div>';
    h += '    <div class="tbl-scroll"><table class="det det-itens"><thead><tr><th>Ação</th><th>PI</th><th>ND</th><th class="num">Valor</th></tr></thead><tbody>';
    for(var j = 0; j < item.itens.length; j++){
      var sub = item.itens[j];
      h += '<tr><td class="mono2">' + bcmsEsc(sub.acao || '—') + '</td><td class="mono2">' + bcmsEsc(sub.pi || '—') + '</td><td class="mono2">' + bcmsEsc(sub.nd || '—') + '</td><td class="num">' + bcmsBRL(sub.val || 0) + '</td></tr>';
    }
    h += '</tbody></table></div>';
    h += '  </div>';
  }

  /* Rodapé de Ações Executivo */
  h += '  <div class="m-footer-actions-v2">';
  h += '    <button type="button" class="m-btn-pill" onclick="bcmsVerNoHistorico(\'' + bcmsEsc(item.nc) + '\')">🔍 Localizar no Histórico Geral</button>';
  h += '    <button type="button" class="m-btn-pill" onclick="bcmsCopiarNC(\'' + bcmsEsc(item.nc) + '\', this)">📋 Copiar Nº da NC</button>';
  h += '    <button type="button" class="m-btn-pill primary" onclick="bcmsCelClose()">Fechar Janela ✕</button>';
  h += '  </div>';

  h += '</div>';

  document.getElementById('modal-body').innerHTML = h;
  var m = document.getElementById('modal');
  m.classList.add('open');
  m.setAttribute('aria-hidden', 'false');
  var x = document.querySelector('.modal-x');
  if(x) x.focus();
}

function bcmsCopiarNC(ncNumero, btn){
  if(navigator.clipboard && navigator.clipboard.writeText){
    navigator.clipboard.writeText(ncNumero).then(function(){
      bcmsToast('📋 Nota de Crédito ' + ncNumero + ' copiada com sucesso!');
      if(btn){
        var old = btn.innerHTML;
        btn.innerHTML = '✓ Copiado!';
        setTimeout(function(){ btn.innerHTML = old; }, 1800);
      }
    }).catch(function(){
      bcmsToast('NC: ' + ncNumero);
    });
  } else {
    var ta = document.createElement('textarea');
    ta.value = ncNumero;
    document.body.appendChild(ta);
    ta.select();
    try { 
      document.execCommand('copy'); 
      bcmsToast('📋 Nota de Crédito ' + ncNumero + ' copiada!');
      if(btn){
        var old = btn.innerHTML;
        btn.innerHTML = '✓ Copiado!';
        setTimeout(function(){ btn.innerHTML = old; }, 1800);
      }
    } catch(e){}
    document.body.removeChild(ta);
  }
}

function bcmsCopiarTexto(btn){
  var txt = btn.getAttribute('data-copy') || '';
  if(!txt) return;
  if(navigator.clipboard && navigator.clipboard.writeText){
    navigator.clipboard.writeText(txt).then(function(){
      bcmsToast('📋 Finalidade copiada para a área de transferência!');
      var old = btn.innerHTML;
      btn.innerHTML = '✓ Copiado!';
      setTimeout(function(){ btn.innerHTML = old; }, 1800);
    }).catch(function(){
      bcmsToast('Texto copiado');
    });
  } else {
    var ta = document.createElement('textarea');
    ta.value = txt;
    document.body.appendChild(ta);
    ta.select();
    try { 
      document.execCommand('copy'); 
      bcmsToast('📋 Finalidade copiada!');
      var old = btn.innerHTML;
      btn.innerHTML = '✓ Copiado!';
      setTimeout(function(){ btn.innerHTML = old; }, 1800);
    } catch(e){}
    document.body.removeChild(ta);
  }
}

function bcmsVerNoHistorico(ncNumero){
  bcmsCelClose();
  trocaOMDSPorKey('HISTORICO');
  var inp = document.getElementById('flt-hist-busca');
  if(inp){
    inp.value = ncNumero;
    bcmsFiltraHistorico();
  }
}

function bcmsExportHistoricoExcel(){
  var list = (HIST_FILTERED && HIST_FILTERED.length) ? HIST_FILTERED : ((HISTDATA && HISTDATA.items) ? HISTDATA.items : []);
  if(!list.length){
    bcmsToast('⚠️ Não há registros no histórico para exportar.');
    return;
  }
  var filename = 'historico_notas_credito_2026_' + (new Date().toISOString().slice(0, 10));
  var headers = ['Data', 'Número da NC', 'UG Emitente (Cód)', 'UG Emitente (Nome)', 'UG Favorecida (Cód)', 'UG Favorecida (OM)', 'PTRES / Ação', 'Fonte', 'PI', 'ND', 'Valor Original (R$)', 'Valor Executado (R$)', 'Saldo Atual (R$)', 'Status', 'Justificativa / Objeto'];

  var xml = '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<?mso-application progid="Excel.Sheet"?>\n' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n' +
    ' xmlns:o="urn:schemas-microsoft-com:office:office"\n' +
    ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n' +
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n' +
    '<Styles>\n' +
    ' <Style ss:ID="Default" ss:Name="Normal"><Font ss:FontName="Calibri" ss:Size="11" ss:Color="#000000"/><Alignment ss:Vertical="Center"/></Style>\n' +
    ' <Style ss:ID="Header"><Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1C4A73" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>\n' +
    ' <Style ss:ID="Title"><Font ss:FontName="Calibri" ss:Size="14" ss:Bold="1" ss:Color="#1C4A73"/><Alignment ss:Vertical="Center"/></Style>\n' +
    ' <Style ss:ID="Currency"><NumberFormat ss:Format="&quot;R$&quot; #,##0.00"/><Alignment ss:Horizontal="Right" ss:Vertical="Center"/></Style>\n' +
    ' <Style ss:ID="Center"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>\n' +
    '</Styles>\n' +
    '<Worksheet ss:Name="Histórico de NCs">\n' +
    '<Table ss:DefaultRowHeight="20">\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="80"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="110"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="90"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="200"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="90"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="110"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="80"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="70"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="110"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="90"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="130"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="130"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="130"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="120"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="380"/>\n';

  xml += ' <Row ss:Height="26"><Cell ss:StyleID="Title" ss:MergeAcross="' + (headers.length - 1) + '"><Data ss:Type="String">6º BATALHÃO DE ENGENHARIA DE CONSTRUÇÃO — HISTÓRICO DE NOTAS DE CRÉDITO</Data></Cell></Row>\n';
  xml += ' <Row ss:Height="18"><Cell ss:MergeAcross="' + (headers.length - 1) + '"><Data ss:Type="String">Exercício Financeiro Corrente · ' + list.length + ' registro(s) exportado(s) · Posição SIAFI / Tesouro Gerencial · ' + new Date().toLocaleDateString('pt-BR') + '</Data></Cell></Row>\n';
  xml += ' <Row ss:Height="10"/>\n';
  xml += ' <Row ss:Height="24">\n';
  headers.forEach(function(h){ xml += '  <Cell ss:StyleID="Header"><Data ss:Type="String">' + bcmsEscXml(h) + '</Data></Cell>\n'; });
  xml += ' </Row>\n';

  list.forEach(function(item){
    xml += ' <Row ss:Height="20">\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.dia || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.nc || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.emit_cod || '') + '</Data></Cell>\n';
    xml += '  <Cell><Data ss:Type="String">' + bcmsEscXml(item.emit_nome || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.fav_cod || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.om_sigla || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.ptres || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.fonte || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.pi || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.nd || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Currency"><Data ss:Type="Number">' + (item.prov || 0).toFixed(2) + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Currency"><Data ss:Type="Number">' + (item.emp || 0).toFixed(2) + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Currency"><Data ss:Type="Number">' + (item.cred || 0).toFixed(2) + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.status || '') + '</Data></Cell>\n';
    xml += '  <Cell><Data ss:Type="String">' + bcmsEscXml(item.obj || '') + '</Data></Cell>\n';
    xml += ' </Row>\n';
  });

  xml += '</Table></Worksheet></Workbook>';
  var blob = new Blob([xml], {type: 'application/vnd.ms-excel;charset=utf-8;'});
  var link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename + '.xls';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  bcmsToast('📊 Planilha do Histórico exportada com sucesso (' + list.length + ' linhas)!');
}

/* ==========================================================================
   HISTÓRICO COMPLETO POR UNIDADE (Subaba Analítica da OMDS)
   ========================================================================== */
var UHIST_STATE = {};

function bcmsGetUHistState(sfx){
  if(!UHIST_STATE[sfx]){
    var u = (typeof UNIDADES !== 'undefined' && UNIDADES[sfx]) ? UNIDADES[sfx] : null;
    var allItems = [];
    if(typeof HISTDATA !== 'undefined' && HISTDATA && HISTDATA.items && u){
      allItems = HISTDATA.items.filter(function(it){
        return it.fav_cod === u.ogu || it.fav_cod === u.fex;
      });
    }
    UHIST_STATE[sfx] = {
      items: allItems,
      filtered: allItems.slice(),
      page: 1,
      pageSize: 50,
      sortCol: 'dt',
      sortDir: 'desc'
    };
  }
  return UHIST_STATE[sfx];
}

function bcmsInitUHist(sfx){
  var st = bcmsGetUHistState(sfx);
  bcmsFiltraUHist(sfx);
}

function bcmsFiltraUHist(sfx){
  var st = bcmsGetUHistState(sfx);
  var raw = st.items;
  if(!raw) return;

  var v = function(id){ var e = document.getElementById(id + '-' + sfx); return e ? e.value.trim() : ''; };
  var fPer = v('flt-uhist-periodo');
  var fFonte = v('flt-uhist-fonte');
  var fPtres = v('flt-uhist-ptres');
  var fFaixa = v('flt-uhist-faixa');
  var q = (v('flt-uhist-busca') || '').toLowerCase();

  st.filtered = raw.filter(function(it){
    if(fPer){
      if(fPer.startsWith('T')){
        if(it.tri !== fPer) return false;
      } else {
        if(it.mes !== fPer) return false;
      }
    }
    if(fFonte && it.fonte !== fFonte) return false;
    if(fPtres && it.ptres !== fPtres && it.acao !== fPtres) return false;
    if(fFaixa){
      if(fFaixa === 'saldo_pos' && it.cred <= 0.01) return false;
      if(fFaixa === 'parcial' && (it.status_slug !== 'parcial')) return false;
      if(fFaixa === 'zerada' && (it.cred > 0.01 || it.status_slug === 'canc')) return false;
      if(fFaixa === 'canc' && it.status_slug !== 'canc') return false;
    }
    if(q){
      var match = (it.nc && it.nc.toLowerCase().indexOf(q) > -1) ||
                  (it.obj && it.obj.toLowerCase().indexOf(q) > -1) ||
                  (it.pi && it.pi.toLowerCase().indexOf(q) > -1) ||
                  (it.pi_desc && it.pi_desc.toLowerCase().indexOf(q) > -1) ||
                  (it.emit_nome && it.emit_nome.toLowerCase().indexOf(q) > -1) ||
                  (it.emit_cod && String(it.emit_cod).indexOf(q) > -1) ||
                  (it.ptres && it.ptres.toLowerCase().indexOf(q) > -1) ||
                  (it.nd && it.nd.toLowerCase().indexOf(q) > -1);
      if(!match) return false;
    }
    return true;
  });

  bcmsSortUHistFiltered(sfx);
  bcmsAtualizaUHistKPIs(sfx);
  bcmsRenderUHist(sfx, 1);
}

function bcmsSortUHistFiltered(sfx){
  var st = bcmsGetUHistState(sfx);
  var col = st.sortCol;
  var dir = st.sortDir;
  st.filtered.sort(function(a, b){
    var va = a[col], vb = b[col];
    if(col === 'dt'){
      va = a.dt || '';
      vb = b.dt || '';
    } else if(col === 'prov' || col === 'cred' || col === 'emp'){
      va = a[col] || 0;
      vb = b[col] || 0;
    } else if(col === 'emit'){
      va = a.emit_nome || a.emit_cod || '';
      vb = b.emit_nome || b.emit_cod || '';
    } else if(col === 'ptres'){
      va = (a.ptres || '') + ' ' + (a.nd || '');
      vb = (b.ptres || '') + ' ' + (b.nd || '');
    } else if(col === 'status'){
      va = a.status || '';
      vb = b.status || '';
    } else {
      va = String(a[col] || '').toLowerCase();
      vb = String(b[col] || '').toLowerCase();
    }
    if(va < vb) return dir === 'asc' ? -1 : 1;
    if(va > vb) return dir === 'asc' ? 1 : -1;
    return 0;
  });
}

function bcmsSortUHist(col, th, sfx){
  var st = bcmsGetUHistState(sfx);
  if(st.sortCol === col){
    st.sortDir = (st.sortDir === 'asc') ? 'desc' : 'asc';
  } else {
    st.sortCol = col;
    st.sortDir = (col === 'dt' || col === 'prov' || col === 'cred') ? 'desc' : 'asc';
  }
  var thead = th.closest('thead');
  if(thead){
    thead.querySelectorAll('th').forEach(function(h){
      h.setAttribute('aria-sort', 'none');
      var s = h.querySelector('.sort');
      if(s) s.textContent = '';
    });
    th.setAttribute('aria-sort', st.sortDir === 'asc' ? 'ascending' : 'descending');
    var s = th.querySelector('.sort');
    if(s) s.textContent = st.sortDir === 'asc' ? ' ▲' : ' ▼';
  }
  bcmsSortUHistFiltered(sfx);
  bcmsRenderUHist(sfx, st.page);
}

function bcmsAtualizaUHistKPIs(sfx){
  var st = bcmsGetUHistState(sfx);
  var list = st.filtered;
  var total = list.length;
  var distinctSet = {};
  var totProv = 0, totEmp = 0, totCred = 0;
  for(var i = 0; i < list.length; i++){
    var it = list[i];
    distinctSet[it.nc] = true;
    totProv += it.prov;
    totEmp += it.emp;
    totCred += it.cred;
  }
  var distinctCount = Object.keys(distinctSet).length;

  var k1 = document.getElementById('kpi-uhist-total-' + sfx);
  if(k1) k1.textContent = distinctCount.toLocaleString('pt-BR');

  var k2 = document.getElementById('kpi-uhist-prov-' + sfx);
  if(k2) k2.textContent = bcmsBRL(totProv);

  var k3 = document.getElementById('kpi-uhist-emp-' + sfx);
  if(k3) k3.textContent = bcmsBRL(totEmp);

  var k4 = document.getElementById('kpi-uhist-cred-' + sfx);
  if(k4) k4.textContent = bcmsBRL(totCred);

  var badge = document.getElementById('cnt-uhist-ncs-' + sfx);
  if(badge){
    var totalUnit = st.items.length;
    var ativo = (total !== totalUnit);
    badge.textContent = (ativo ? 'Filtrado: ' : 'Exibindo ') + total + ' de ' + totalUnit + ' NC(s) no histórico (' + distinctCount + ' distintas)';
  }

  /* tfoot */
  var tfLabel = document.getElementById('tf-uhist-label-' + sfx);
  if(tfLabel) tfLabel.textContent = (total !== st.items.length ? 'FILTRADO · ' : 'TOTAL · ') + total + ' Nota(s) de Crédito no histórico';
  var tfProv = document.getElementById('tf-uhist-prov-' + sfx);
  if(tfProv) tfProv.textContent = bcmsBRL(totProv);
  var tfCred = document.getElementById('tf-uhist-cred-' + sfx);
  if(tfCred) tfCred.textContent = bcmsBRL(totCred);

  /* chip de resumo dos filtros */
  var v = function(id){ var e = document.getElementById(id + '-' + sfx); return e ? e.value.trim() : ''; };
  var fPer = v('flt-uhist-periodo');
  var fFonte = v('flt-uhist-fonte');
  var fPtres = v('flt-uhist-ptres');
  var fFaixa = v('flt-uhist-faixa');
  var q = v('flt-uhist-busca');

  var res = document.getElementById('flt-uhist-res-' + sfx);
  if(res){
    var pk = [];
    if(fFonte) pk.push('Fonte ' + fFonte);
    if(fPtres) pk.push('Ação ' + fPtres);
    if(fPer) pk.push(fPer);
    if(fFaixa) pk.push({saldo_pos:'Com Saldo',parcial:'Parcial',zerada:'Zerada',canc:'Cancelada'}[fFaixa] || fFaixa);
    if(q) pk.push('“' + q + '”');
    res.textContent = pk.length ? (pk.join(' · ') + ' — ' + bcmsBRL(totCred)) : '';
    res.className = 'flt-resumo' + (pk.length ? ' on' : '');
  }
}

function bcmsLimpaFiltrosUHist(sfx){
  ['flt-uhist-periodo','flt-uhist-fonte','flt-uhist-ptres','flt-uhist-faixa'].forEach(function(baseId){
    var el = document.getElementById(baseId + '-' + sfx);
    if(el) el.value = '';
  });
  var b = document.getElementById('flt-uhist-busca-' + sfx);
  if(b) b.value = '';
  bcmsFiltraUHist(sfx);
  bcmsToast('Filtros do histórico de ' + sfx + ' redefinidos.');
}

function bcmsRenderUHist(sfx, page){
  var st = bcmsGetUHistState(sfx);
  st.page = page;
  var total = st.filtered.length;
  var totalPages = Math.max(1, Math.ceil(total / st.pageSize));
  if(st.page > totalPages) st.page = totalPages;
  if(st.page < 1) st.page = 1;

  var start = (st.page - 1) * st.pageSize;
  var end = Math.min(start + st.pageSize, total);
  var pageItems = st.filtered.slice(start, end);

  var tbody = document.getElementById('tbody-uhist-ncs-' + sfx);
  if(!tbody) return;

  if(pageItems.length === 0){
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:48px 16px;color:var(--ink-muted);"><span style="font-size:2rem;display:block;margin-bottom:8px;">🔍</span>Nenhuma Nota de Crédito encontrada para os filtros selecionados.<br><button type="button" class="flt-limpa" style="margin-top:12px;" onclick="bcmsLimpaFiltrosUHist(\'' + sfx + '\')">✕ Limpar Filtros</button></td></tr>';
  } else {
    var rowsHtml = '';
    for(var i = 0; i < pageItems.length; i++){
      var item = pageItems[i];
      var statusCls = 'status-' + item.status_slug;
      var emitNomeCurto = (item.emit_nome && item.emit_nome.length > 22) ? item.emit_nome.substring(0, 22) + '…' : (item.emit_nome || '—');
      var descCompleta = item.obj || '';
      var descResumo = descCompleta.length > 118 ? (descCompleta.substring(0, 118) + '…') : descCompleta;
      var ncFull = String(item.nc || '');
      var mNc = ncFull.match(/NC(\d+)$/);
      var ncLbl = mNc ? ('<span class="nc-num">NC ' + mNc[1] + '</span> <span class="nc-ug">· ' + bcmsEsc(ncFull.substring(0,6)) + '</span>') : ('<span class="nc-num">' + bcmsEsc(ncFull) + '</span>');
      var acaoNd = (item.ptres && item.nd) ? (item.ptres + ' · ' + item.nd) : (item.ptres || item.nd || '—');

      rowsHtml += '<tr class="cel-row" data-hid="' + bcmsEsc(item.hid) + '" tabindex="0" role="button" onclick="bcmsDetalheNC(\'' + bcmsEsc(item.hid) + '\')" '
        + 'title="Clique para abrir o detalhamento completo da NC ' + bcmsEsc(ncFull) + '" '
        + 'onkeydown="if(event.key===\'Enter\'||event.key===\' \'){event.preventDefault();bcmsDetalheNC(\'' + bcmsEsc(item.hid) + '\')}">'
        + '<td><span class="pill-fonte">' + bcmsEsc(item.fonte) + '</span></td>'
        + '<td class="mono2" title="' + bcmsEsc(ncFull) + '">' + ncLbl + '</td>'
        + '<td class="mono2">' + bcmsEsc(acaoNd) + '</td>'
        + '<td class="obj" title="' + bcmsEsc(descCompleta) + '" data-full-desc="' + bcmsEsc(descCompleta) + '">' + bcmsEsc(descResumo) + '</td>'
        + '<td title="' + bcmsEsc(item.emit_nome) + '"><span class="ug-pill emit">' + bcmsEsc(item.emit_cod) + '</span> <small>' + bcmsEsc(emitNomeCurto) + '</small></td>'
        + '<td class="mono2">' + bcmsEsc(item.dia || '—') + '</td>'
        + '<td class="num" data-sort="' + item.prov.toFixed(2) + '">' + bcmsBRL(item.prov) + '</td>'
        + '<td class="num anchor" data-sort="' + item.cred.toFixed(2) + '">' + bcmsBRL(item.cred) + '</td>'
        + '<td class="num"><span class="pill-status ' + statusCls + '">' + bcmsEsc(item.status) + '</span><i class="chev" aria-hidden="true">›</i></td>'
        + '</tr>';
    }
    tbody.innerHTML = rowsHtml;
  }

  var pagInfo = document.getElementById('pag-info-uhist-' + sfx);
  if(pagInfo){
    if(total === 0){
      pagInfo.textContent = 'Página 0 de 0 (0 registros)';
    } else {
      pagInfo.textContent = 'Página ' + st.page + ' de ' + totalPages + ' (Exibindo ' + (start + 1) + '–' + end + ' de ' + total + ' NCs)';
    }
  }

  var btnAnt = document.getElementById('btn-pag-uhist-ant-' + sfx);
  if(btnAnt) btnAnt.disabled = (st.page <= 1);
  var btnProx = document.getElementById('btn-pag-uhist-prox-' + sfx);
  if(btnProx) btnProx.disabled = (st.page >= totalPages);
}

function bcmsPaginaUHist(delta, sfx){
  var st = bcmsGetUHistState(sfx);
  bcmsRenderUHist(sfx, st.page + delta);
  var scrollCont = document.getElementById('scroll-uhist-ncs-' + sfx);
  if(scrollCont) scrollCont.scrollIntoView({behavior:'smooth', block:'nearest'});
}

function bcmsExportUHistExcel(sfx, sigla){
  var st = bcmsGetUHistState(sfx);
  var list = st.filtered;
  if(!list || !list.length){
    bcmsToast('⚠️ Não há registros para exportar.');
    return;
  }
  var omNome = sigla || sfx;
  var filename = 'historico_ncs_' + omNome.toLowerCase().replace(/[\s·]+/g, '_') + '_' + (new Date().toISOString().slice(0, 10));
  var headers = ['Data', 'Número da NC', 'UG Emitente (Cód)', 'UG Emitente (Nome)', 'UG Favorecida (Cód)', 'UG Favorecida (OM)', 'PTRES / Ação', 'Fonte', 'PI', 'ND', 'Valor Original (R$)', 'Valor Executado (R$)', 'Saldo Atual (R$)', 'Status', 'Justificativa / Objeto'];

  var xml = '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<?mso-application progid="Excel.Sheet"?>\n' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n' +
    ' xmlns:o="urn:schemas-microsoft-com:office:office"\n' +
    ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n' +
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n' +
    '<Styles>\n' +
    ' <Style ss:ID="Default" ss:Name="Normal"><Font ss:FontName="Calibri" ss:Size="11" ss:Color="#000000"/><Alignment ss:Vertical="Center"/></Style>\n' +
    ' <Style ss:ID="Header"><Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1C4A73" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>\n' +
    ' <Style ss:ID="Title"><Font ss:FontName="Calibri" ss:Size="14" ss:Bold="1" ss:Color="#1C4A73"/><Alignment ss:Vertical="Center"/></Style>\n' +
    ' <Style ss:ID="Currency"><NumberFormat ss:Format="&quot;R$&quot; #,##0.00"/><Alignment ss:Horizontal="Right" ss:Vertical="Center"/></Style>\n' +
    ' <Style ss:ID="Center"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>\n' +
    '</Styles>\n' +
    '<Worksheet ss:Name="Histórico ' + bcmsEscXml(omNome) + '">\n' +
    '<Table ss:DefaultRowHeight="20">\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="80"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="110"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="90"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="200"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="90"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="110"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="80"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="70"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="110"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="90"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="130"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="130"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="130"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="120"/>\n' +
    ' <Column ss:AutoFitWidth="1" ss:Width="380"/>\n';

  xml += ' <Row ss:Height="26"><Cell ss:StyleID="Title" ss:MergeAcross="' + (headers.length - 1) + '"><Data ss:Type="String">6º BATALHÃO DE ENGENHARIA DE CONSTRUÇÃO — NOTAS DE CRÉDITO (' + bcmsEscXml(omNome) + ')</Data></Cell></Row>\n';
  xml += ' <Row ss:Height="18"><Cell ss:MergeAcross="' + (headers.length - 1) + '"><Data ss:Type="String">Exercício Financeiro Corrente · ' + list.length + ' registro(s) exportado(s) · Posição SIAFI / Tesouro Gerencial · ' + new Date().toLocaleDateString('pt-BR') + '</Data></Cell></Row>\n';
  xml += ' <Row ss:Height="10"/>\n';
  xml += ' <Row ss:Height="24">\n';
  headers.forEach(function(h){ xml += '  <Cell ss:StyleID="Header"><Data ss:Type="String">' + bcmsEscXml(h) + '</Data></Cell>\n'; });
  xml += ' </Row>\n';

  list.forEach(function(item){
    xml += ' <Row ss:Height="20">\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.dia || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.nc || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.emit_cod || '') + '</Data></Cell>\n';
    xml += '  <Cell><Data ss:Type="String">' + bcmsEscXml(item.emit_nome || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.fav_cod || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.om_sigla || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.ptres || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.fonte || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.pi || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.nd || '') + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Currency"><Data ss:Type="Number">' + (item.prov || 0).toFixed(2) + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Currency"><Data ss:Type="Number">' + (item.emp || 0).toFixed(2) + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Currency"><Data ss:Type="Number">' + (item.cred || 0).toFixed(2) + '</Data></Cell>\n';
    xml += '  <Cell ss:StyleID="Center"><Data ss:Type="String">' + bcmsEscXml(item.status || '') + '</Data></Cell>\n';
    xml += '  <Cell><Data ss:Type="String">' + bcmsEscXml(item.obj || '') + '</Data></Cell>\n';
    xml += ' </Row>\n';
  });

  xml += '</Table></Worksheet></Workbook>';
  var blob = new Blob([xml], {type: 'application/vnd.ms-excel;charset=utf-8;'});
  var link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename + '.xls';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  bcmsToast('📊 Planilha do Histórico de ' + omNome + ' exportada com sucesso (' + list.length + ' linhas)!');
}



/* ==========================================================================
   MELHORIAS DE DETALHAMENTO & AUDITORIA MULTI-UG / OMDS / KPIS
   ========================================================================== */

function bcmsDetalheOMDS(key){
  if(!key) return;
  if(key === 'BEC6'){
    trocaOMDSPorKey('BEC6');
    return;
  }
  var u = (typeof OMDSDATA !== 'undefined' && OMDSDATA) ? OMDSDATA[key] : null;
  if(!u){
    if(typeof CATRDATA !== 'undefined' && CATRDATA && CATRDATA.por_ug){
      for(var i = 0; i < CATRDATA.por_ug.length; i++){
        if(CATRDATA.por_ug[i].cod === key || CATRDATA.por_ug[i].nome.indexOf(key) !== -1){
          bcmsDetalheUG(CATRDATA.por_ug[i].cod);
          return;
        }
      }
    }
    return;
  }

  var prov = u.prov || 0;
  var emp  = u.emp || 0;
  var cred = u.cred || 0;
  var liq  = u.liq || 0;
  var pag  = u.pag || 0;
  var pctEmp = prov > 0 ? (emp / prov * 100) : 0;
  var pctLiq = emp > 0 ? (liq / emp * 100) : 0;
  var pctPag = liq > 0 ? (pag / liq * 100) : 0;

  var semCor = pctEmp >= 90 ? 'var(--ok, #10B981)' : (pctEmp >= 80 ? 'var(--gold, #F59E0B)' : 'var(--bad, #EF4444)');
  var statusBadge = cred > 0.01 ?
    '<span class="m-badge-status-lg status-ok">● SALDO DISPONÍVEL: ' + bcmsFmtBRL(cred) + '</span>' :
    '<span class="m-badge-status-lg status-warn">● 100% EMPENHADO / ZERADO</span>';

  var h = '';
  h += '<div class="m-accent-bar" style="background:' + (u.accent || '#15803D') + ';"></div>';
  h += '<div class="m-content-wrap">';
  h += '  <div class="m-header-v2">';
  h += '    <div class="m-header-meta-row">';
  h += '      <span class="m-badge-op" style="background:' + (u.accent || '#15803D') + ';color:#fff;">🏛️ ORGANIZAÇÃO MILITAR · 12ª RM / CMA</span>';
  h += '      ' + statusBadge;
  h += '    </div>';
  h += '    <div class="m-title-row">';
  h += '      <div class="m-nc-code-block" style="display:flex;align-items:center;gap:12px;">';
  h += '        <img src="assets/logos/' + (u.logo || '12RM.png') + '" alt="" style="width:40px;height:40px;object-fit:contain;" onerror="this.src=\'assets/logos/12RM.png\'">';
  h += '        <div><h3 id="modal-title" style="margin:0;">' + bcmsEsc(u.sigla) + ' — ' + bcmsEsc(u.nome) + '</h3>';
  h += '        <span style="font-size:0.8125rem;color:var(--ink-muted);">UASGs: ' + bcmsEsc(u.ogu) + ' (OGU) e ' + bcmsEsc(u.fex) + ' (FEx) · Exercício 2026</span></div>';
  h += '      </div>';
  h += '    </div>';
  h += '  </div>';

  h += '  <div class="m-fin-section">';
  h += '    <div class="m-fin-grid">';
  h += '      <div class="m-fin-card"><span class="m-fin-label">Provisão Recebida</span><span class="m-fin-val col-prov">' + bcmsFmtBRL(prov) + '</span><span class="m-fin-sub">Dotação consolidada</span></div>';
  h += '      <div class="m-fin-card"><span class="m-fin-label">Despesas Empenhadas</span><span class="m-fin-val" style="color:var(--primary-600);">' + bcmsFmtBRL(emp) + '</span><span class="m-fin-sub">' + pctEmp.toFixed(1) + '% de execução</span></div>';
  h += '      <div class="m-fin-card"><span class="m-fin-label">Crédito Disponível</span><span class="m-fin-val ' + (cred > 0.01 ? 'col-cred' : '') + '">' + bcmsFmtBRL(cred) + '</span><span class="m-fin-sub">' + (cred > 0.01 ? 'Saldo livre em tela' : '100% comprometido') + '</span></div>';
  h += '      <div class="m-fin-card"><span class="m-fin-label">Liquidado / Pago</span><span class="m-fin-val" style="font-size:1.15rem;">' + bcmsFmtBRL(liq) + '</span><span class="m-fin-sub">Pago: ' + bcmsFmtBRL(pag) + ' (' + pctPag.toFixed(1) + '%)</span></div>';
  h += '    </div>';

  h += '    <div class="m-exec-pipeline" style="margin-top:16px;">';
  h += '      <div class="m-pipeline-header"><span class="m-pipeline-title">Funil de Execução Orçamentária da OMDS</span><span class="m-pipeline-pct" style="color:' + semCor + ';">' + pctEmp.toFixed(1) + '% executado</span></div>';
  h += '      <div class="m-pipeline-stages">';
  h += '        <div class="m-stage"><div class="m-stage-head"><span>1. Dotação Recebida</span><b>100%</b></div><div class="m-stage-track"><div class="m-stage-fill" style="width:100%;background:var(--prov);"></div></div><span class="m-stage-val">' + bcmsFmtBRL(prov) + '</span></div>';
  h += '        <div class="m-stage"><div class="m-stage-head"><span>2. Empenhado</span><b>' + pctEmp.toFixed(1) + '%</b></div><div class="m-stage-track"><div class="m-stage-fill" style="width:' + Math.min(100, pctEmp) + '%;background:var(--emp);"></div></div><span class="m-stage-val">' + bcmsFmtBRL(emp) + '</span></div>';
  h += '        <div class="m-stage"><div class="m-stage-head"><span>3. Liquidado</span><b>' + pctLiq.toFixed(1) + '% do emp.</b></div><div class="m-stage-track"><div class="m-stage-fill" style="width:' + Math.min(100, pctLiq) + '%;background:var(--gold);"></div></div><span class="m-stage-val">' + bcmsFmtBRL(liq) + '</span></div>';
  h += '        <div class="m-stage"><div class="m-stage-head"><span>4. Pago</span><b>' + pctPag.toFixed(1) + '% do liq.</b></div><div class="m-stage-track"><div class="m-stage-fill" style="width:' + Math.min(100, pctPag) + '%;background:var(--pag);"></div></div><span class="m-stage-val">' + bcmsFmtBRL(pag) + '</span></div>';
  h += '      </div>';
  h += '    </div>';

  h += '    <div style="margin-top:20px;display:flex;gap:12px;justify-content:flex-end;">';
  h += '      <button type="button" class="btn-flt" onclick="bcmsIrParaCatrimaniUG(\'' + bcmsEsc(u.ogu) + '\')" style="background:var(--primary);color:#fff;border:none;padding:8px 16px;border-radius:6px;font-weight:700;cursor:pointer;">🎖️ Ver Operações na Catrimani II ›</button>';
  h += '    </div>';
  h += '  </div>';
  h += '</div>';

  var mb = document.getElementById('modal-body');
  if(mb) mb.innerHTML = h;
  var m = document.getElementById('modal');
  if(m){ m.classList.add('open'); m.setAttribute('aria-hidden', 'false'); }
  var x = document.querySelector('.modal-x');
  if(x) x.focus();
}

function bcmsIrParaCatrimaniUG(codUg){
  var m = document.getElementById('modal');
  if(m) m.classList.remove('open');
  trocaOMDSPorKey('CATRIMANI');
  setTimeout(function(){
    var sel = document.getElementById('flt-catr-ug');
    if(sel){ sel.value = String(codUg); bcmsFiltraCatrimani(); }
  }, 150);
}

function bcmsModalKpi(tipo){
  var titulo = '', sub = '', formula = '', corpo = '';
  if(tipo === 'prov'){
    titulo = 'Dotação / Provisão Recebida';
    sub = 'Crédito orçamentário descentralizado pelos órgãos superiores (COTER, COLOG, EME)';
    formula = 'Dotação Inicial + Provisões Recebidas − Provisões Concedidas = Dotação Líquida';
    corpo = '<p>Representa a autorização orçamentária total transferida para a Unidade Gestora na LOA 2026. Este recurso torna-se apto para empenho imediato conforme os Planos de Aplicação e cronogramas de desembolso aprovados.</p>';
  } else if(tipo === 'emp'){
    titulo = 'Despesas Empenhadas';
    sub = 'Primeiro estágio da execução da despesa pública (Art. 58 da Lei nº 4.320/1964)';
    formula = 'Dotação Empenhada = Compromisso formal de pagamento criado pela Nota de Empenho (NE)';
    corpo = '<p>O empenho reserva o crédito orçamentário específico para contratação ou aquisição de bens e serviços. Garante que o crédito não seja recolhido por expiração de prazo (como as 11 NCs com vencimento em 30 de setembro).</p>';
  } else if(tipo === 'cred'){
    titulo = 'Crédito Disponível Líquido';
    sub = 'Saldo livre em tela para emissão de novos empenhos ou contratos';
    formula = 'Crédito Disponível = Provisão Recebida − Empenhado (Saldo Livre em Tela)';
    corpo = '<p>Valor residual exato disponível na UG para emissão de novas Notas de Empenho (NE). Créditos com prazo até 30 de setembro devem ser empenhados com prioridade absoluta para evitar devolução compulsória ao COTER.</p>';
  } else if(tipo === 'liq'){
    titulo = 'Despesas Liquidadas';
    sub = 'Segundo estágio da execução da despesa (Art. 63 da Lei nº 4.320/1964)';
    formula = 'Liquidação = Atesto do bem entregue ou serviço prestado via Termo de Recebimento';
    corpo = '<p>Comprova o cumprimento integral da obrigação pelo fornecedor contratado, mediante nota fiscal e atesto de conformidade técnica pelos fiscais do contrato. Precede a ordem bancária de pagamento.</p>';
  } else if(tipo === 'pag'){
    titulo = 'Despesas Pagas';
    sub = 'Terceiro e último estágio da despesa pública (Art. 64 da Lei nº 4.320/1964)';
    formula = 'Pagamento = Ordem Bancária (OB) transmitida à Conta Única do Tesouro Nacional';
    corpo = '<p>Conclusão financeira da despesa pública mediante crédito na conta corrente bancária do fornecedor. Quita a dívida do Estado e baixa a responsabilidade da administração militar.</p>';
  }

  var h = '';
  h += '<div class="m-accent-bar" style="background:var(--primary);"></div>';
  h += '<div class="m-content-wrap">';
  h += '  <div class="m-header-v2">';
  h += '    <div class="m-header-meta-row"><span class="m-badge-op">ℹ️ REGRA ORÇAMENTÁRIA & CONTABILIDADE SIAFI</span></div>';
  h += '    <h3 id="modal-title">' + titulo + '</h3>';
  h += '    <p class="m-sub" style="margin-top:4px;">' + sub + '</p>';
  h += '  </div>';
  h += '  <div style="background:var(--bg-subtle);padding:14px;border-radius:8px;border:1px solid var(--border);margin:16px 0;">';
  h += '    <span style="font-size:0.75rem;font-weight:700;color:var(--ink-muted);text-transform:uppercase;letter-spacing:0.05em;">Equação Contábil / Regra de Ouro</span>';
  h += '    <div style="font-family:var(--mono);font-size:0.95rem;font-weight:700;color:var(--primary);margin-top:4px;">' + formula + '</div>';
  h += '  </div>';
  h += '  <div style="font-size:0.9rem;line-height:1.6;color:var(--ink);">' + corpo + '</div>';
  h += '  <div style="margin-top:20px;text-align:right;">';
  h += '    <button type="button" class="btn-flt" onclick="document.getElementById(\'modal\').classList.remove(\'open\')" style="background:var(--primary);color:#fff;border:none;padding:8px 18px;border-radius:6px;font-weight:700;cursor:pointer;">Entendido</button>';
  h += '  </div>';
  h += '</div>';

  var mb = document.getElementById('modal-body');
  if(mb) mb.innerHTML = h;
  var m = document.getElementById('modal');
  if(m){ m.classList.add('open'); m.setAttribute('aria-hidden', 'false'); }
  var x = document.querySelector('.modal-x');
  if(x) x.focus();
}

function bcmsLimparFiltrosCatrimani(){
  var inBusca = document.getElementById('flt-catr-busca'); if(inBusca) inBusca.value = '';
  var inUg = document.getElementById('flt-catr-ug'); if(inUg) inUg.value = '';
  var inFonte = document.getElementById('flt-catr-fonte'); if(inFonte) inFonte.value = '';
  var inSaldo = document.getElementById('flt-catr-saldo'); if(inSaldo) inSaldo.value = '';
  bcmsFiltraCatrimani();
  bcmsToast('Filtros da Operação Catrimani redefinidos');
}

/* ==========================================================================
   MÓDULO OPERAÇÃO CATRIMANI II (Multi-UGs & Ação 21EM)
   ========================================================================== */
var CATR_INITIALIZED = false;
var CATR_FILTERED = [];
var CATR_PAGE = 1;
var CATR_PER_PAGE = 25;

function bcmsInitCatrimani(){
  CATR_INITIALIZED = true;
  if(typeof CATRDATA !== 'undefined' && CATRDATA.linhas){
    CATR_FILTERED = CATRDATA.linhas.slice();
    bcmsRenderCatrimani(1);
  }
}

function bcmsFiltraCatrimani(){
  if(!CATR_INITIALIZED) bcmsInitCatrimani();
  if(!CATRDATA || !CATRDATA.linhas) return;
  var fug = (document.getElementById('flt-catr-ug') ? document.getElementById('flt-catr-ug').value.trim() : '');
  var ffonte = (document.getElementById('flt-catr-fonte') ? document.getElementById('flt-catr-fonte').value.trim() : '');
  var fsaldo = (document.getElementById('flt-catr-saldo') ? document.getElementById('flt-catr-saldo').value.trim() : '');
  var q = (document.getElementById('flt-catr-busca') ? document.getElementById('flt-catr-busca').value.toLowerCase().trim() : '');

  var tokens = q ? q.split(/\s+/).filter(Boolean) : [];

  CATR_FILTERED = CATRDATA.linhas.filter(function(it){
    if(fug && it.ug !== fug) return false;
    if(ffonte && String(it.ug).indexOf(ffonte) !== 0) return false;
    if(fsaldo){
      if(fsaldo === 'com_saldo' && it.cred <= 0.01) return false;
      if(fsaldo === 'zerada' && it.cred > 0.01) return false;
    }
    if(tokens.length > 0){
      var rowText = (
        (it.nc || '') + ' ' +
        (it.ug || '') + ' ' +
        (it.ug_nome || '') + ' ' +
        (it.emit || '') + ' ' +
        (it.emit_nome || '') + ' ' +
        (it.acao || '') + ' ' +
        (it.pi || '') + ' ' +
        (it.pi_nome || '') + ' ' +
        (it.nd || '') + ' ' +
        (it.nd_desc || '') + ' ' +
        (it.obj || '')
      ).toLowerCase();
      for(var k=0; k<tokens.length; k++){
        if(rowText.indexOf(tokens[k]) === -1) return false;
      }
    }
    return true;
  });

  bcmsRenderCatrimani(1);
}

function bcmsFmtBRL(v){
  return bcmsBRL(v || 0);
}

function bcmsRenderCatrimani(pag){
  var p = parseInt(pag, 10);
  if(isNaN(p) || p < 1) p = 1;
  CATR_PAGE = p;

  var perPage = parseInt(CATR_PER_PAGE, 10) || 25;
  var total = (CATR_FILTERED && CATR_FILTERED.length) ? CATR_FILTERED.length : 0;
  var totPages = Math.ceil(total / perPage) || 1;
  if(CATR_PAGE > totPages) CATR_PAGE = totPages;
  if(CATR_PAGE < 1) CATR_PAGE = 1;

  var start = total > 0 ? (CATR_PAGE - 1) * perPage : 0;
  var end = Math.min(start + perPage, total);
  var slice = total > 0 ? CATR_FILTERED.slice(start, end) : [];

  var tbody = document.getElementById('tbody-catr-ncs');
  if(!tbody) return;

  var h = '';
  if(!slice.length){
    h = '<tr><td colspan="10" style="text-align:center;padding:32px;color:var(--ink-muted);">Nenhuma Nota de Crédito encontrada com os filtros selecionados.</td></tr>';
  } else {
    slice.forEach(function(it){
      var stCor = it.cred > 0.01 ? 'var(--ok, #10B981)' : 'var(--ink-muted)';
      var stTxt = it.cred > 0.01 ? 'Com Saldo' : 'Empenhada';
      h += '<tr class="tr-click" tabindex="0" role="button" onclick="bcmsOpenNCModalManual(\'' + (it.nc || '') + '\')" ' +
           'onkeydown="if(event.key===\'Enter\'||event.key===\' \'){event.preventDefault();bcmsOpenNCModalManual(\'' + (it.nc || '') + '\');}" ' +
           'title="Clique para abrir a ficha cadastral completa desta NC">' +
           '<td>' + (it.dia || '') + '</td>' +
           '<td><b class="nc-mono">' + (it.nc || '') + '</b></td>' +
           '<td><b>' + (it.ug || '') + '</b> <span class="tbl-om-sub">' + (it.ug_nome || '') + '</span></td>' +
           '<td><span class="pill-ptres">' + (it.acao || '') + '</span> · <span class="pill-pi">' + (it.pi || '') + '</span></td>' +
           '<td>' + (it.nd || '') + ' <span class="tbl-om-sub">' + ((it.nd_desc || '').slice(0, 20)) + '</span></td>' +
           '<td class="wrap-txt" title="' + (it.obj || '') + '">' + ((it.obj || '').slice(0, 55)) + '...</td>' +
           '<td class="num">' + bcmsFmtBRL(it.prov) + '</td>' +
           '<td class="num">' + bcmsFmtBRL(it.emp) + '</td>' +
           '<td class="num anchor" style="font-weight:700;">' + bcmsFmtBRL(it.cred) + '</td>' +
           '<td><span class="pill-nd" style="background:var(--track);color:' + stCor + ';font-weight:700;">' + stTxt + '</span></td>' +
           '</tr>';
    });
  }
  tbody.innerHTML = h;

  var totalGlobal = (typeof CATRDATA !== 'undefined' && CATRDATA && CATRDATA.linhas) ? CATRDATA.linhas.length : total;

  /* Contagem de NCs distintas no conjunto filtrado */
  var distinctNcs = {};
  var sumProv = 0, sumEmp = 0, sumCred = 0;
  if(CATR_FILTERED){
    CATR_FILTERED.forEach(function(it){
      if(it.nc) distinctNcs[it.nc] = true;
      sumProv += (it.prov || 0);
      sumEmp  += (it.emp || 0);
      sumCred += (it.cred || 0);
    });
  }
  var numDistinct = Object.keys(distinctNcs).length;

  var cnt = document.getElementById('cnt-catr-ncs');
  if(cnt){
    if(total === totalGlobal){
      cnt.textContent = 'Exibindo ' + (total > 0 ? (start + 1) + '–' + end : '0') + ' de ' + total + ' lançamentos (' + numDistinct + ' NCs distintas)';
    } else {
      cnt.textContent = 'Filtrado: ' + total + ' de ' + totalGlobal + ' itens (' + numDistinct + ' NCs distintas) · Exibindo ' + (total > 0 ? (start + 1) + '–' + end : '0');
    }
  }

  /* Recalcular dinamicamente o rodapé da tabela */
  var tfoot = document.querySelector('#tab-catrimani-ncs tfoot');
  if(tfoot){
    tfoot.innerHTML = '<tr>' +
      '<td colspan="6"><b>' + (total === totalGlobal ? 'TOTAL CONSOLIDADO' : 'SUBTOTAL FILTRADO') + ' · ' + numDistinct + ' NOTAS DE CRÉDITO (' + total + ' ITENS)</b></td>' +
      '<td class="num"><b>' + bcmsFmtBRL(sumProv) + '</b></td>' +
      '<td class="num"><b>' + bcmsFmtBRL(sumEmp) + '</b></td>' +
      '<td class="num anchor"><b>' + bcmsFmtBRL(sumCred) + '</b></td>' +
      '<td>—</td>' +
      '</tr>';
  }

  var ptxt = document.getElementById('pag-catr-txt');
  if(ptxt){
    ptxt.textContent = 'Página ' + CATR_PAGE + ' de ' + totPages + ' (Exibindo ' + (total > 0 ? (start + 1) + '–' + end : '0') + ' de ' + total + ')';
  }

  var btnAnt = document.getElementById('btn-catr-ant');
  var btnProx = document.getElementById('btn-catr-prox');
  if(btnAnt) btnAnt.disabled = (CATR_PAGE <= 1);
  if(btnProx) btnProx.disabled = (CATR_PAGE >= totPages || total === 0);
}

function bcmsPaginaCatrimani(delta){
  bcmsRenderCatrimani(CATR_PAGE + delta);
}

function bcmsOpenNCModalManual(ncNum){
  if(!ncNum) return;
  bcmsDetalheNC(ncNum);
}

function bcmsDetalheUG(codUg){
  if(!codUg) return;
  var u = null;
  if(typeof CATRDATA !== 'undefined' && CATRDATA && CATRDATA.por_ug){
    for(var i = 0; i < CATRDATA.por_ug.length; i++){
      if(CATRDATA.por_ug[i].cod === String(codUg)){
        u = CATRDATA.por_ug[i];
        break;
      }
    }
  }
  if(!u) return;

  var prov = u.prov || 0;
  var emp  = u.emp || 0;
  var cred = u.cred || 0;
  var liq  = u.liq || 0;
  var pag  = u.pag || 0;
  var pctEmp = prov > 0 ? (emp / prov * 100) : 0;
  var pctLiq = emp > 0 ? (liq / emp * 100) : 0;
  var pctPag = liq > 0 ? (pag / liq * 100) : 0;

  var semCor = pctEmp >= 80 ? 'var(--ok, #10B981)' : (pctEmp >= 60 ? 'var(--gold, #F59E0B)' : 'var(--bad, #EF4444)');
  var statusBadge = cred > 0.01 ?
    '<span class="m-badge-status-lg status-ok">● SALDO DISPONÍVEL: ' + bcmsFmtBRL(cred) + '</span>' :
    '<span class="m-badge-status-lg status-warn">● 100% EMPENHADO / ZERADO</span>';

  var h = '';
  h += '<div class="m-accent-bar" style="background:linear-gradient(90deg, #15803D 0%, #10B981 50%, #3B82F6 100%);"></div>';
  h += '<div class="m-content-wrap">';

  /* Cabeçalho */
  h += '  <div class="m-header-v2">';
  h += '    <div class="m-header-meta-row">';
  h += '      <span class="m-badge-op">🛡️ OPERAÇÃO CATRIMANI II · MULTI-UG</span>';
  h += '      ' + statusBadge;
  h += '    </div>';
  h += '    <div class="m-title-row">';
  h += '      <div class="m-nc-code-block">';
  h += '        <h3 id="modal-title">UG ' + bcmsEsc(u.cod) + ' — ' + bcmsEsc(u.nome) + '</h3>';
  h += '      </div>';
  h += '      <button type="button" class="m-btn-pill" onclick="bcmsFiltrarPorUG(\'' + bcmsEsc(u.cod) + '\')" title="Ver no Extrato de Notas de Crédito">';
  h += '        🔍 Filtrar no Extrato';
  h += '      </button>';
  h += '    </div>';
  h += '    <div class="m-sub-meta">';
  h += '      <span class="m-meta-chip">🏛️ Unidade Gestora Executora da Amazônia</span>';
  h += '      <span class="m-meta-chip">📑 ' + (u.n_ncs || (u.ncs ? u.ncs.length : 0)) + ' Notas de Crédito Vinculadas</span>';
  h += '      <span class="m-meta-chip">⚙️ Ação 21EM · Exercício 2026</span>';
  h += '    </div>';
  h += '  </div>';

  /* Balanço Financeiro */
  h += '  <div class="m-fin-section">';
  h += '    <div class="m-fin-grid">';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Recebido (Dotação)</span>';
  h += '        <span class="m-fin-val col-prov">' + bcmsFmtBRL(prov) + '</span>';
  h += '        <span class="m-fin-sub">Total descentralizado à UG</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Empenhado</span>';
  h += '        <span class="m-fin-val" style="color:var(--primary-600);">' + bcmsFmtBRL(emp) + '</span>';
  h += '        <span class="m-fin-sub">' + pctEmp.toFixed(1) + '% da dotação</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Crédito Disponível Líquido</span>';
  h += '        <span class="m-fin-val ' + (cred > 0.01 ? 'col-cred' : '') + '">' + bcmsFmtBRL(cred) + '</span>';
  h += '        <span class="m-fin-sub">' + (cred > 0.01 ? 'Saldo livre p/ novos empenhos' : 'Dotação 100% comprometida') + '</span>';
  h += '      </div>';
  h += '      <div class="m-fin-card">';
  h += '        <span class="m-fin-label">Liquidado / Pago</span>';
  h += '        <span class="m-fin-val" style="font-size:1.15rem;">' + bcmsFmtBRL(liq) + '</span>';
  h += '        <span class="m-fin-sub">Pago: ' + bcmsFmtBRL(pag) + ' (' + pctPag.toFixed(1) + '%)</span>';
  h += '      </div>';
  h += '    </div>';

  /* Pipeline de Execução Orçamentária */
  h += '    <div class="m-exec-pipeline">';
  h += '      <div class="m-pipeline-header">';
  h += '        <span class="m-pipeline-title">Funil de Execução da Unidade Gestora</span>';
  h += '        <span class="m-pipeline-pct" style="color:' + semCor + ';">' + pctEmp.toFixed(1) + '% executado</span>';
  h += '      </div>';
  h += '      <div class="m-pipeline-stages">';
  h += '        <div class="m-stage-item">';
  h += '          <div class="m-stage-info"><span class="m-stage-name">1. Empenho</span><span class="m-stage-val">' + pctEmp.toFixed(1) + '% (' + bcmsFmtBRL(emp) + ')</span></div>';
  h += '          <div class="m-stage-track"><div class="m-stage-fill" style="width:' + Math.min(100, Math.max(2, pctEmp)) + '%;background:' + semCor + ';"></div></div>';
  h += '        </div>';
  h += '        <div class="m-stage-item">';
  h += '          <div class="m-stage-info"><span class="m-stage-name">2. Liquidação</span><span class="m-stage-val">' + pctLiq.toFixed(1) + '% s/ empenho (' + bcmsFmtBRL(liq) + ')</span></div>';
  h += '          <div class="m-stage-track"><div class="m-stage-fill" style="width:' + Math.min(100, Math.max(2, pctLiq)) + '%;background:#3B82F6;"></div></div>';
  h += '        </div>';
  h += '        <div class="m-stage-item">';
  h += '          <div class="m-stage-info"><span class="m-stage-name">3. Pagamento</span><span class="m-stage-val">' + pctPag.toFixed(1) + '% s/ liquidado (' + bcmsFmtBRL(pag) + ')</span></div>';
  h += '          <div class="m-stage-track"><div class="m-stage-fill" style="width:' + Math.min(100, Math.max(2, pctPag)) + '%;background:#8B5CF6;"></div></div>';
  h += '        </div>';
  h += '      </div>';
  h += '    </div>';
  h += '  </div>';

  /* Seção Despesas por ND */
  if(u.nds_list && u.nds_list.length > 0){
    h += '  <div class="m-justif-card" style="border-left-color:#15803D;margin-top:14px;">';
    h += '    <div class="m-justif-header"><span class="m-justif-title">📊 Desdobramento por Natureza de Despesa (' + u.nds_list.length + ' NDs)</span></div>';
    h += '    <div class="tbl-scroll"><table class="det det-compact" style="width:100%;font-size:0.78125rem;">';
    h += '      <thead><tr><th>ND</th><th>Descrição</th><th class="num">Recebido</th><th class="num">Empenhado</th><th class="num">Disponível</th><th class="num">% Exec</th></tr></thead><tbody>';
    for(var k = 0; k < u.nds_list.length; k++){
      var ndo = u.nds_list[k];
      var ndPct = ndo.prov > 0 ? (ndo.emp / ndo.prov * 100) : 0;
      h += '<tr>' +
           '<td class="mono2"><b>' + bcmsEsc(ndo.nd) + '</b></td>' +
           '<td>' + bcmsEsc(ndo.nd_desc || '—') + '</td>' +
           '<td class="num">' + bcmsFmtBRL(ndo.prov) + '</td>' +
           '<td class="num">' + bcmsFmtBRL(ndo.emp) + '</td>' +
           '<td class="num anchor" style="font-weight:700;">' + bcmsFmtBRL(ndo.cred) + '</td>' +
           '<td class="num"><b>' + ndPct.toFixed(1) + '%</b></td>' +
           '</tr>';
    }
    h += '    </tbody></table></div>';
    h += '  </div>';
  }

  /* Seção Notas de Crédito da UG */
  if(u.ncs && u.ncs.length > 0){
    h += '  <div class="m-justif-card" style="border-left-color:var(--primary-600);margin-top:14px;">';
    h += '    <div class="m-justif-header" style="display:flex;justify-content:space-between;align-items:center;">';
    h += '      <span class="m-justif-title">📋 Notas de Crédito Recebidas (' + u.ncs.length + ' NCs)</span>';
    h += '      <span style="font-size:0.75rem;color:var(--ink-muted);">Clique em qualquer NC para ver a ficha cadastral</span>';
    h += '    </div>';
    h += '    <div class="tbl-scroll" style="max-height:260px;"><table class="det det-compact" style="width:100%;font-size:0.78125rem;">';
    h += '      <thead><tr><th>Emissão</th><th>Número NC</th><th>ND</th><th class="num">Recebido</th><th class="num">Empenhado</th><th class="num">Saldo Disp.</th><th>Status</th></tr></thead><tbody>';
    for(var mIdx = 0; mIdx < u.ncs.length; mIdx++){
      var nco = u.ncs[mIdx];
      var sColor = nco.cred > 0.01 ? 'var(--ok, #10B981)' : 'var(--ink-muted)';
      var sText = nco.cred > 0.01 ? 'Com Saldo' : 'Empenhada';
      h += '<tr class="tr-click" onclick="bcmsOpenNCModalManual(\'' + bcmsEsc(nco.nc) + '\')" title="Abrir ficha da NC ' + bcmsEsc(nco.nc) + '">' +
           '<td>' + bcmsEsc(nco.dia || '—') + '</td>' +
           '<td><b class="nc-mono" style="color:var(--primary-600);">' + bcmsEsc(nco.nc) + '</b></td>' +
           '<td class="mono2">' + bcmsEsc(nco.nd || '—') + '</td>' +
           '<td class="num">' + bcmsFmtBRL(nco.prov) + '</td>' +
           '<td class="num">' + bcmsFmtBRL(nco.emp) + '</td>' +
           '<td class="num anchor" style="font-weight:700;">' + bcmsFmtBRL(nco.cred) + '</td>' +
           '<td><span class="pill-nd" style="color:' + sColor + ';font-weight:700;">' + sText + '</span></td>' +
           '</tr>';
    }
    h += '    </tbody></table></div>';
    h += '  </div>';
  }

  /* Rodapé de Ações */
  h += '  <div class="m-footer-actions-v2">';
  h += '    <button type="button" class="m-btn-pill primary" onclick="bcmsFiltrarPorUG(\'' + bcmsEsc(u.cod) + '\')">🔍 Ver NCs Desta UG no Extrato</button>';
  h += '    <button type="button" class="m-btn-pill" onclick="bcmsCelClose()">Fechar Janela ✕</button>';
  h += '  </div>';

  h += '</div>';

  document.getElementById('modal-body').innerHTML = h;
  var modalEl = document.getElementById('modal');
  modalEl.classList.add('open');
  modalEl.setAttribute('aria-hidden', 'false');
  var xBtn = document.querySelector('.modal-x');
  if(xBtn) xBtn.focus();
}

function bcmsFiltrarPorUG(codUg){
  bcmsCelClose();
  var sel = document.getElementById('flt-catr-ug');
  if(sel){
    sel.value = codUg;
  }
  var busca = document.getElementById('flt-catr-busca');
  if(busca) busca.value = '';
  bcmsFiltraCatrimani();
  var tab = document.getElementById('tab-catrimani-ncs');
  if(tab && tab.scrollIntoView){
    tab.scrollIntoView({behavior: 'smooth', block: 'start'});
  }
  bcmsToast('🔍 Extrato filtrado pela UG ' + codUg);
}

function bcmsExportCatrimaniExcel(){
  if(!CATRDATA || !CATRDATA.linhas) return;
  var list = CATR_FILTERED && CATR_FILTERED.length ? CATR_FILTERED : CATRDATA.linhas;

  var xml = '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<?mso-application progid="Excel.Sheet"?>\n' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n' +
    ' xmlns:o="urn:schemas-microsoft-com:office:office"\n' +
    ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n' +
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n' +
    ' <Styles>\n' +
    '  <Style ss:ID="Default" ss:Name="Normal"><Font ss:FontName="Calibri" ss:Size="11"/></Style>\n' +
    '  <Style ss:ID="sHeader"><Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#15803D" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/></Style>\n' +
    '  <Style ss:ID="sMoeda"><NumberFormat ss:Format="R$ #,##0.00"/></Style>\n' +
    ' </Styles>\n' +
    ' <Worksheet ss:Name="Operacao_Catrimani_II">\n' +
    '  <Table>\n' +
    '   <Row ss:StyleID="sHeader">\n' +
    '    <Cell><Data ss:Type="String">EMISSÃO</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">NOTA DE CRÉDITO</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">CÓD. UG</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">UG EXECUTORA</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">AÇÃO</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">PLANO INTERNO</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">DESCRIÇÃO PI</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">NATUREZA DESPESA</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">DESCRIÇÃO ND</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">OBJETO / JUSTIFICATIVA</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">RECEBIDO (R$)</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">EMPENHADO (R$)</Data></Cell>\n' +
    '    <Cell><Data ss:Type="String">DISPONÍVEL (R$)</Data></Cell>\n' +
    '   </Row>\n';

  list.forEach(function(it){
    xml += '   <Row>\n' +
      '    <Cell><Data ss:Type="String">' + (it.dia || '') + '</Data></Cell>\n' +
      '    <Cell><Data ss:Type="String">' + (it.nc || '') + '</Data></Cell>\n' +
      '    <Cell><Data ss:Type="String">' + (it.ug || '') + '</Data></Cell>\n' +
      '    <Cell><Data ss:Type="String">' + (it.ug_nome || '') + '</Data></Cell>\n' +
      '    <Cell><Data ss:Type="String">' + (it.acao || '') + '</Data></Cell>\n' +
      '    <Cell><Data ss:Type="String">' + (it.pi || '') + '</Data></Cell>\n' +
      '    <Cell><Data ss:Type="String">' + (it.pi_nome || '') + '</Data></Cell>\n' +
      '    <Cell><Data ss:Type="String">' + (it.nd || '') + '</Data></Cell>\n' +
      '    <Cell><Data ss:Type="String">' + (it.nd_desc || '') + '</Data></Cell>\n' +
      '    <Cell><Data ss:Type="String">' + (it.obj || '') + '</Data></Cell>\n' +
      '    <Cell ss:StyleID="sMoeda"><Data ss:Type="Number">' + (it.prov || 0) + '</Data></Cell>\n' +
      '    <Cell ss:StyleID="sMoeda"><Data ss:Type="Number">' + (it.emp || 0) + '</Data></Cell>\n' +
      '    <Cell ss:StyleID="sMoeda"><Data ss:Type="Number">' + (it.cred || 0) + '</Data></Cell>\n' +
      '   </Row>\n';
  });

  xml += '  </Table>\n </Worksheet>\n</Workbook>';

  var blob = new Blob([xml], {type: 'application/vnd.ms-excel;charset=utf-8;'});
  var link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'operacao_catrimani_credito_' + new Date().toISOString().slice(0,10) + '.xls';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  bcmsToast('📊 Planilha da Operação Catrimani exportada com sucesso (' + list.length + ' linhas)!');
}

"""

# ---------------- main ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--local", help="caminho de um xlsx ou csv local (teste)")
    ap.add_argument("--url", help="URL do CSV ou planilha publicada no Google Sheets")
    ap.add_argument("--date", help="data do snapshot YYYY-MM-DD (default: hoje)")
    ap.add_argument("--file-id", default=(os.environ.get("DRIVE_FILE_ID") or DEFAULT_FILE_ID))
    args = ap.parse_args()

    data_str = args.date or datetime.date.today().isoformat()
    target = args.local or args.url or os.environ.get("SHEETS_CSV_URL") or args.file_id
    path = args.local if (args.local and os.path.exists(args.local)) else baixar(target)
    print("Fonte:", path)
    res, periodo, alertas, catrimani_data, omds_totais = etl(path)
    for a in alertas:
        print("[ALERTA]", a)
    hist = atualizar_historico(res, data_str)
    html_out = montar_pagina(res, hist, data_str, periodo, alertas, catrimani_data, omds_totais)

    os.makedirs(SITE, exist_ok=True)
    os.makedirs(os.path.join(SITE, "data"), exist_ok=True)
    src_logos = os.path.join(HERE, "assets", "logos")
    dst_logos = os.path.join(SITE, "assets", "logos")
    os.makedirs(dst_logos, exist_ok=True)
    if os.path.isdir(src_logos):
        for fn in os.listdir(src_logos):
            if fn.lower().endswith((".png", ".jpg", ".jpeg", ".webp", ".svg")):
                try:
                    shutil.copyfile(os.path.join(src_logos, fn), os.path.join(dst_logos, fn))
                except Exception as e:
                    print("[AVISO] não copiei", fn, ":", e)
    with open(os.path.join(SITE, "index.html"), "w", encoding="utf-8") as f:
        f.write(html_out)
    with open(os.path.join(SITE, "data", "history.json"), "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=1)

    tot = {k: sum(res[c][k] for c, _ in ALVOS) for k in ("prov", "cred", "emp", "liq", "pag")}
    print(f"OK -> {os.path.join(SITE,'index.html')}")
    print(f"Periodo={periodo} | Credito Disp total={brl(tot['cred'])} | historico {len(hist)} dia(s)")

if __name__ == "__main__":
    main()
