"""Solicitação 9 — Referência Local.
Conta posts por local (multirrótulo, separador ";"), sem agrupar variantes de grafia;
lista pares de variantes para aprovação da autora.
Uso: python3 solicitacao9.py <planilha.xlsx>
Gera Solicitacao9_Referencia_local.xlsx (sem gráfico: não solicitado).
"""
import sys
import re
import itertools
import unicodedata
import difflib
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao9_Referencia_local.xlsx"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
COL = "Referência Local"
NI = "Não identificado"
AMPLOS = ["Imperatriz", "Maranhão"]
TOP = 30
LIMIAR = 0.85  # similaridade mínima (difflib) entre formas normalizadas

# ---- Decisões da autora (04/11/2026), em decisoes_locais.py
sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
from decisoes_locais import DECISOES

# Respostas que contradizem outras (detectadas por transitividade): NÃO aplicadas até confirmação
CONFLITOS = {}  # os dois conflitos de 04/11 foram resolvidos pela autora (ver decisoes_locais.py)
_par = {}


def _f(x):
    _par.setdefault(x, x)
    while _par[x] != x:
        _par[x] = _par[_par[x]]
        x = _par[x]
    return x


for a, b, d, _ in DECISOES:
    if d == "mesma" and (a, b, d) not in CONFLITOS:
        _par[_f(a)] = _f(b)
DISTINTOS = {frozenset({a, b}) for a, b, d, _ in DECISOES if d == "diferente"}
for par_ in DISTINTOS:  # nenhuma decisão "diferente" pode cair dentro de um grupo
    x, y = tuple(par_)
    assert _f(x) != _f(y), par_
_grupos = {}
for x in list(_par):
    _grupos.setdefault(_f(x), []).append(x)
_grupos = [sorted(v) for v in _grupos.values() if len(v) > 1]


df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
N_tot = {**N.to_dict(), "Total": len(df)}

# separar por ";", tirar espaços nas pontas, um local no máximo uma vez por post; grafia mantida como está
df["loc_orig"] = df[COL].astype(str).str.split(";").apply(lambda l: sorted({x.strip() for x in l if x.strip()}))
# forma adotada de cada grupo = grafia mais frequente (posts); empate -> a que aparece primeiro como Variante A
_freq = df["loc_orig"].explode().value_counts()
_ordemA = {a: i for i, (a, *_ ) in enumerate(DECISOES)}
AGRUPAR = {}
for g in _grupos:
    for v_ in g:
        assert v_ in _freq.index, f"grafia não encontrada na planilha: {v_}"
    canon = sorted(g, key=lambda v_: (-_freq[v_], _ordemA.get(v_, 999), v_))[0]
    for v_ in g:
        if v_ != canon:
            AGRUPAR[v_] = canon
# aplica os agrupamentos aprovados; o set garante que o local conte no máximo uma vez por post
df["loc"] = df["loc_orig"].apply(lambda l: sorted({AGRUPAR.get(x, x) for x in l}))
POSTS_DUPLA = int(df["loc_orig"].apply(lambda l: len(l) != len({AGRUPAR.get(x, x) for x in l})).sum())
assert df["loc"].apply(len).min() >= 1
assert not df["loc"].apply(lambda l: NI in l and len(l) > 1).any()
long = df[["url ", "Perfil", "loc"]].explode("loc").reset_index(drop=True)

# ---- Não identificado
ni = df.assign(ni=df["loc"].apply(lambda l: l == [NI])).groupby("Perfil")["ni"].sum().reindex(PERFIS)
t_ni = pd.DataFrame({"Perfil": PERFIS + ["Total"],
                     "N de posts do perfil": [N[p] for p in PERFIS] + [len(df)],
                     "Não identificado (N)": list(ni.values) + [int(ni.sum())]})
t_ni["Não identificado (%)"] = t_ni["Não identificado (N)"] / t_ni["N de posts do perfil"]
assert t_ni.iloc[-1]["Não identificado (N)"] == 390


def top(grupo, excluir):
    d = long if grupo == "Total" else long[long["Perfil"] == grupo]
    c = d.loc[~d["loc"].isin(excluir), "loc"].value_counts().rename_axis("Local").reset_index(name="N de posts")
    c = c.sort_values(["N de posts", "Local"], ascending=[False, True])
    c["Posição"] = c["N de posts"].rank(method="min", ascending=False).astype(int)
    corte = c[c["Posição"] <= TOP].copy()          # empates na 30ª posição são incluídos
    rep = corte["Posição"].duplicated(keep=False)
    corte["Posição"] = corte["Posição"].astype(str) + rep.map({True: "=", False: ""})
    corte["% dos posts do perfil"] = corte["N de posts"] / N_tot[grupo]
    corte["Local"] = [l + (" *" if n < 10 else "") for l, n in zip(corte["Local"], corte["N de posts"])]
    return corte[["Posição", "Local", "N de posts", "% dos posts do perfil"]].reset_index(drop=True)


listas = {}
for g in PERFIS + ["Total"]:
    listas[g] = (top(g, [NI]), top(g, [NI] + AMPLOS))

# o que foi excluído na versão B
excl = []
for g in PERFIS + ["Total"]:
    d = long if g == "Total" else long[long["Perfil"] == g]
    for a in AMPLOS:
        k = int((d["loc"] == a).sum())
        excl.append([g, a, k, k / N_tot[g]])
t_excl = pd.DataFrame(excl, columns=["Perfil", "Local excluído na versão B", "N de posts", "% dos posts do perfil"])

# ---- Variantes para aprovação (calculadas sobre as grafias ORIGINAIS; não agrupa; só lista)
long_o = df[["url ", "Perfil", "loc_orig"]].explode("loc_orig").reset_index(drop=True).rename(columns={"loc_orig": "loc"})
cont = long_o["loc"].value_counts()
perf_loc = long_o.groupby("loc")["Perfil"].apply(lambda s: ", ".join(p for p in PERFIS if p in set(s)))
PREF = r"^(avenida|av|rua|r|bairro da|bairro do|bairro|park|parque|pq)\s+"


def norm(x):
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode().lower()
    x = re.sub(r"[^a-z0-9 ]", " ", x)
    return re.sub(r"\s+", " ", x).strip()


def sem_prefixo(x):
    return re.sub(PREF, "", norm(x)).replace(" de ", " ").replace(" da ", " ").replace(" do ", " ")


def num(x):
    m = re.search(r"\b(i{1,3}|iv|[0-9]+)$", norm(x))
    return m.group(1) if m else ""


def tipo(a, b):
    na, nb = norm(a), norm(b)
    if na == nb:
        return "Acento, maiúscula ou pontuação"
    if sem_prefixo(a) == sem_prefixo(b):
        return "Abreviação ou prefixo (Av./Avenida, Rua, Bairro, Parque/Park)"
    ra = re.sub(r"\s*\b(i{1,3}|iv|[0-9]+)$", "", na)
    rb = re.sub(r"\s*\b(i{1,3}|iv|[0-9]+)$", "", nb)
    if ra == rb and num(a) != num(b):
        romano = {"i": "1", "ii": "2", "iii": "3", "iv": "4"}
        if romano.get(num(a), num(a)) == romano.get(num(b), num(b)) and num(a) and num(b):
            return "Numeração escrita de forma diferente (romano × arábico)"
        return "Numeração diferente (pode ser outro local: confirmar)"
    if norm(a).replace("-ma", "").strip() == norm(b).replace("-ma", "").strip() or \
       na.replace(" ma", "") == nb.replace(" ma", ""):
        return "Sufixo de UF (-MA)"
    return "Grafia semelhante (1 a 2 letras; pode ser outro local: confirmar)"


pares = []
locais = [l for l in cont.index if l != NI]
for a, b in itertools.combinations(locais, 2):
    na, nb = norm(a), norm(b)
    r = difflib.SequenceMatcher(None, na, nb).ratio()
    rs = difflib.SequenceMatcher(None, sem_prefixo(a), sem_prefixo(b)).ratio()
    if na == nb or sem_prefixo(a) == sem_prefixo(b) or r >= LIMIAR or rs >= 0.95:
        A, B = (a, b) if cont[a] >= cont[b] else (b, a)
        pares.append([A, int(cont[A]), B, int(cont[B]),
                      f"A: {perf_loc[A]} | B: {perf_loc[B]}", tipo(A, B), round(max(r, rs), 2)])
t_var = pd.DataFrame(pares, columns=["Variante A", "N", "Variante B", "N ", "Perfil onde aparece",
                                     "Tipo de diferença (sugestão de leitura)", "Similaridade (0–1)"])
ordem_tipo = {t: i for i, t in enumerate([
    "Acento, maiúscula ou pontuação", "Abreviação ou prefixo (Av./Avenida, Rua, Bairro, Parque/Park)",
    "Sufixo de UF (-MA)", "Numeração escrita de forma diferente (romano × arábico)",
    "Grafia semelhante (1 a 2 letras; pode ser outro local: confirmar)",
    "Numeração diferente (pode ser outro local: confirmar)"])}
_conf_pares = {frozenset({a, b}) for a, b, _ in CONFLITOS}
_sem = {frozenset({a, b}) for a, b, d, _ in DECISOES if d == "sem_decisao"}


def _status(a, b):
    k = frozenset({a, b})
    if k in _conf_pares:
        return "CONFLITO entre respostas — mantido separado; confirmar"
    if k in DISTINTOS:
        return "REJEITADO em 04/11 — locais distintos (não agrupar)"
    if AGRUPAR.get(a, a) == AGRUPAR.get(b, b):
        extra = " (por transitividade; sem resposta direta)" if k in _sem else ""
        return f"APROVADO em 04/11 — agrupado como \"{AGRUPAR.get(a, a)}\"" + extra
    if k in _sem:
        return "Sem decisão da autora"
    return "Aguardando decisão"


t_var["Status"] = [_status(a, b) for a, b in zip(t_var["Variante A"], t_var["Variante B"])]
t_var["Comentário da autora"] = [
    " / ".join(c for a2, b2, d, c in DECISOES if frozenset({a2, b2}) == frozenset({a, b}) and c) for a, b in
    zip(t_var["Variante A"], t_var["Variante B"])]
t_var["_o"] = t_var["Tipo de diferença (sugestão de leitura)"].map(ordem_tipo)
_os = lambda x: 0 if x.startswith("CONFLITO") else 1 if x.startswith("Sem") else 2 if x.startswith("Aguard") else 3 if x.startswith("APROV") else 4
t_var["_s"] = t_var["Status"].map(_os)
t_var = t_var.sort_values(["_s", "_o", "N"], ascending=[True, True, False]).drop(columns=["_s", "_o"]).reset_index(drop=True)
# aba de agrupamentos aplicados
_ag = []
for g in _grupos:
    canon = AGRUPAR.get(g[0], g[0])
    vs = sorted(g, key=lambda v_: -_freq[v_])
    _ag.append([canon, "; ".join(f"{v_} ({int(_freq[v_])})" for v_ in vs),
                int(df["loc"].apply(lambda l: canon in l).sum())])
t_ag = pd.DataFrame(_ag, columns=["Forma adotada", "Grafias agrupadas (N de posts de cada uma, antes do agrupamento)",
                                  "N de posts do grupo (após agrupamento)"]).sort_values(
    "N de posts do grupo (após agrupamento)", ascending=False).reset_index(drop=True)

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia; original não alterado)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "\"Não identificado\" fora das listas Top 30 (390 posts; ver aba Nao_identificado). Versão B das listas: também sem \"Imperatriz\" e \"Maranhão\" (ver aba Excluidos_versao_B)."],
    ["Contagem", "Célula separada por \";\" e espaços das pontas removidos; cada local conta no máximo uma vez por post (posts, não células). Grafias mantidas como na planilha, exceto os agrupamentos aprovados (ver linha \"Agrupamentos aprovados\")."],
    ["Fórmula do %", "% = nº de posts do perfil que citam o local ÷ total de posts do perfil (Ego 414; Roma 1.131; IO 2.721; Total 4.266) × 100. Um post cita em média vários locais; a soma passa de 100%."],
    ["Top 30 e empates", "Ranking de competição (1, 2, 2, 4...). \"=\" indica empate. Locais empatados na 30ª posição foram todos incluídos; por isso algumas listas têm mais de 30 linhas."],
    ["*", "Local com menos de 10 posts no grupo da lista."],
    ["Versões A e B", "A = sem \"Não identificado\". B = sem \"Não identificado\", \"Imperatriz\" e \"Maranhão\". As posições foram recalculadas na versão B."],
    ["Variantes", "Pares listados por (1) forma igual após retirar acentos, maiúsculas e pontuação; (2) forma igual após retirar prefixos Av./Avenida, Rua, Bairro, Parque/Park e de/da/do; ou (3) similaridade ≥ 0,85 (difflib.SequenceMatcher) entre as formas normalizadas. A coluna \"Tipo de diferença\" é sugestão de leitura; pares com numeração diferente (ex.: Parque Alvorada I × II) ou 1–2 letras diferentes (ex.: Vila Machado × Vila Macedo) podem ser locais distintos. NADA FOI AGRUPADO: depende da aprovação da autora."],
    ["Agrupamentos aprovados", f"Decisões da autora de 04/11/2026 (arquivo decisoes_locais.py). Aplicadas por transitividade (se A = B e B = C, então A = C). Forma adotada de cada grupo = grafia mais frequente; em empate, a que aparece como Variante A na resposta da autora. {len(_grupos)} grupos, {len(AGRUPAR)} grafias substituídas (ver aba Agrupamentos_aplicados). Posts que tinham duas grafias do mesmo grupo passaram a contar uma vez: {POSTS_DUPLA}."],
    ["Conflitos resolvidos", "Vila Redenção × Vila Redenção 2: diferentes (Vila Redenção = I; Vila Redenção 2 = Vila Redenção II). Avenida Ceará × Ceará: diferentes (Avenida Ceará = Rua Ceará; Ceará = estado). Correções da autora em 04/11/2026."],
    ["Confirmados em 04/11", "Pedro Neiva de Santana = Avenida Pedro Neiva de Santana; Caema = Bairro da CAEMA. Não há outros pares sem decisão."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a Referência Local."],
    ["Script", "solicitacao9.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t_ni.to_excel(w, sheet_name="Nao_identificado", index=False)
    for g in PERFIS + ["Total"]:
        nome = "Top30_" + g.replace(" ", "_")
        a, b = listas[g]
        pd.DataFrame([[f"Versão A — {g}: sem \"Não identificado\" (N de posts = {N_tot[g]:,})".replace(",", ".")]]).to_excel(
            w, sheet_name=nome, index=False, header=False, startrow=0, startcol=0)
        a.to_excel(w, sheet_name=nome, index=False, startrow=1, startcol=0)
        pd.DataFrame([[f"Versão B — {g}: sem \"Não identificado\", \"Imperatriz\" e \"Maranhão\""]]).to_excel(
            w, sheet_name=nome, index=False, header=False, startrow=0, startcol=5)
        b.to_excel(w, sheet_name=nome, index=False, startrow=1, startcol=5)
    t_excl.to_excel(w, sheet_name="Excluidos_versao_B", index=False)
    t_var.to_excel(w, sheet_name="Variantes_para_aprovacao", index=False)
    t_ag.to_excel(w, sheet_name="Agrupamentos_aplicados", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
ws = wb["Nao_identificado"]
for c in ws[1] + ws[ws.max_row]:
    c.font = Font(bold=True)
for r in ws.iter_rows(min_row=2):
    r[1].number_format = r[2].number_format = "#,##0"
    r[3].number_format = "0.0%"
for L, wd in zip("ABCD", (20, 22, 22, 22)):
    ws.column_dimensions[L].width = wd
for g in PERFIS + ["Total"]:
    ws = wb["Top30_" + g.replace(" ", "_")]
    for r in (1, 2):
        for c in ws[r]:
            c.font = Font(bold=True)
    for row in ws.iter_rows(min_row=3):
        for idx in (2, 7):
            if row[idx].value is not None:
                row[idx].number_format = "#,##0"
        for idx in (3, 8):
            if row[idx].value is not None:
                row[idx].number_format = "0.0%"
    for L, wd in zip("ABCDEFGHI", (9, 30, 12, 20, 3, 9, 30, 12, 20)):
        ws.column_dimensions[L].width = wd
ws = wb["Excluidos_versao_B"]
for c in ws[1]:
    c.font = Font(bold=True)
for r in ws.iter_rows(min_row=2):
    r[2].number_format = "#,##0"
    r[3].number_format = "0.0%"
for L, wd in zip("ABCD", (20, 28, 12, 20)):
    ws.column_dimensions[L].width = wd
ws = wb["Variantes_para_aprovacao"]
for c in ws[1]:
    c.font = Font(bold=True)
    c.alignment = Alignment(wrap_text=True)
for L, wd in zip("ABCDEFGHI", (28, 8, 28, 8, 40, 48, 12, 45, 50)):
    ws.column_dimensions[L].width = wd
ws = wb["Agrupamentos_aplicados"]
for c in ws[1]:
    c.font = Font(bold=True)
    c.alignment = Alignment(wrap_text=True)
for L, wd in zip("ABC", (30, 90, 22)):
    ws.column_dimensions[L].width = wd
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
    row[2].number_format = "#,##0"
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

pd.set_option("display.width", 250)
print(t_ni.to_string()); print(t_excl.to_string())
for g in PERFIS + ["Total"]:
    a, b = listas[g]
    print(g, len(a), len(b)); print(a.head(10).to_string()); print(b.head(10).to_string())
print(len(t_var)); print(t_var.to_string())

# ---- Tabela derivada: Referência Local original × agrupada (a planilha original não é alterada)
DERIV = "Solicitacao9_Referencia_local_agrupada.xlsx"
der = df[["url ", "Perfil", COL]].copy()
der["Referência Local (agrupada)"] = df["loc"].apply(lambda l: "; ".join(l))
der["Mudou com o agrupamento?"] = (df["loc"] != df["loc_orig"]).map({True: "Sim", False: "Não"})
assert len(der) == 4266 and der["url "].is_unique
with pd.ExcelWriter(DERIV, engine="openpyxl") as w:
    der.to_excel(w, sheet_name="Dados", index=False)
    pd.DataFrame([
        ["Arquivo de origem", NOME_ORIGINAL + " (versão congelada de 04/11/2026; original não alterado)"],
        ["Linhas", "4.266 (uma por post; chave = url)"],
        ["Coluna agrupada", "Locais separados por \";\", agrupamentos aprovados aplicados (decisoes_locais.py), cada local uma vez por post, em ordem alfabética, unidos por \"; \"."],
        ["Mudou com o agrupamento?", f"Sim em {int((der['Mudou com o agrupamento?'] == 'Sim').sum())} posts (grafia substituída ou local repetido no post)."],
        ["Script", "solicitacao9.py"],
    ], columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)
print("derivada:", der["Mudou com o agrupamento?"].value_counts().to_dict())
