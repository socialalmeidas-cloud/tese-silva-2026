"""Solicitação 10 — Calendário (Período Especial; multirrótulo, separador ";").
Uso: python3 solicitacao10.py <planilha.xlsx>
Gera Solicitacao10_Calendario.xlsx e .png.
"""
import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao10_Calendario.xlsx"
PNG = "Solicitacao10_Calendario.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
COL = "Período Especial"
NH = "Não há"
TOP = 15
MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
INV = {v: k for k, v in MESES.items()}

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
N_tot = {**N.to_dict(), "Total": len(df)}

df["ev"] = df[COL].astype(str).str.split(";").apply(lambda l: sorted({x.strip() for x in l if x.strip()}))
assert df["ev"].apply(len).min() >= 1
assert not df["ev"].apply(lambda l: NH in l and len(l) > 1).any()
df["algum"] = df["ev"].apply(lambda l: l != [NH])
p = df["Período"].astype(str).str.strip().str.strip("()").str.split("-", expand=True)
df["data"] = pd.to_datetime(dict(year=p[1].astype(int), month=p[0].str.lower().map(MESES), day=1))
long = df[["url ", "Perfil", "data", "ev"]].explode("ev").reset_index(drop=True)
ev_long = long[long["ev"] != NH]

# ---- Por_perfil
cont = ev_long["ev"].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
cont["pos"] = cont["n"].rank(method="min", ascending=False).astype(int)
TOPEV = cont[cont["pos"] <= TOP]["e"].tolist()  # empates na 15ª posição incluídos
ct = pd.crosstab(ev_long["ev"], ev_long["Perfil"]).reindex(index=TOPEV, columns=PERFIS, fill_value=0)
alg = df.groupby("Perfil")["algum"].sum().reindex(PERFIS)
nh = (~df["algum"]).groupby(df["Perfil"]).sum().reindex(PERFIS)
linhas = pd.concat([pd.DataFrame([alg.values, nh.values], index=["Algum evento (qualquer rótulo diferente de Não há)", NH],
                                 columns=PERFIS), ct])
linhas["Total"] = linhas.sum(axis=1)
t1 = pd.DataFrame(index=linhas.index)
for c in PERFIS + ["Total"]:
    t1[f"{c} (N)"] = linhas[c]
    t1[f"{c} (%)"] = linhas[c] / N_tot[c]
pos = dict(zip(cont["e"], cont["pos"]))
rep = cont[cont["pos"] <= TOP]["pos"].duplicated(keep=False).values
pos_txt = {e: f"{pos[e]}{'=' if r else ''}" for e, r in zip(TOPEV, rep)}
t1.index = [e if e not in pos_txt else f"{pos_txt[e]}. {e}" + (" *" if linhas.loc[e, "Total"] < 5 else "") for e in t1.index]
t1.loc["N de posts"] = [N_tot[c] if i % 2 == 0 else np.nan for c in PERFIS + ["Total"] for i in (0, 1)]
t1 = t1.reset_index().rename(columns={"index": "Evento"})
for c in [x for x in t1.columns if "(N)" in x]:
    t1[c] = t1[c].astype(int)
for c in PERFIS + ["Total"]:
    assert t1.iloc[0][f"{c} (N)"] + t1.iloc[1][f"{c} (N)"] == N_tot[c]


# ---- Por_mes
def por_mes(sub):
    b = df if sub == "Total" else df[df["Perfil"] == sub]
    e = ev_long if sub == "Total" else ev_long[ev_long["Perfil"] == sub]
    rows = []
    for d_ in sorted(df["data"].unique()):
        bm = b[b["data"] == d_]
        em = e[e["data"] == d_]["ev"].value_counts()
        n, a = len(bm), int(bm["algum"].sum())
        if len(em):
            top_n = int(em.max())
            princ = " / ".join(sorted(em[em == top_n].index))
        else:
            top_n, princ = 0, "—"
        rows.append([f"{INV[pd.Timestamp(d_).month]}/{pd.Timestamp(d_).year}", n, a, a / n if n else np.nan, princ, top_n])
    t = pd.DataFrame(rows, columns=["Mês/Ano", "N de posts", "Posts com algum evento (N)", "Posts com algum evento (%)",
                                    "Principal evento do mês", "N desse evento"])
    tot = [["Total", int(t["N de posts"].sum()), int(t["Posts com algum evento (N)"].sum()),
            t["Posts com algum evento (N)"].sum() / t["N de posts"].sum(), "", np.nan]]
    t = pd.concat([t, pd.DataFrame(tot, columns=t.columns)], ignore_index=True)
    assert t.iloc[-1]["N de posts"] == N_tot[sub]
    return t


mes = {g: por_mes(g) for g in ["Total"] + PERFIS}

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma."],
    ["Período Especial é multirrótulo", "A célula pode ter mais de um evento, separados por \";\". Cada post conta uma vez para cada evento que tem (posts, não células). \"Não há\" nunca aparece com outro rótulo. 28 posts têm mais de um evento (25 com 2 e 3 com 3); por isso a soma das linhas de eventos passa do N de \"Algum evento\"."],
    ["Critério de registro", "O contexto temporal (evento) foi registrado só quando estava presente na legenda, no texto na imagem ou na mídia do post. Posts publicados durante um evento, mas sem menção a ele, foram codificados como \"Não há\"."],
    ["Fórmula do %", "Por_perfil: % = nº de posts do perfil com o rótulo ÷ total de posts do perfil (Ego 414; Roma 1.131; IO 2.721; Total 4.266) × 100. Por_mes: % = posts com algum evento no mês ÷ posts do mês (dentro do perfil) × 100."],
    ["Eventos listados", f"Os 15 eventos mais frequentes do total geral (ranking de competição; empates na 15ª posição incluídos: {len(TOPEV)} eventos). Total de rótulos de evento distintos: {cont.shape[0]}."],
    ["Principal evento do mês", "Rótulo mais frequente no mês (exceto \"Não há\"). Empates são mostrados juntos, separados por \" / \". \"—\" = nenhum evento no mês."],
    ["Período", "Coluna Período no formato \"(mmm-aaaa)\", convertida para ordem cronológica (jan/2024 a dez/2024)."],
    ["*", "Evento com menos de 5 posts no total (nenhum caso entre os 15 listados)."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a Período Especial."],
    ["Script", "solicitacao10.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t1.to_excel(w, sheet_name="Por_perfil", index=False)
    r = 0
    for g in ["Total"] + PERFIS:
        pd.DataFrame([[f"{g} (N = {N_tot[g]:,})".replace(",", ".")]]).to_excel(w, sheet_name="Por_mes", index=False, header=False, startrow=r)
        mes[g].to_excel(w, sheet_name="Por_mes", index=False, startrow=r + 1)
        r += len(mes[g]) + 4
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
ws = wb["Por_perfil"]
for c in ws[1]:
    c.font = Font(bold=True)
for rr in (2, 3, ws.max_row):
    for c in ws[rr]:
        c.font = Font(bold=True)
for col in ws.columns:
    h = str(col[0].value)
    for c in col[1:]:
        c.number_format = "0.0%" if "(%)" in h else "#,##0" if "(N)" in h else c.number_format
    ws.column_dimensions[col[0].column_letter].width = max(12, len(h) + 2)
ws.column_dimensions["A"].width = 52
ws = wb["Por_mes"]
for row in ws.iter_rows():
    v0 = str(row[0].value)
    if row[1].value is None and row[0].value is not None or v0 == "Mês/Ano" or v0 == "Total":
        for c in row:
            c.font = Font(bold=True)
    if v0 not in ("Mês/Ano", "None") and row[1].value is not None:
        row[1].number_format = row[2].number_format = row[5].number_format = "#,##0"
        row[3].number_format = "0.0%"
for L, wd in zip("ABCDEF", (14, 12, 16, 16, 45, 12)):
    ws.column_dimensions[L].width = wd
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 32
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico: % de posts com algum evento por mês, barras agrupadas por perfil
cores = {"Ego": "#E69F00", "Roma": "#0072B2", "Imperatriz Online": "#009E73"}
hach = {"Ego": "", "Roma": "//", "Imperatriz Online": ".."}
meses_lab = mes["Total"]["Mês/Ano"].iloc[:-1].tolist()
x = np.arange(len(meses_lab))
w_ = 0.27
fig, ax = plt.subplots(figsize=(14, 6))
ymax = 0
for k, g in enumerate(PERFIS):
    v = mes[g]["Posts com algum evento (%)"].iloc[:-1].values * 100
    ymax = max(ymax, np.nanmax(v))
    xs = x + (k - 1) * w_
    ax.bar(xs, v, width=w_, color=cores[g], hatch=hach[g], edgecolor="white", linewidth=0.5,
           label=f"{g} (N = {N_tot[g]:,})".replace(",", "."))
    for xx, vv in zip(xs, v):
        ax.text(xx, vv + 0.4, f"{vv:.1f}".replace(".", ","), ha="center", va="bottom", fontsize=6.5, rotation=90)
ax.set_xticks(x)
ax.set_xticklabels(meses_lab)
ax.set_ylim(0, ymax * 1.25)
ax.set_ylabel("% dos posts do mês com algum evento")
ax.set_xlabel("Mês/Ano")
ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
ax.spines[["top", "right"]].set_visible(False)
ax.legend(frameon=False, ncol=3, loc="upper left")
fig.text(0.99, 0.005, "Valores em %; denominador = posts do perfil no mês.", ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")

pd.set_option("display.width", 250)
print(t1.to_string())
for g in ["Total"] + PERFIS:
    print(g); print(mes[g].to_string())
