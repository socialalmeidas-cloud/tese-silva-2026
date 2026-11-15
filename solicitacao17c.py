"""Solicitação 17C — Mapa de calor: Temáticas Secundárias × Eixo Temático.
Linhas: união das 5 temáticas mais frequentes de cada um dos 10 eixos (empates na 5ª posição entram todos).
Colunas: os 10 eixos mais frequentes. Célula: % dos posts do eixo que têm a temática.
Ordem das linhas: agrupadas pelo eixo em que a temática tem a maior %, e dentro do grupo pela % decrescente.
Mesmas regras da Solicitação 17: posts (cada rótulo uma vez por post); variantes aprovadas pela autora aplicadas (v2, 07/11/2026).
Uso: python3 -I solicitacao17c.py <planilha.xlsx>
"""
import sys
import numpy as np
import os
import pandas as pd
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # pasta do script (python -I não a inclui)
from variantes_aprovadas import rotulos_do_post
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA, PNG = "Solicitacao17C_Mapa_de_calor_Tematica_x_Eixo.xlsx", "Solicitacao17C_Mapa_de_calor_Tematica_x_Eixo.png"
COL, EIXO = "Temática Secundária ", "Eixo Temático"
K = 5
br_int = lambda v: f"{int(v):,}".replace(",", ".")

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
df[EIXO] = df[EIXO].astype(str).str.strip()
df["rot"] = df[COL].apply(rotulos_do_post)          # variantes aprovadas pela autora aplicadas (v2, 07/11/2026)
fe = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
EIXOS = fe["e"].iloc[:10].tolist()
NE = df[EIXO].value_counts()

sub = df[df[EIXO].isin(EIXOS)].explode("rot").dropna(subset=["rot"]).reset_index(drop=True)
N = pd.crosstab(sub["rot"], sub[EIXO]).reindex(columns=EIXOS, fill_value=0)   # posts do eixo com a temática

# linhas: top 5 de cada eixo, incluindo empates na 5ª posição
sel = []
for e in EIXOS:
    s = N[e].sort_values(ascending=False)
    sel += s[s >= s.iloc[K - 1]].index.tolist()
sel = list(dict.fromkeys(sel))
P = N.loc[sel] / NE[EIXOS].values                                          # % dos posts do eixo
dono = P.idxmax(axis=1)
ordem = sorted(sel, key=lambda r: (EIXOS.index(dono[r]), -P.loc[r, dono[r]], r))
N, P = N.loc[ordem], P.loc[ordem]
assert (N.values <= NE[EIXOS].values).all()

# ---- planilha
def com_total(m, fmt_n):
    out = m.copy()
    out.columns = [f"{e} (N = {br_int(NE[e])})" for e in EIXOS]
    out.insert(0, "Eixo de maior %", [dono[r] for r in m.index])
    out.index.name = "Temática"
    return out.reset_index()

tN, tP = com_total(N, True), com_total(P, False)
tN.loc[len(tN)] = ["N de posts do eixo", ""] + [int(NE[e]) for e in EIXOS]
notas = pd.DataFrame([
    ["Arquivo", f"{NOME_ORIGINAL} (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "3.678 posts dos 10 eixos mais frequentes (de 4.266). Excluídos: 588 posts dos 24 eixos fora do top 10."],
    ["Linhas", f"{len(ordem)} temáticas: união das 5 mais frequentes de cada eixo (empates na 5ª posição incluídos)."],
    ["Célula", "% = N de posts do eixo com a temática ÷ N de posts do eixo × 100. Um post pode ter vários rótulos: as colunas não somam 100%. "
     "Valores para todas as temáticas das linhas, inclusive nos eixos onde elas não estão no top 5."],
    ["Ordem", "Linhas agrupadas pelo eixo em que a temática tem a maior % (coluna \"Eixo de maior %\") e, dentro do grupo, pela % decrescente."],
    ["Contagem", "Posts, não células; cada rótulo uma vez por post. Variantes de grafia aprovadas pela autora (07/11/2026) aplicadas."],
    ["Cuidado de leitura", "Eixos pequenos (Período Eleitoral, 73; Esporte, 58) têm % altas com poucos posts: ver a aba N."],
    ["Variantes (v2)", "Versão 2 (07/11/2026): variantes de grafia aprovadas pela autora aplicadas (ver aba Variantes_aplicadas da Solicitação 17). Versão 1 salva como *_v1_sem_variantes."],
    ["Script", "solicitacao17c.py"],
], columns=["Item", "Descrição"])
with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    tP.to_excel(w, sheet_name="Percentual", index=False)
    tN.to_excel(w, sheet_name="N", index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)
wb = load_workbook(SAIDA)
for nome in ["Percentual", "N"]:
    ws = wb[nome]
    for c in ws[1]:
        c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True)
    for row in ws.iter_rows(min_row=2, min_col=3):
        for c in row:
            c.number_format = "0.0%" if nome == "Percentual" else "#,##0"
    ws.column_dimensions["A"].width = 26; ws.column_dimensions["B"].width = 18
ws = wb["Notas"]; ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 18; ws.column_dimensions["B"].width = 120
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- mapa de calor (cividis; valor escrito em toda célula > 0)
M = P.values * 100
fig, ax = plt.subplots(figsize=(13, 0.33 * len(ordem) + 2.5))
cmap = matplotlib.colormaps["cividis_r"].copy()
cmap.set_bad("white")                                   # célula 0 em branco, para o vazio não competir com os valores
im = ax.imshow(np.ma.masked_equal(M, 0), cmap=cmap, aspect="auto", vmin=0, vmax=M.max())
ax.set_facecolor("white")
for i in range(M.shape[0] + 1):
    ax.axhline(i - 0.5, color="#e6e6e6", linewidth=0.5)
for j in range(M.shape[1] + 1):
    ax.axvline(j - 0.5, color="#e6e6e6", linewidth=0.5)
for i in range(M.shape[0]):
    for j in range(M.shape[1]):
        v = M[i, j]
        if v > 0:
            ax.text(j, i, f"{v:.1f}".replace(".", ","), ha="center", va="center", fontsize=7.5,
                    color="white" if v > M.max() * 0.55 else "black")
ax.set_xticks(range(len(EIXOS)))
ax.set_xticklabels([f"{e}\n(N = {br_int(NE[e])})" for e in EIXOS], fontsize=8.5, rotation=35, ha="right")
ax.set_yticks(range(len(ordem))); ax.set_yticklabels(ordem, fontsize=8.5)
ax.xaxis.tick_top(); plt.setp(ax.get_xticklabels(), rotation=35, ha="left")
# linhas separando os grupos (eixo de maior %)
grupos = [dono[r] for r in ordem]
for i in range(1, len(grupos)):
    if grupos[i] != grupos[i - 1]:
        ax.axhline(i - 0.5, color="#555555", linewidth=1.4)
ax.set_xticks(np.arange(-0.5, len(EIXOS)), minor=True)
ax.tick_params(which="minor", length=0)
for s in ax.spines.values():
    s.set_visible(False)
cb = fig.colorbar(im, ax=ax, fraction=0.025, pad=0.02)
cb.set_label("% dos posts do eixo com a temática")
cb.ax.yaxis.set_major_formatter(lambda v_, _: f"{v_:.0f}%")
fig.text(0.99, 0.005, "Linhas: 5 temáticas mais frequentes de cada eixo · agrupadas pelo eixo de maior % · célula em branco = 0 · "
         "um post pode ter vários rótulos · variantes aprovadas aplicadas", ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(rect=(0, 0.01, 1, 1))
fig.savefig(PNG, dpi=300, bbox_inches="tight")
print(len(ordem)); print((P * 100).round(1).assign(dono=dono).to_string())
