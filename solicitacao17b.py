"""Solicitação 17B — Temáticas Secundárias × Eixo Temático (gráfico em painéis).
Mesma base e mesmas regras da Solicitação 17 (aba Por_eixo): 4.266 posts, rótulos separados por ";",
contagem de posts (cada rótulo uma vez por post); variantes aprovadas pela autora aplicadas (v2, 07/11/2026).
Uso: python3 -I solicitacao17b.py <planilha.xlsx>
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
SAIDA, PNG = "Solicitacao17B_Tematica_secundaria_x_Eixo.xlsx", "Solicitacao17B_Tematica_secundaria_x_Eixo.png"
COL, EIXO = "Temática Secundária ", "Eixo Temático"
K = 10                                  # temáticas por eixo
br_int = lambda v: f"{int(v):,}".replace(",", ".")
br_pct = lambda v: f"{v:.1f}%".replace(".", ",")

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
df[EIXO] = df[EIXO].astype(str).str.strip()
df["rot"] = df[COL].apply(rotulos_do_post)          # variantes aprovadas pela autora aplicadas (v2, 07/11/2026)

fe = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
EIXOS = fe["e"].iloc[:10].tolist()      # os mesmos 10 eixos das Solicitações 4, 6 e 17

blocos = []
for e in EIXOS:
    d = df[df[EIXO] == e]
    c = d["rot"].explode().dropna().value_counts().rename_axis("Temática").reset_index(name="N de posts")
    c = c.sort_values(["N de posts", "Temática"], ascending=[False, True]).reset_index(drop=True)
    corte = c["N de posts"].iloc[K - 1]
    t = c.head(K).copy()
    t.insert(0, "Posição", range(1, len(t) + 1))
    t.insert(0, "N do eixo", len(d))
    t.insert(0, "Eixo", e)
    t["% dos posts do eixo"] = t["N de posts"] / len(d)
    t["Temática"] = [f"{x} *" if n < 5 else x for x, n in zip(t["Temática"], t["N de posts"])]
    t["Empatados com a 10ª posição e fora do gráfico"] = int((c["N de posts"] == corte).sum() - (t["N de posts"] == corte).sum())
    blocos.append(t)
tab = pd.concat(blocos, ignore_index=True)

# Conferência com a aba Por_eixo da Solicitação 17
s17 = pd.read_excel("Solicitacao17_Tematica_secundaria.xlsx", sheet_name="Por_eixo")
assert s17["Temática"].tolist() == tab["Temática"].tolist() and s17["N de posts"].tolist() == tab["N de posts"].tolist()

notas = pd.DataFrame([
    ["Arquivo", f"{NOME_ORIGINAL} (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "3.678 posts dos 10 eixos mais frequentes (de 4.266). Excluídos: 588 posts dos 24 eixos fora do top 10."],
    ["Base", "Mesma tabela da aba Por_eixo da Solicitação 17 (conferida pelo script)."],
    ["Contagem", "Posts, não células: cada rótulo conta no máximo uma vez por post. Variantes de grafia aprovadas pela autora (07/11/2026) aplicadas."],
    ["Fórmula do %", "% = N de posts do eixo com o rótulo ÷ N de posts do eixo × 100. Um post pode ter vários rótulos: as % não somam 100%."],
    ["Asterisco (*)", "Rótulos com menos de 5 posts no eixo."],
    ["Empates", "Empates em ordem alfabética, só para fixar a posição. Rótulos empatados com a 10ª posição que ficaram de fora: "
     + "; ".join(f"{e} {n}" for e, n in tab.groupby("Eixo", sort=False)["Empatados com a 10ª posição e fora do gráfico"].first().items() if n) + "."],
    ["Gráfico", "Um painel por eixo, com escala própria em cada painel (comparar as barras dentro do painel, não entre painéis)."],
    ["Variantes (v2)", "Versão 2 (07/11/2026): variantes de grafia aprovadas pela autora aplicadas (ver aba Variantes_aplicadas da Solicitação 17). Versão 1 salva como *_v1_sem_variantes."],
    ["Script", "solicitacao17b.py"],
], columns=["Item", "Descrição"])

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    tab.to_excel(w, sheet_name="Tematica_x_Eixo", index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)
wb = load_workbook(SAIDA)
ws = wb["Tematica_x_Eixo"]
for c in ws[1]:
    c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True)
for row in ws.iter_rows(min_row=2):
    row[5].number_format = "0.0%"; row[1].number_format = "#,##0"
for col, wd in zip("ABCDEFG", [20, 10, 9, 28, 11, 12, 20]):
    ws.column_dimensions[col].width = wd
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 18; ws.column_dimensions["B"].width = 120
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico: 10 painéis (2 linhas × 5), barras horizontais, % dos posts do eixo
fig, axs = plt.subplots(2, 5, figsize=(20, 9.5))
for ax, e in zip(axs.flat, EIXOS):
    t = tab[tab["Eixo"] == e]
    pct = t["% dos posts do eixo"].values * 100
    y = np.arange(len(t))[::-1]
    ax.barh(y, pct, color="#0072B2", height=0.7)
    for yy, v in zip(y, pct):
        ax.text(v + pct.max() * 0.03, yy, br_pct(v), va="center", fontsize=8)
    ax.set_yticks(y); ax.set_yticklabels(t["Temática"], fontsize=8.5)
    ax.set_xlim(0, pct.max() * 1.35)
    ax.xaxis.set_major_formatter(lambda v_, _: f"{v_:.0f}%")
    ax.tick_params(axis="x", labelsize=8)
    fora = t["Empatados com a 10ª posição e fora do gráfico"].iloc[0]
    ax.set_title(f"{e} (N = {br_int(t['N do eixo'].iloc[0])})" + (f"\n+{fora} empatados com o 10º" if fora else ""),
                 fontsize=10, loc="left")                 # nome do painel, não título da figura
    ax.spines[["top", "right"]].set_visible(False)
fig.supxlabel("% dos posts do eixo (um post pode ter vários rótulos)", fontsize=10)
fig.text(0.99, 0.003, "Escala própria em cada painel · * menos de 5 posts · empates em ordem alfabética · variantes de grafia aprovadas aplicadas",
         ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(PNG, dpi=300, bbox_inches="tight")
print(tab.round(3).to_string())
