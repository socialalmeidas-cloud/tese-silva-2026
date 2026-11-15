"""Solicitação 20 — Eixo Temático × Mediação da Informação, sem o eixo Publicidade (gráfico).
Base: os 10 eixos mais frequentes do total geral (Solicitação 4), sem Publicidade = 9 eixos (mesma base da 4B).
Uso: python3 -I solicitacao20.py <planilha.xlsx>
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
SAIDA, PNG = "Solicitacao20_Eixo_x_Mediacao_sem_Publicidade.xlsx", "Solicitacao20_Eixo_x_Mediacao_sem_Publicidade.png"
EIXO, MED = "Eixo Temático", "Mediação da Informação"
MEDS = ["Factual", "Publicitária", "De Entretenimento", "Institucional", "Opinativa"]
PERFIS = [("Total", None), ("Ego", "Ego"), ("Roma", "Roma"), ("Imperatriz_Online", "Imperatriz Online")]

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
df[EIXO] = df[EIXO].astype(str).str.strip()
df[MED] = df[MED].astype(str).str.strip()
assert set(df[MED]) == set(MEDS)

freq = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
TOP10 = freq["e"].iloc[:10].tolist()
EIXOS = [e for e in TOP10 if e != "Publicidade"]
assert len(EIXOS) == 9
sub = df[df[EIXO].isin(EIXOS)]
assert len(sub) == 1936


def tabela(d):
    ct = pd.crosstab(d[EIXO], d[MED]).reindex(index=EIXOS, columns=MEDS, fill_value=0)
    ct.loc["Total"] = ct.sum()
    n = ct.sum(axis=1)
    out = pd.DataFrame({"Eixo": [f"{e} *" if (n[e] < 10 and e != "Total") else e for e in ct.index], "N do eixo": n.values})
    for m in MEDS:
        out[f"{m} (N)"] = ct[m].values
        out[f"{m} (%)"] = np.where(n.values > 0, ct[m].values / np.where(n.values > 0, n.values, 1), np.nan)
    assert (ct.sum(axis=1) == n).all()
    return out


tabs = {nome: tabela(sub if p is None else sub[sub["Perfil"] == p]) for nome, p in PERFIS}
assert sum(tabs[k]["N do eixo"].iloc[-1] for k in ["Ego", "Roma", "Imperatriz_Online"]) == tabs["Total"]["N do eixo"].iloc[-1] == 1936

notas = pd.DataFrame([
    ["Arquivo", f"{NOME_ORIGINAL} (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "1.936 posts dos 9 eixos (Ego " + str(tabs["Ego"]["N do eixo"].iloc[-1]) + "; Roma " + f'{tabs["Roma"]["N do eixo"].iloc[-1]:,}'.replace(",", ".")
     + "; Imperatriz Online " + f'{tabs["Imperatriz_Online"]["N do eixo"].iloc[-1]:,}'.replace(",", ".") + ")"],
    ["Exclusões", "Eixo Publicidade (1.742 posts) e os 24 eixos fora do top 10 (588 posts). Total excluído: 2.330 de 4.266."],
    ["Base", "Mesmos 9 eixos e mesma base da Solicitação 4B (tabela a), já confirmada. A Mediação é rótulo único por post; cada linha soma 100%."],
    ["Fórmula do %", "% = N da mediação no eixo ÷ N do eixo × 100."],
    ["Mediação Publicitária", "Continua aparecendo (24 posts) porque a exclusão é do EIXO Publicidade, não da mediação Publicitária: são posts de outros eixos codificados com mediação publicitária."],
    ["Asterisco (*)", "Eixos com menos de 10 posts no perfil: percentuais instáveis."],
    ["Gráfico", "Aba Total. Valores escritos dentro dos segmentos ≥ 6%; segmentos menores listados à direita da barra com a inicial da mediação. Cores Okabe-Ito com hachuras."],
    ["Teste", "Não refeito aqui: ver Solicitação 4B (versão b/c: χ² = 798,5; gl 24; V de Cramér = 0,371; p ≈ 0,0001 por Monte Carlo)."],
    ["Script", "solicitacao20.py"],
], columns=["Item", "Descrição"])

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    for nome, _ in PERFIS:
        tabs[nome].to_excel(w, sheet_name=nome, index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)
wb = load_workbook(SAIDA)
for nome, _ in PERFIS:
    ws = wb[nome]
    for c in ws[1]:
        c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for c in row[1:]:
            c.number_format = "0.0%" if "(%)" in ws.cell(1, c.column).value else "#,##0"
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 22
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 26; ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico: barras horizontais empilhadas a 100% (aba Total)
cores = ["#0072B2", "#E69F00", "#009E73", "#CC79A7", "#56B4E9"]
hach = ["", "//", "", "..", "xx"]
t = tabs["Total"].iloc[:9].iloc[::-1].reset_index(drop=True)
br = lambda v: f"{v:.1f}%".replace(".", ",")
fig, ax = plt.subplots(figsize=(11, 6))
esq = np.zeros(len(t))
for k, m in enumerate(MEDS):
    vals = t[f"{m} (%)"].values * 100
    ax.barh(range(len(t)), vals, left=esq, color=cores[k], hatch=hach[k], edgecolor="white", linewidth=0.8, label=m)
    for y, (v_, l_) in enumerate(zip(vals, esq)):
        if v_ >= 6:
            ax.text(l_ + v_ / 2, y, br(v_), ha="center", va="center", fontsize=8.5,
                    color="white" if k in (0, 2) else "black",
                    bbox=None if k in (0, 2) else dict(facecolor="white", edgecolor="none", alpha=0.8, pad=1))
    esq += vals
sig = {"Factual": "F", "Publicitária": "P", "De Entretenimento": "E", "Institucional": "I", "Opinativa": "O"}
for y, r in t.iterrows():
    peq = [f"{sig[m]} {br(r[f'{m} (%)'] * 100)}" for m in MEDS if r[f"{m} (N)"] > 0 and r[f"{m} (%)"] * 100 < 6]
    if peq:
        ax.text(101.5, y, " · ".join(peq), va="center", fontsize=8)
ax.set_yticks(range(len(t)))
ax.set_yticklabels([f"{e} (N = {n:,})".replace(",", ".") for e, n in zip(t["Eixo"], t["N do eixo"])])
ax.set_xlim(0, 100)
ax.set_xlabel("% dos posts do eixo (N = 1.936; sem o eixo Publicidade)")
ax.xaxis.set_major_formatter(lambda v_, _: f"{v_:.0f}%")
ax.spines[["top", "right"]].set_visible(False)
ax.legend(title="Mediação da Informação", frameon=False, ncol=5, loc="upper center", bbox_to_anchor=(0.5, -0.12))
fig.text(0.99, 0.01, "À direita: segmentos < 6% (F Factual · P Publicitária · E De Entretenimento · I Institucional · O Opinativa)",
         ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")

for nome, _ in PERFIS:
    print(nome); print(tabs[nome][["Eixo", "N do eixo"] + [f"{m} (%)" for m in MEDS]].round(3).to_string())
