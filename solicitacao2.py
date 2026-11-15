"""Solicitação 2 — Agenda temática (Eixo Temático).
Uso: python3 solicitacao2.py <planilha.xlsx>
Gera Solicitacao2_Agenda_tematica.xlsx e .png.
"""
import sys
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao2_Agenda_tematica.xlsx"
PNG = "Solicitacao2_Agenda_tematica.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
COL = "Eixo Temático"
TOP = 10

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
assert not df[COL].astype(str).str.contains("; ").any()  # rótulo único
assert df[COL].nunique() == 34

# ---- Aba 1: Eixo por perfil
ct = pd.crosstab(df[COL], df["Perfil"])[PERFIS]
ct["Total"] = ct.sum(axis=1)
ct = ct.reset_index().sort_values(["Total", COL], ascending=[False, True]).set_index(COL)
t1 = pd.DataFrame(index=ct.index)
for c in PERFIS:
    t1[f"{c} (N)"] = ct[c]
    t1[f"{c} (%)"] = ct[c] / N[c]
t1["Total (N)"] = ct["Total"]
t1["Total (%)"] = ct["Total"] / len(df)
t1.index = [e + (" *" if t1.loc[e, "Total (N)"] < 10 else "") for e in t1.index]
tot = {f"{c} (N)": N[c] for c in PERFIS} | {f"{c} (%)": 1.0 for c in PERFIS}
tot |= {"Total (N)": len(df), "Total (%)": 1.0}
t1.loc["Total"] = pd.Series(tot)
t1 = t1.reset_index().rename(columns={"index": "Eixo"})
for c in [x for x in t1.columns if "(N)" in x]:
    t1[c] = t1[c].astype(int)
corpo = t1.iloc[:-1]
for c in PERFIS + ["Total"]:
    assert corpo[f"{c} (N)"].sum() == t1.iloc[-1][f"{c} (N)"]
    assert abs(corpo[f"{c} (%)"].sum() - 1) < 1e-9

# ---- Aba 2: Top 10 por perfil (ranking de competição; empates no corte são incluídos)
blocos = {}
for p in PERFIS:
    s = (df.loc[df["Perfil"] == p, COL].value_counts()
         .rename_axis("Eixo").reset_index(name="N")
         .sort_values(["N", "Eixo"], ascending=[False, True]))
    s["Posição"] = s["N"].rank(method="min", ascending=False).astype(int)
    s = s[s["Posição"] <= TOP].copy()
    repet = s["Posição"].duplicated(keep=False)
    s["Posição"] = s["Posição"].astype(str) + repet.map({True: "=", False: ""})
    s["% do perfil"] = s["N"] / N[p]
    s["Eixo"] = [e + (" *" if n < 10 else "") for e, n in zip(s["Eixo"], s["N"])]
    blocos[p] = s[["Posição", "Eixo", "N", "% do perfil"]].reset_index(drop=True)

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma."],
    ["Variável", "Eixo Temático: rótulo único por post; 34 eixos. Percentuais de cada coluna somam 100%."],
    ["Fórmula do %", "% = N do eixo no perfil ÷ total de posts do perfil (Ego 414; Roma 1.131; IO 2.721; Total 4.266) × 100."],
    ["Ordenação", "Eixo_por_perfil: decrescente pelo Total (N); empates em ordem alfabética."],
    ["* em Eixo_por_perfil", "Eixo com menos de 10 posts no total do corpus."],
    ["* em Top10_por_perfil", "Eixo com menos de 10 posts dentro do perfil."],
    ["Empates no Top 10", "Ranking de competição (1, 2, 2, 4...). Posição com \"=\" indica empate. Quando há empate na 10ª posição, todos os eixos empatados foram incluídos; por isso o bloco do Ego tem mais de 10 linhas (4 eixos empatados com 4 posts na 10ª posição)."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a Eixo Temático. Lista completa: Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix."],
    ["Script", "solicitacao2.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t1.to_excel(w, sheet_name="Eixo_por_perfil", index=False)
    col0 = 0
    for p in PERFIS:
        pd.DataFrame([[p]]).to_excel(w, sheet_name="Top10_por_perfil", index=False,
                                     header=False, startrow=0, startcol=col0)
        blocos[p].to_excel(w, sheet_name="Top10_por_perfil", index=False,
                           startrow=1, startcol=col0)
        col0 += 5
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
ws = wb["Eixo_por_perfil"]
for cell in ws[1]:
    cell.font = Font(bold=True)
for c in ws[ws.max_row]:
    c.font = Font(bold=True)
for col in ws.columns:
    h = str(col[0].value)
    for c in col[1:]:
        c.number_format = "0.0%" if "(%)" in h else "#,##0" if "(N)" in h else c.number_format
    ws.column_dimensions[col[0].column_letter].width = max(12, max(len(str(c.value or "")) for c in col) + 2)
ws = wb["Top10_por_perfil"]
for r in (1, 2):
    for c in ws[r]:
        c.font = Font(bold=True)
for col in ws.columns:
    h = str(col[1].value)
    for c in col[2:]:
        if h == "% do perfil" and c.value is not None:
            c.number_format = "0.0%"
        elif h == "N" and c.value is not None:
            c.number_format = "#,##0"
    ws.column_dimensions[col[0].column_letter].width = max(10, max(len(str(c.value or "")) for c in col[1:]) + 2)
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico: um painel por perfil, barras horizontais, x em %
cores = {"Ego": "#E69F00", "Roma": "#0072B2", "Imperatriz Online": "#009E73"}
altura = max(len(b) for b in blocos.values())
fig, axes = plt.subplots(1, 3, figsize=(15, 0.42 * altura + 1.6))
for ax, p in zip(axes, PERFIS):
    b = blocos[p].iloc[::-1]
    v = b["% do perfil"] * 100
    ax.barh(b["Eixo"], v, color=cores[p])
    for y, (val, n) in enumerate(zip(v, b["N"])):
        ax.text(val + 0.8, y, f"{val:.1f}%".replace(".", ",") + f" ({n:,})".replace(",", "."),
                va="center", fontsize=8)
    ax.set_xlim(0, max(bb["% do perfil"].max() for bb in blocos.values()) * 100 * 1.35)
    ax.set_xlabel(f"% dos posts do perfil (N = {N[p]:,})".replace(",", "."))
    ax.set_title(p, fontsize=11)  # identificação do painel, não título do gráfico
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="y", labelsize=9)
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:.0f}%")
fig.tight_layout()
fig.savefig(PNG, dpi=300)

for p in PERFIS:
    print(p); print(blocos[p].to_string())
print(t1.head(10).to_string())
