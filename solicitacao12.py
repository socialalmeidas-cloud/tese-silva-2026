"""Solicitação 12 — Collab (rótulo único: Sim/Não).
Uso: python3 solicitacao12.py <planilha.xlsx>
Gera Solicitacao12_Collab.xlsx e .png.
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
SAIDA = "Solicitacao12_Collab.xlsx"
PNG = "Solicitacao12_Collab.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
EIXO = "Eixo Temático"
MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
INV = {v: k for k, v in MESES.items()}

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
N_tot = {**N.to_dict(), "Total": len(df)}
assert set(df["Collab"].astype(str).str.strip()) == {"Sim", "Não"}
df["sim"] = df["Collab"].astype(str).str.strip() == "Sim"
p = df["Período"].astype(str).str.strip().str.strip("()").str.split("-", expand=True)
df["data"] = pd.to_datetime(dict(year=p[1].astype(int), month=p[0].str.lower().map(MESES), day=1))


def linha(nome, b):
    return [nome, len(b), int(b["sim"].sum()), b["sim"].mean() if len(b) else np.nan]


COLS = ["N de posts", "Collab Sim (N)", "Collab Sim (%)"]
t_perf = pd.DataFrame([linha(p_, df[df["Perfil"] == p_]) for p_ in PERFIS] + [linha("Total", df)], columns=["Perfil"] + COLS)
assert t_perf.iloc[:-1]["Collab Sim (N)"].sum() == t_perf.iloc[-1]["Collab Sim (N)"] == 550

freq = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
EIXOS = freq["e"].iloc[:10].tolist()
outros = df[~df[EIXO].isin(EIXOS)]
t_eixo = pd.DataFrame([linha(e, df[df[EIXO] == e]) for e in EIXOS] + [linha("Outros eixos (24)", outros), linha("Total", df)],
                      columns=["Eixo", "N do eixo", "Collab Sim (N)", "Collab Sim (%)"])
assert t_eixo.iloc[:-1]["N do eixo"].sum() == len(df)


def por_mes(sub):
    b = df if sub == "Total" else df[df["Perfil"] == sub]
    rows = [linha(f"{INV[pd.Timestamp(d_).month]}/{pd.Timestamp(d_).year}", b[b["data"] == d_]) for d_ in sorted(df["data"].unique())]
    rows.append(linha("Total", b))
    t = pd.DataFrame(rows, columns=["Mês/Ano"] + COLS)
    assert t.iloc[:-1]["N de posts"].sum() == N_tot[sub]
    return t


mes = {g: por_mes(g) for g in ["Total"] + PERFIS}

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma."],
    ["Variável", "Collab: rótulo único por post (Sim/Não). Collab Sim + Collab Não = N do grupo."],
    ["Fórmula do %", "% = posts com Collab = Sim no grupo ÷ N de posts do grupo (perfil, eixo ou mês dentro do perfil) × 100."],
    ["Eixos", "Os 10 eixos mais frequentes do total geral (mesmos da Solicitação 4); acrescentadas as linhas \"Outros eixos (24)\" e \"Total\" para fechar o N."],
    ["Período", "Coluna Período no formato \"(mmm-aaaa)\", convertida para ordem cronológica (jan/2024 a dez/2024)."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a Collab."],
    ["Script", "solicitacao12.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t_perf.to_excel(w, sheet_name="Por_perfil", index=False)
    t_eixo.to_excel(w, sheet_name="Por_eixo", index=False)
    r = 0
    for g in ["Total"] + PERFIS:
        pd.DataFrame([[f"{g} (N = {N_tot[g]:,})".replace(",", ".")]]).to_excel(w, sheet_name="Por_mes", index=False, header=False, startrow=r)
        mes[g].to_excel(w, sheet_name="Por_mes", index=False, startrow=r + 1)
        r += len(mes[g]) + 4
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
for nome in ["Por_perfil", "Por_eixo", "Por_mes"]:
    ws = wb[nome]
    for row in ws.iter_rows():
        v0 = str(row[0].value)
        if v0 in ("Perfil", "Eixo", "Mês/Ano", "Total") or (nome == "Por_mes" and row[1].value is None and row[0].value is not None):
            for c in row:
                c.font = Font(bold=True)
        if len(row) >= 4 and isinstance(row[1].value, (int, float)):
            row[1].number_format = row[2].number_format = "#,##0"
            row[3].number_format = "0.0%"
    for L, wd in zip("ABCD", (24, 14, 16, 16)):
        ws.column_dimensions[L].width = wd
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico de linhas
estilo = {"Ego": ("#E69F00", "o", "-"), "Roma": ("#0072B2", "s", "--"), "Imperatriz Online": ("#009E73", "^", "-.")}
lab = mes["Total"]["Mês/Ano"].iloc[:-1].tolist()
x = np.arange(len(lab))
vals = {g: mes[g]["Collab Sim (%)"].iloc[:-1].values * 100 for g in PERFIS}
fig, ax = plt.subplots(figsize=(11, 5.5))
for g in PERFIS:
    cor, mk, ls = estilo[g]
    ax.plot(x, vals[g], color=cor, marker=mk, linestyle=ls, linewidth=2, label=f"{g} (N = {N_tot[g]:,})".replace(",", "."))
for i in x:
    ordem = sorted(PERFIS, key=lambda g: vals[g][i])
    ult = -99
    for g in ordem:  # rótulos empilhados para não sobrepor
        v = vals[g][i]
        y = max(v, ult + 2.6)
        ax.annotate(f"{v:.1f}".replace(".", ","), (i, v), xytext=(i + 0.14, y), fontsize=7.5, color=estilo[g][0], va="center",
                    bbox=dict(facecolor="white", edgecolor="none", alpha=0.9, pad=0.6), zorder=5)
        ult = y
ax.set_xticks(x)
ax.set_xticklabels(lab, rotation=45, ha="right")
ax.set_xlabel("Mês/Ano")
ax.set_ylabel("% dos posts do mês com Collab = Sim")
ax.set_ylim(0, max(v.max() for v in vals.values()) * 1.2)
ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
ax.spines[["top", "right"]].set_visible(False)
ax.grid(axis="y", alpha=0.3)
ax.legend(frameon=False, ncol=3, loc="upper left")
fig.text(0.99, 0.005, "Valores em %; denominador = posts do perfil no mês.", ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")

pd.set_option("display.width", 200)
print(t_perf.to_string()); print(t_eixo.to_string())
for g in ["Total"] + PERFIS:
    print(g); print(mes[g].to_string())
