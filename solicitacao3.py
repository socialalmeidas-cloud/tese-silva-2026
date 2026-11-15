"""Solicitação 3 — Função da informação (Mediação da Informação).
Uso: python3 solicitacao3.py <planilha.xlsx>
Gera Solicitacao3_Mediacao.xlsx e .png.
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
SAIDA = "Solicitacao3_Mediacao.xlsx"
PNG = "Solicitacao3_Mediacao.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
COL = "Mediação da Informação"
CATS = {"Factual", "Publicitária", "De Entretenimento", "Institucional", "Opinativa"}

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
assert not df[COL].astype(str).str.contains("; ").any()  # rótulo único
assert set(df[COL].unique()) == CATS, set(df[COL].unique())

ct = pd.crosstab(df[COL], df["Perfil"])[PERFIS]
ct["Total"] = ct.sum(axis=1)
ct = ct.reset_index().sort_values(["Total", COL], ascending=[False, True]).set_index(COL)
t = pd.DataFrame(index=ct.index)
for c in PERFIS:
    t[f"{c} (N)"] = ct[c]
    t[f"{c} (%)"] = ct[c] / N[c]
t["Total (N)"] = ct["Total"]
t["Total (%)"] = ct["Total"] / len(df)
ordem = list(t.index)
t.index = [e + (" *" if t.loc[e, "Total (N)"] < 10 else "") for e in t.index]
tot = {f"{c} (N)": N[c] for c in PERFIS} | {f"{c} (%)": 1.0 for c in PERFIS}
tot |= {"Total (N)": len(df), "Total (%)": 1.0}
t.loc["Total"] = pd.Series(tot)
t = t.reset_index().rename(columns={"index": "Mediação"})
for c in [x for x in t.columns if "(N)" in x]:
    t[c] = t[c].astype(int)
corpo = t.iloc[:-1]
for c in PERFIS + ["Total"]:
    assert corpo[f"{c} (N)"].sum() == t.iloc[-1][f"{c} (N)"]
    assert abs(corpo[f"{c} (%)"].sum() - 1) < 1e-9

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma."],
    ["Variável", "Mediação da Informação: rótulo único por post; 5 categorias. Percentuais de cada coluna somam 100%."],
    ["Fórmula do %", "% = N da categoria no perfil ÷ total de posts do perfil (Ego 414; Roma 1.131; IO 2.721; Total 4.266) × 100."],
    ["Ordenação", "Decrescente pelo Total (N)."],
    ["*", "Categoria com menos de 10 posts no total (nenhum caso nesta tabela)."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a Mediação da Informação. Lista completa: Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix."],
    ["Gráfico", "Barras empilhadas a 100%; cada barra = um perfil; segmentos na mesma ordem da tabela (de baixo para cima). Segmentos com menos de 4% têm o rótulo (com o nome da categoria) ao lado da barra; segmentos com N = 0 não recebem rótulo (Institucional no Ego)."],
    ["Script", "solicitacao3.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t.to_excel(w, sheet_name="Mediacao_por_perfil", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
ws = wb["Mediacao_por_perfil"]
for c in ws[1] + ws[ws.max_row]:
    c.font = Font(bold=True)
for col in ws.columns:
    h = str(col[0].value)
    for c in col[1:]:
        c.number_format = "0.0%" if "(%)" in h else "#,##0" if "(N)" in h else c.number_format
    ws.column_dimensions[col[0].column_letter].width = max(12, max(len(str(c.value or "")) for c in col) + 2)
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico: barras empilhadas a 100% (Okabe-Ito + hachuras para reforçar a distinção)
cores = ["#0072B2", "#E69F00", "#009E73", "#CC79A7", "#56B4E9"]
hach = ["", "//", "", "..", "xx"]
fmt = lambda v: f"{v:.1f}%".replace(".", ",")
fig, ax = plt.subplots(figsize=(11, 6.5))
x = [0, 2, 4]  # espaçamento largo para caber os rótulos externos
base = [0.0] * len(PERFIS)
pequenos = {i: [] for i in range(len(PERFIS))}
for k, cat in enumerate(ordem):
    vals = [ct.loc[cat, p] / N[p] * 100 for p in PERFIS]
    ns = [ct.loc[cat, p] for p in PERFIS]
    ax.bar(x, vals, bottom=base, width=0.8, color=cores[k], hatch=hach[k],
           edgecolor="white", linewidth=0.8, label=cat)
    for i, (v, n) in enumerate(zip(vals, ns)):
        meio = base[i] + v / 2
        txt = f"{fmt(v)} ({f'{n:,}'.replace(',', '.')})"
        if v >= 4:
            ax.text(x[i], meio, txt, ha="center", va="center", fontsize=9,
                    color="white" if k in (0, 2) else "black",
                    bbox=None if k in (0, 2) else dict(facecolor="white", edgecolor="none", alpha=0.8, pad=1))
        elif n > 0:  # segmentos vazios (N = 0) não recebem rótulo
            pequenos[i].append((meio, f"{cat}: {txt}", cores[k]))
    base = [b + v for b, v in zip(base, vals)]
# rótulos dos segmentos pequenos, à direita da barra, sem sobreposição
for i, itens in pequenos.items():
    ultimo = -10
    for meio, txt, cor in sorted(itens):
        y = max(meio, ultimo + 4)
        ax.annotate(txt, xy=(x[i] + 0.4, meio), xytext=(x[i] + 0.55, y), fontsize=8.5,
                    va="center", arrowprops=dict(arrowstyle="-", color="gray", lw=0.6))
        ultimo = y
ax.set_xticks(x)
ax.set_xticklabels([f"{p}\n(N = {N[p]:,})".replace(",", ".") for p in PERFIS])
ax.set_ylabel("% dos posts do perfil")
ax.set_ylim(0, 100)
ax.set_xlim(-0.7, x[-1] + 1.9)
ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
ax.spines[["top", "right"]].set_visible(False)
h, l = ax.get_legend_handles_labels()
ax.legend(h[::-1], l[::-1], title="Mediação da Informação", frameon=False,
          bbox_to_anchor=(1.0, 0.5), loc="center left")
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")
print(t.to_string())
