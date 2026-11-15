"""Solicitação 5 — Regime de Visibilidade (multirrótulo, separador "; ").
Uso: python3 solicitacao5.py <planilha.xlsx>
Gera Solicitacao5_Regime.xlsx e .png.
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
SAIDA = "Solicitacao5_Regime.xlsx"
PNG = "Solicitacao5_Regime.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
COL = "Regime de Visibilidade"
EMERG = {"Criminoso", "Animal", "Bens materiais", "Candidatos", "Outro perfil"}

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}

# separar rótulos; um rótulo conta no máximo uma vez por post
df["rot"] = df[COL].astype(str).str.split(";").apply(lambda l: sorted({x.strip() for x in l if x.strip()}))
assert df["rot"].apply(len).min() >= 1
assert not df["rot"].apply(lambda l: "Não há" in l and len(l) > 1).any()
long = df[["url ", "Perfil", "rot"]].explode("rot").reset_index(drop=True)
assert long["rot"].nunique() == 13, long["rot"].nunique()

# ---- Aba 1: posts com cada rótulo
ct = pd.crosstab(long["rot"], long["Perfil"])[PERFIS]
ct["Total"] = ct.sum(axis=1)
ct = ct.reset_index().sort_values(["Total", "rot"], ascending=[False, True]).set_index("rot")
t1 = pd.DataFrame(index=ct.index)
for c in PERFIS:
    t1[f"{c} (N)"] = ct[c]
    t1[f"{c} (%)"] = ct[c] / N[c]
t1["Total (N)"] = ct["Total"]
t1["Total (%)"] = ct["Total"] / len(df)
ordem = list(t1.index)
rotulo = {r: r + (" †" if r in EMERG else "") + (" *" if ct.loc[r, "Total"] < 10 else "") for r in ordem}
t1.index = [rotulo[r] for r in ordem]
t1.loc["N de posts"] = [N["Ego"], np.nan, N["Roma"], np.nan, N["Imperatriz Online"], np.nan, len(df), np.nan]
t1 = t1.reset_index().rename(columns={"index": "Regime"})
for c in [x for x in t1.columns if "(N)" in x]:
    t1[c] = t1[c].astype(int)
# conferência: soma de rótulos = soma do nº de rótulos por post
for p in PERFIS:
    assert ct[p].sum() == df.loc[df["Perfil"] == p, "rot"].apply(len).sum()

# ---- Aba 2: nº de rótulos por post
df["n_rot"] = df["rot"].apply(len)
df["faixa"] = pd.cut(df["n_rot"], [0, 1, 2, np.inf], labels=["1 rótulo", "2 rótulos", "3 ou mais rótulos"])
ct2 = pd.crosstab(df["faixa"], df["Perfil"])[PERFIS]
ct2["Total"] = ct2.sum(axis=1)
t2 = pd.DataFrame(index=ct2.index.astype(str))
for c in PERFIS + ["Total"]:
    den = N[c] if c != "Total" else len(df)
    t2[f"{c} (N)"] = ct2[c].values
    t2[f"{c} (%)"] = ct2[c].values / den
tot = [N["Ego"], 1.0, N["Roma"], 1.0, N["Imperatriz Online"], 1.0, len(df), 1.0]
t2.loc["Total"] = tot
med = [df.loc[df["Perfil"] == p, "n_rot"].mean() for p in PERFIS] + [df["n_rot"].mean()]
maxi = [df.loc[df["Perfil"] == p, "n_rot"].max() for p in PERFIS] + [df["n_rot"].max()]
t2 = t2.reset_index().rename(columns={"index": "Nº de rótulos por post", "faixa": "Nº de rótulos por post"})
for c in [x for x in t2.columns if "(N)" in x]:
    t2[c] = t2[c].astype(int)
for c in PERFIS + ["Total"]:
    assert t2.iloc[:-1][f"{c} (N)"].sum() == t2.iloc[-1][f"{c} (N)"]
medias = pd.DataFrame({"Medida": ["Média de rótulos por post", "Máximo de rótulos em um post"],
                       **{p: [m, x] for p, m, x in zip(PERFIS + ["Total"], med, maxi)}})
detalhe = pd.crosstab(df["n_rot"], df["Perfil"])[PERFIS]
detalhe["Total"] = detalhe.sum(axis=1)
detalhe = detalhe.reset_index().rename(columns={"n_rot": "Nº exato de rótulos"})

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma. Nenhum post tem a célula vazia e nenhum tem \"Não há\" em Regime de Visibilidade."],
    ["Contagem", "Multirrótulo: a célula foi separada por \"; \" e cada post conta uma vez para cada rótulo que tem (posts, não células)."],
    ["Fórmula do %", "% = nº de posts do perfil com o rótulo ÷ total de posts do perfil (Ego 414; Roma 1.131; IO 2.721; Total 4.266) × 100."],
    ["Soma > 100%", "Como um post pode ter vários rótulos, a soma dos percentuais de cada coluna passa de 100% em Regime_por_perfil. Em Rotulos_por_post, cada coluna soma 100%."],
    ["*", "Rótulo com menos de 10 posts no total."],
    ["† Categorias emergentes", "Criminoso †, Animal †, Bens materiais †, Candidatos †, Outro perfil † surgiram por \"Outro (especificar)\" e foram aplicadas apenas a partir do momento em que surgiram, de forma desigual entre lotes. Em Criminoso, a proporção de posts marcados entre os de Criminalidade e Segurança Pública variou de 0% a 71% entre lotes. Resultados indicativos: comparações entre perfis ou codificadoras podem refletir quando a categoria foi aplicada, não diferenças reais entre os perfis."],
    ["Lista completa de emergentes", "Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix."],
    ["Script", "solicitacao5.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t1.to_excel(w, sheet_name="Regime_por_perfil", index=False)
    t2.to_excel(w, sheet_name="Rotulos_por_post", index=False)
    medias.to_excel(w, sheet_name="Rotulos_por_post", index=False, startrow=len(t2) + 2)
    detalhe.to_excel(w, sheet_name="Rotulos_por_post", index=False, startrow=len(t2) + 6)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
ws = wb["Regime_por_perfil"]
for c in ws[1] + ws[ws.max_row]:
    c.font = Font(bold=True)
for col in ws.columns:
    h = str(col[0].value)
    for c in col[1:]:
        c.number_format = "0.0%" if "(%)" in h else "#,##0" if "(N)" in h else c.number_format
    ws.column_dimensions[col[0].column_letter].width = max(12, len(h) + 2)
ws.column_dimensions["A"].width = 42
ws = wb["Rotulos_por_post"]
for col in ws.columns:
    h = str(col[0].value)
    for c in col[1:len(t2) + 1]:
        c.number_format = "0.0%" if "(%)" in h else "#,##0" if "(N)" in h else c.number_format
    ws.column_dimensions[col[0].column_letter].width = max(12, len(h) + 2)
ws.column_dimensions["A"].width = 30
for r in (1, len(t2) + 1, len(t2) + 3, len(t2) + 7):
    for c in ws[r]:
        c.font = Font(bold=True)
for c in ws[len(t2) + 4][1:]:
    c.number_format = "0.00"
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 28
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico: barras horizontais agrupadas
cores = {"Ego": "#E69F00", "Roma": "#0072B2", "Imperatriz Online": "#009E73"}
hach = {"Ego": "", "Roma": "//", "Imperatriz Online": ".."}
ords = ordem[::-1]
y = np.arange(len(ords))
h = 0.27
fig, ax = plt.subplots(figsize=(11, 10))
for k, p in enumerate(PERFIS):
    vals = np.array([ct.loc[r, p] / N[p] * 100 for r in ords])
    pos = y + (1 - k) * h
    ax.barh(pos, vals, height=h, color=cores[p], hatch=hach[p], edgecolor="white", linewidth=0.5,
            label=f"{p} (N = {N[p]:,})".replace(",", "."))
    for yy, v in zip(pos, vals):
        ax.text(v + 0.5, yy, f"{v:.1f}%".replace(".", ","), va="center", fontsize=7)
ax.set_yticks(y)
ax.set_yticklabels([rotulo[r] for r in ords], fontsize=9)
ax.set_xlabel("% dos posts do perfil (um post pode ter mais de um rótulo)")
ax.xaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
ax.set_xlim(0, max(ct.loc[r, p] / N[p] * 100 for r in ords for p in PERFIS) * 1.12)
ax.spines[["top", "right"]].set_visible(False)
ax.legend(frameon=False, loc="lower right")
fig.text(0.99, 0.005, "† categoria emergente (resultados indicativos) · * menos de 10 posts no total",
         ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")

pd.set_option("display.width", 250)
print(t1.to_string()); print(t2.to_string()); print(medias.to_string()); print(detalhe.to_string())
