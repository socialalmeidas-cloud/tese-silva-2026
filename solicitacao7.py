"""Solicitação 7 — Conteúdo Sensível (multirrótulo, separador "; ").
Uso: python3 solicitacao7.py <planilha.xlsx>
Gera Solicitacao7_Conteudo_sensivel.xlsx e .png.
"""
import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao7_Conteudo_sensivel.xlsx"
PNG = "Solicitacao7_Conteudo_sensivel.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
COL, EIXO = "Conteúdo Sensível", "Eixo Temático"
EMERG = {"Procedimento", "Animal morto", "Restrição de idade", "Suicídio", "Nudez"}
ALGUM = "Algum conteúdo sensível (qualquer rótulo diferente de Não há)"

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}

df["rot"] = df[COL].astype(str).str.split(";").apply(lambda l: sorted({x.strip() for x in l if x.strip()}))
assert df["rot"].apply(len).min() >= 1
assert not df["rot"].apply(lambda l: "Não há" in l and len(l) > 1).any()
df["algum"] = df["rot"].apply(lambda l: l != ["Não há"])
long = df[["url ", "Perfil", EIXO, "rot"]].explode("rot").reset_index(drop=True)
cont = long["rot"].value_counts().rename_axis("r").reset_index(name="n").sort_values(["n", "r"], ascending=[False, True])
SENS = [r for r in cont["r"] if r != "Não há"]
assert len(SENS) + 1 == 10
ORDEM = ["Não há"] + SENS
rot = {r: r + (" †" if r in EMERG else "") + (" *" if cont.set_index("r").loc[r, "n"] < 10 else "") for r in ORDEM}

# ---- Aba Por_perfil
ct = pd.crosstab(long["rot"], long["Perfil"]).reindex(index=ORDEM, columns=PERFIS, fill_value=0)
ct.loc[ALGUM] = df.groupby("Perfil")["algum"].sum().reindex(PERFIS)
ct["Total"] = ct.sum(axis=1)
t1 = pd.DataFrame(index=ct.index)
for c in PERFIS:
    t1[f"{c} (N)"] = ct[c]
    t1[f"{c} (%)"] = ct[c] / N[c]
t1["Total (N)"] = ct["Total"]
t1["Total (%)"] = ct["Total"] / len(df)
t1.index = [rot.get(r, r) for r in t1.index]
t1.loc["N de posts"] = [N["Ego"], np.nan, N["Roma"], np.nan, N["Imperatriz Online"], np.nan, len(df), np.nan]
t1 = t1.reset_index().rename(columns={"index": "Conteúdo Sensível"})
for c in [x for x in t1.columns if "(N)" in x]:
    t1[c] = t1[c].astype(int)
# conferência: Não há + Algum = N do perfil
for c in PERFIS + ["Total"]:
    assert ct.loc["Não há", c] + ct.loc[ALGUM, c] == (N[c] if c != "Total" else len(df))

# ---- Aba Por_eixo
freq = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
EIXOS = freq["e"].iloc[:10].tolist()
n_eixo = df[EIXO].value_counts().reindex(EIXOS)
ce = pd.crosstab(long[EIXO], long["rot"]).reindex(index=EIXOS, columns=SENS, fill_value=0)
t2 = ce.copy()
t2.columns = [f"{rot[r]} (N)" for r in SENS]
t2.insert(0, "N do eixo", n_eixo.values)
alg = df.groupby(EIXO)["algum"].sum().reindex(EIXOS)
t2["Algum conteúdo sensível (N)"] = alg.values
t2["Algum conteúdo sensível (%)"] = alg.values / n_eixo.values
t2.index.name = "Eixo"
t2 = t2.reset_index()
outros = df[~df[EIXO].isin(EIXOS)]
linha_outros = {"Eixo": "Outros eixos (24)", "N do eixo": len(outros),
                **{f"{rot[r]} (N)": int(outros["rot"].apply(lambda l: r in l).sum()) for r in SENS},
                "Algum conteúdo sensível (N)": int(outros["algum"].sum()),
                "Algum conteúdo sensível (%)": outros["algum"].mean()}
linha_tot = {"Eixo": "Total", "N do eixo": len(df),
             **{f"{rot[r]} (N)": int(df["rot"].apply(lambda l: r in l).sum()) for r in SENS},
             "Algum conteúdo sensível (N)": int(df["algum"].sum()),
             "Algum conteúdo sensível (%)": df["algum"].mean()}
t2 = pd.concat([t2, pd.DataFrame([linha_outros, linha_tot])], ignore_index=True)
assert t2.iloc[:-1]["N do eixo"].sum() == len(df)
assert t2.iloc[:-1]["Algum conteúdo sensível (N)"].sum() == df["algum"].sum()

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma. Em Por_eixo, os 24 eixos fora do top 10 foram agrupados na linha \"Outros eixos (24)\" (588 posts), acrescentada para fechar o total."],
    ["Contagem", "Multirrótulo: a célula foi separada por \"; \" e cada post conta uma vez para cada rótulo (posts, não células). \"Não há\" nunca aparece com outro rótulo."],
    ["Algum conteúdo sensível", "Post com pelo menos um rótulo diferente de \"Não há\". Cada post conta uma vez, mesmo que tenha vários rótulos sensíveis. \"Não há\" + \"Algum conteúdo sensível\" = N do perfil."],
    ["Fórmula do %", "Por_perfil: % = nº de posts do perfil com o rótulo ÷ total de posts do perfil × 100. Por_eixo: % = nº de posts do eixo com algum conteúdo sensível ÷ N do eixo × 100."],
    ["Soma > 100%", "Em Por_perfil, a soma dos rótulos sensíveis passa do número de posts com algum conteúdo sensível porque 40 posts têm mais de um rótulo; a linha \"Algum conteúdo sensível\" não deve ser somada às demais."],
    ["*", "Rótulo com menos de 10 posts no total."],
    ["† CATEGORIAS EMERGENTES", "Procedimento †, Animal morto †, Restrição de idade †, Suicídio †, Nudez † (destacadas em amarelo). Surgiram por \"Outro (especificar)\" e foram aplicadas apenas a partir do momento em que surgiram, de forma desigual entre lotes. Resultados indicativos: comparações entre perfis ou codificadoras podem refletir quando a categoria foi aplicada, não diferenças reais entre os perfis."],
    ["Retiradas em 04/11", "Incêndio foi retirado de Conteúdo Sensível e reclassificado como Acidente (o eixo Incêndio permanece)."],
    ["Lista completa de emergentes", "Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix."],
    ["Script", "solicitacao7.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t1.to_excel(w, sheet_name="Por_perfil", index=False)
    t2.to_excel(w, sheet_name="Por_eixo", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

amarelo = PatternFill("solid", fgColor="FFF2CC")
wb = load_workbook(SAIDA)
ws = wb["Por_perfil"]
for c in ws[1]:
    c.font = Font(bold=True)
for row in ws.iter_rows(min_row=2):
    lab = str(row[0].value)
    if "†" in lab:
        for c in row:
            c.fill = amarelo
    if lab.startswith("Algum") or lab == "N de posts":
        for c in row:
            c.font = Font(bold=True)
for col in ws.columns:
    h = str(col[0].value)
    for c in col[1:]:
        c.number_format = "0.0%" if "(%)" in h else "#,##0" if "(N)" in h else c.number_format
    ws.column_dimensions[col[0].column_letter].width = max(12, len(h) + 2)
ws.column_dimensions["A"].width = 60
ws = wb["Por_eixo"]
for c in ws[1]:
    c.font = Font(bold=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[1].height = 45
for c in ws[ws.max_row]:
    c.font = Font(bold=True)
for col in ws.columns:
    h = str(col[0].value)
    for c in col[1:]:
        c.number_format = "0.0%" if "(%)" in h else "#,##0"
        if "†" in h:
            c.fill = amarelo
    if "†" in h:
        col[0].fill = amarelo
    ws.column_dimensions[col[0].column_letter].width = 14
ws.column_dimensions["A"].width = 20
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
    if row[0].value == "† CATEGORIAS EMERGENTES":
        for c in row:
            c.fill = amarelo
            c.font = Font(bold=True)
wb.save(SAIDA)

# ---- Gráfico: % de posts com algum conteúdo sensível por eixo (10 eixos), ordenado
g = t2.iloc[:10].sort_values("Algum conteúdo sensível (%)")
fig, ax = plt.subplots(figsize=(9, 5.5))
vals = g["Algum conteúdo sensível (%)"].values * 100
ax.barh(range(len(g)), vals, color="#0072B2")
for y, (v, n, ne) in enumerate(zip(vals, g["Algum conteúdo sensível (N)"], g["N do eixo"])):
    ax.text(v + 0.6, y, f"{v:.1f}%".replace(".", ",") + f" ({n} de {ne:,})".replace(",", "."), va="center", fontsize=8.5)
ax.set_yticks(range(len(g)))
ax.set_yticklabels(g["Eixo"])
ax.set_xlim(0, vals.max() * 1.3)
ax.set_xlabel("% dos posts do eixo com algum conteúdo sensível")
ax.xaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
ax.spines[["top", "right"]].set_visible(False)
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")

pd.set_option("display.width", 300)
print(t1.to_string()); print(t2.to_string())
