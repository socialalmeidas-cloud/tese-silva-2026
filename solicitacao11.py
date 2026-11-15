"""Solicitação 11 — Chamadas para ação (CTA; multirrótulo, separador ";").
Uso: python3 solicitacao11.py <planilha.xlsx>
Gera Solicitacao11_CTA.xlsx e .png.
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
SAIDA = "Solicitacao11_CTA.xlsx"
PNG = "Solicitacao11_CTA.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
COL, EIXO, MED = "CTA", "Eixo Temático", "Mediação da Informação"
NH = "Não há"
EMERG = {"Pix"}
ALGUM = "Algum CTA (qualquer rótulo diferente de Não há)"

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
N_tot = {**N.to_dict(), "Total": len(df)}

df["cta"] = df[COL].astype(str).str.split(";").apply(lambda l: sorted({x.strip() for x in l if x.strip()}))
assert df["cta"].apply(len).min() >= 1
assert not df["cta"].apply(lambda l: NH in l and len(l) > 1).any()
df["algum"] = df["cta"].apply(lambda l: l != [NH])
df["n_cta"] = df["cta"].apply(lambda l: 0 if l == [NH] else len(l))
long = df[["url ", "Perfil", EIXO, MED, "cta"]].explode("cta").reset_index(drop=True)
cont = long["cta"].value_counts().rename_axis("r").reset_index(name="n").sort_values(["n", "r"], ascending=[False, True])
ROTS = [r for r in cont["r"] if r != NH]
assert len(ROTS) + 1 == 7
rot = {r: r + (" †" if r in EMERG else "") for r in ROTS + [NH]}

# ---- Por_perfil
ct = pd.crosstab(long["cta"], long["Perfil"]).reindex(index=[NH] + ROTS, columns=PERFIS, fill_value=0)
ct.loc[ALGUM] = df.groupby("Perfil")["algum"].sum().reindex(PERFIS)
ct["Total"] = ct.sum(axis=1)
t1 = pd.DataFrame(index=ct.index)
for c in PERFIS + ["Total"]:
    t1[f"{c} (N)"] = ct[c]
    t1[f"{c} (%)"] = ct[c] / N_tot[c]
t1.index = [rot.get(r, r) + (" *" if r in ct.index and r != ALGUM and ct.loc[r, "Total"] < 10 else "") for r in t1.index]
t1.loc["N de posts"] = [N_tot[c] if i == 0 else np.nan for c in PERFIS + ["Total"] for i in (0, 1)]
t1 = t1.reset_index().rename(columns={"index": "CTA"})
for c in [x for x in t1.columns if "(N)" in x]:
    t1[c] = t1[c].astype(int)
for c in PERFIS + ["Total"]:
    assert ct.loc[NH, c] + ct.loc[ALGUM, c] == N_tot[c]


# ---- Por_eixo / Por_mediacao
def por_grupo(col, grupos):
    c = pd.crosstab(long[col], long["cta"]).reindex(index=grupos, columns=ROTS, fill_value=0)
    n = df[col].value_counts().reindex(grupos)
    a = df.groupby(col)["algum"].sum().reindex(grupos)
    out = c.copy()
    out.columns = [f"{rot[r]} (N)" for r in ROTS]
    out.insert(0, f"N do {'eixo' if col == EIXO else 'grupo'}", n.values)
    out["Algum CTA (N)"] = a.values
    out["Algum CTA (%)"] = a.values / n.values
    out.index.name = "Eixo" if col == EIXO else "Mediação"
    out = out.reset_index()
    out.iloc[:, 0] = [g + (" *" if k < 10 else "") for g, k in zip(out.iloc[:, 0], n.values)]
    return out


freq = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
EIXOS = freq["e"].iloc[:10].tolist()
t_eixo = por_grupo(EIXO, EIXOS)
MEDS = df[MED].value_counts().index.tolist()
assert len(MEDS) == 5
t_med = por_grupo(MED, MEDS).rename(columns={"N do grupo": "N da mediação"})
tot_med = {"Mediação": "Total", "N da mediação": len(df), **{f"{rot[r]} (N)": int(t_med[f"{rot[r]} (N)"].sum()) for r in ROTS},
           "Algum CTA (N)": int(df["algum"].sum()), "Algum CTA (%)": df["algum"].mean()}
t_med = pd.concat([t_med, pd.DataFrame([tot_med])], ignore_index=True)
assert t_med.iloc[:-1]["N da mediação"].sum() == len(df)

# ---- CTA_por_post
df["faixa"] = pd.cut(df["n_cta"], [-1, 0, 1, np.inf], labels=["0 (Não há)", "1 CTA", "2 ou mais CTA"])
c2 = pd.crosstab(df["faixa"], df["Perfil"]).reindex(columns=PERFIS)
c2["Total"] = c2.sum(axis=1)
t2 = pd.DataFrame(index=c2.index.astype(str))
for c in PERFIS + ["Total"]:
    t2[f"{c} (N)"] = c2[c].values
    t2[f"{c} (%)"] = c2[c].values / N_tot[c]
t2.loc["Total"] = [N_tot[c] if i == 0 else 1.0 for c in PERFIS + ["Total"] for i in (0, 1)]
t2 = t2.reset_index().rename(columns={"index": "Nº de CTA no post", "faixa": "Nº de CTA no post"})
for c in [x for x in t2.columns if "(N)" in x]:
    t2[c] = t2[c].astype(int)
for c in PERFIS + ["Total"]:
    assert t2.iloc[:-1][f"{c} (N)"].sum() == N_tot[c]
medias = pd.DataFrame({"Medida": ["Média de CTA por post (Não há = 0)", "Média de CTA entre posts com algum CTA", "Máximo de CTA em um post"],
                       **{c: [b["n_cta"].mean(), b.loc[b["algum"], "n_cta"].mean(), b["n_cta"].max()]
                          for c, b in [(p, df[df["Perfil"] == p]) for p in PERFIS] + [("Total", df)]}})
detalhe = pd.crosstab(df["n_cta"], df["Perfil"]).reindex(columns=PERFIS)
detalhe["Total"] = detalhe.sum(axis=1)
detalhe = detalhe.reset_index().rename(columns={"n_cta": "Nº exato de CTA"})

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma."],
    ["Contagem", "CTA é multirrótulo: célula separada por \";\"; cada post conta uma vez para cada rótulo que tem (posts, não células). \"Não há\" nunca aparece com outro rótulo."],
    ["Algum CTA", "Post com pelo menos um rótulo diferente de \"Não há\"; conta uma vez por post. \"Não há\" + \"Algum CTA\" = N do perfil."],
    ["Fórmula do %", "Por_perfil: % = nº de posts do perfil com o rótulo ÷ total de posts do perfil × 100. Por_eixo e Por_mediacao: % = posts do grupo com algum CTA ÷ N do grupo × 100. CTA_por_post: % = posts na faixa ÷ posts do perfil × 100."],
    ["Soma > 100%", "Como um post pode ter vários CTA, a soma dos rótulos passa do número de posts com algum CTA; a linha \"Algum CTA\" não deve ser somada às demais."],
    ["Média de CTA", "Nº de rótulos de CTA por post, contando \"Não há\" como 0."],
    ["*", "Rótulo ou grupo com menos de 10 posts."],
    ["† Categoria emergente", "Pix † surgiu por \"Outro (especificar)\" e foi aplicada apenas a partir do momento em que surgiu, de forma desigual entre lotes. Resultados indicativos: comparações entre perfis ou codificadoras podem refletir quando a categoria foi aplicada, não diferenças reais entre os perfis."],
    ["Retirada em 04/11", "Imperativo foi removido de CTA."],
    ["Lista completa de emergentes", "Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix."],
    ["Script", "solicitacao11.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t1.to_excel(w, sheet_name="Por_perfil", index=False)
    t_eixo.to_excel(w, sheet_name="Por_eixo", index=False)
    t_med.to_excel(w, sheet_name="Por_mediacao", index=False)
    t2.to_excel(w, sheet_name="CTA_por_post", index=False)
    medias.to_excel(w, sheet_name="CTA_por_post", index=False, startrow=len(t2) + 2)
    detalhe.to_excel(w, sheet_name="CTA_por_post", index=False, startrow=len(t2) + 7)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

amarelo = PatternFill("solid", fgColor="FFF2CC")
wb = load_workbook(SAIDA)
for nome in ["Por_perfil", "Por_eixo", "Por_mediacao"]:
    ws = wb[nome]
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[1].height = 32
    for col in ws.columns:
        h = str(col[0].value)
        for c in col[1:]:
            c.number_format = "0.0%" if "(%)" in h else "#,##0" if ("(N)" in h or h.startswith("N d")) else c.number_format
            if "†" in h:
                c.fill = amarelo
        if "†" in h:
            col[0].fill = amarelo
        ws.column_dimensions[col[0].column_letter].width = 14
    ws.column_dimensions["A"].width = 48 if nome == "Por_perfil" else 22
    for row in ws.iter_rows(min_row=2):
        lab = str(row[0].value)
        if "†" in lab:
            for c in row:
                c.fill = amarelo
        if lab.startswith(("Algum", "N de posts", "Total")):
            for c in row:
                c.font = Font(bold=True)
ws = wb["CTA_por_post"]
for r in (1, len(t2) + 1, len(t2) + 3, len(t2) + 8):
    for c in ws[r]:
        c.font = Font(bold=True)
for col in ws.columns:
    h = str(col[0].value)
    for c in col[1:len(t2) + 1]:
        c.number_format = "0.0%" if "(%)" in h else "#,##0" if "(N)" in h else c.number_format
    ws.column_dimensions[col[0].column_letter].width = 14
ws.column_dimensions["A"].width = 40
for rr in (len(t2) + 4, len(t2) + 5):
    for c in ws[rr][1:]:
        c.number_format = "0.00"
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 28
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
    if str(row[0].value).startswith("†"):
        for c in row:
            c.fill = amarelo
            c.font = Font(bold=True)
wb.save(SAIDA)

# ---- Gráfico: proporção de posts com algum CTA por mediação (empilhado a 100%)
g = t_med.iloc[:-1].copy()
g["base"] = g["Mediação"].str.replace(" *", "", regex=False)
g = g.sort_values("Algum CTA (%)")
sim = g["Algum CTA (%)"].values * 100
nao = 100 - sim
fig, ax = plt.subplots(figsize=(10, 4.8))
y = np.arange(len(g))
ax.barh(y, sim, color="#0072B2", label="Algum CTA", edgecolor="white")
ax.barh(y, nao, left=sim, color="#E69F00", hatch="//", edgecolor="white", label="Não há")
br = lambda v: f"{v:.1f}%".replace(".", ",")
milhar = lambda n: f"{int(n):,}".replace(",", ".")
for yy, (s_, n_, a_, tt) in enumerate(zip(sim, nao, g["Algum CTA (N)"], g["N da mediação"])):
    if s_ >= 6:
        ax.text(s_ / 2, yy, f"{br(s_)} ({milhar(a_)})",
                ha="center", va="center", color="white", fontsize=9)
    else:
        ax.text(s_ + 0.8, yy, f"{br(s_)} ({a_})", ha="left", va="center", fontsize=8.5,
                bbox=dict(facecolor="white", edgecolor="none", alpha=0.85, pad=1))
    if n_ >= 6:
        ax.text(s_ + n_ / 2, yy, br(n_), ha="center", va="center", fontsize=9,
                bbox=dict(facecolor="white", edgecolor="none", alpha=0.8, pad=1))
    else:  # segmento pequeno: rótulo à direita da barra
        ax.text(101, yy, f"Não há {br(n_)}", ha="left", va="center", fontsize=8.5)
ax.set_yticks(y)
ax.set_yticklabels([f"{m} (N = {k:,})".replace(",", ".") for m, k in zip(g["Mediação"], g["N da mediação"])])
ax.set_xlim(0, 100)
ax.set_xlabel("% dos posts da mediação")
ax.set_xticks(range(0, 101, 20))
ax.xaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
ax.spines[["top", "right"]].set_visible(False)
ax.legend(frameon=False, ncol=2, loc="upper center", bbox_to_anchor=(0.5, -0.15))
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")

pd.set_option("display.width", 260)
print(t1.to_string()); print(t_eixo.to_string()); print(t_med.to_string()); print(t2.to_string()); print(medias.to_string()); print(detalhe.to_string())
