"""Solicitação 21 — Recursos Audiovisuais × Tipo de Conteúdo (tabela e gráfico).
Uso: python3 -I solicitacao21.py <planilha.xlsx>
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
SAIDA, PNG = "Solicitacao21_Recursos_x_Tipo_de_conteudo.xlsx", "Solicitacao21_Recursos_x_Tipo_de_conteudo.png"
REC, TIPO = "Recursos Audiovisuais", "Tipo de Conteúdo"
TIPOS = ["Vídeo", "Image", "Sidecar"]
ROT_TIPO = {"Vídeo": "Vídeo", "Image": "Image (imagem única)", "Sidecar": "Sidecar (carrossel)"}
EMERG = {"Vídeo sem som", "Desenho"}
PERFIS = [("Total", None), ("Ego", "Ego"), ("Roma", "Roma"), ("Imperatriz_Online", "Imperatriz Online")]

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
df[REC] = df[REC].astype(str).str.strip()
df[TIPO] = df[TIPO].astype(str).str.strip()
assert not df[REC].str.contains(";").any()          # rótulo único por post
assert set(df[TIPO]) == set(TIPOS)
assert not df[REC].str.lower().str.startswith("não analisado").any()

ORDEM = df[REC].value_counts().rename_axis("r").reset_index(name="n").sort_values(["n", "r"], ascending=[False, True])["r"].tolist()


def marca(r, n):
    s = r + (" †" if r in EMERG else "")
    return s + (" *" if n < 10 else "")


def tabela(d):
    ct = pd.crosstab(d[REC], d[TIPO]).reindex(index=ORDEM, columns=TIPOS, fill_value=0)
    ct.loc["Total"] = ct.sum()
    n = ct.sum(axis=1)
    out = pd.DataFrame({"Recurso Audiovisual": [r if r == "Total" else marca(r, n[r]) for r in ct.index],
                        "N do recurso": n.values})
    for t in TIPOS:
        out[f"{ROT_TIPO[t]} (N)"] = ct[t].values
    for t in TIPOS:   # % de linha: composição de cada recurso por tipo
        out[f"{ROT_TIPO[t]} (% da linha)"] = np.where(n.values > 0, ct[t].values / np.where(n.values > 0, n.values, 1), np.nan)
    tot = ct.loc["Total"]
    for t in TIPOS:   # % de coluna: composição de cada tipo por recurso
        out[f"{ROT_TIPO[t]} (% da coluna)"] = ct[t].values / tot[t] if tot[t] else np.nan
    out["% do total de posts"] = n.values / n["Total"]
    return out


tabs = {nome: tabela(df if p is None else df[df["Perfil"] == p]) for nome, p in PERFIS}
assert tabs["Total"]["N do recurso"].iloc[-1] == 4266
assert sum(tabs[k]["N do recurso"].iloc[-1] for k in ["Ego", "Roma", "Imperatriz_Online"]) == 4266

notas = pd.DataFrame([
    ["Arquivo", f"{NOME_ORIGINAL} (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)."],
    ["Exclusões", "Nenhuma: as duas colunas estão preenchidas em todos os posts e não há \"Não analisado\"."],
    ["Variáveis", "Recursos Audiovisuais (codificação da pesquisa) e Tipo de Conteúdo (metadado do programa de coleta). Ambas têm um rótulo por post."],
    ["Nomenclatura", "Tipo de Conteúdo mantém a nomenclatura do programa de coleta: Vídeo, Image (imagem única), Sidecar (carrossel), como na Solicitação 1."],
    ["Percentuais", "% da linha = N do tipo no recurso ÷ N do recurso (cada linha soma 100%). % da coluna = N do recurso no tipo ÷ N do tipo (cada coluna soma 100%)."],
    ["Asterisco (*)", "Recursos com menos de 10 posts na aba: percentuais instáveis."],
    ["† Categorias emergentes", "Vídeo sem som † e Desenho † surgiram por \"Outro (especificar)\" e foram aplicadas apenas a partir do momento em que surgiram; resultados indicativos."],
    ["Combinações a conferir", "Há combinações que podem ser legítimas ou erro de codificação: Vídeo com Tipo Image (1), Fotografia com som com Tipo Vídeo (7), Vídeo/Vídeo com texto em Sidecar (10; o carrossel pode conter vídeo). Nada foi alterado."],
    ["Gráfico", "Aba Total, em números absolutos (N de posts). Valores dentro dos segmentos com ≥ 40 posts; o total de cada recurso à direita da barra."],
    ["Script", "solicitacao21.py"],
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
            c.number_format = "0.0%" if "%" in ws.cell(1, c.column).value else "#,##0"
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 34
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 26; ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico: barras horizontais empilhadas (N absoluto), aba Total
cores = ["#0072B2", "#E69F00", "#009E73"]
hach = ["", "//", ".."]
t = tabs["Total"].iloc[:-1].iloc[::-1].reset_index(drop=True)
fmt = lambda v: f"{int(v):,}".replace(",", ".")
fig, ax = plt.subplots(figsize=(11, 6.5))
esq = np.zeros(len(t))
for k, tp in enumerate(TIPOS):
    vals = t[f"{ROT_TIPO[tp]} (N)"].values.astype(float)
    ax.barh(range(len(t)), vals, left=esq, color=cores[k], hatch=hach[k], edgecolor="white", linewidth=0.8, label=ROT_TIPO[tp])
    for y, (v_, l_) in enumerate(zip(vals, esq)):
        if v_ >= 40:
            ax.text(l_ + v_ / 2, y, fmt(v_), ha="center", va="center", fontsize=8.5,
                    color="white" if k == 0 else "black",
                    bbox=None if k == 0 else dict(facecolor="white", edgecolor="none", alpha=0.8, pad=1))
    esq += vals
for y, n in enumerate(t["N do recurso"]):
    ax.text(n + 12, y, f"N = {fmt(n)}", va="center", fontsize=8.5)
ax.set_yticks(range(len(t)))
ax.set_yticklabels(t["Recurso Audiovisual"])
ax.set_xlim(0, t["N do recurso"].max() * 1.12)
ax.set_xlabel("Nº de posts (N = 4.266)")
ax.xaxis.set_major_formatter(lambda v_, _: fmt(v_))
ax.spines[["top", "right"]].set_visible(False)
ax.legend(title="Tipo de Conteúdo (programa de coleta)", frameon=False, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.11))
fig.text(0.99, 0.01, "Valores dentro dos segmentos com ≥ 40 posts · † categoria emergente (indicativo) · * menos de 10 posts",
         ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")

for nome, _ in PERFIS:
    print(nome)
    print(tabs[nome][["Recurso Audiovisual", "N do recurso"] + [f"{ROT_TIPO[x]} (N)" for x in TIPOS] + [f"{ROT_TIPO[x]} (% da linha)" for x in TIPOS]].round(3).to_string())
