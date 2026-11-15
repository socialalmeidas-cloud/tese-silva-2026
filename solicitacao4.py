"""Solicitação 4 — Eixo Temático × Mediação da Informação (percentual de linha).
Uso: python3 solicitacao4.py <planilha.xlsx>
Gera Solicitacao4_Eixo_x_Mediacao.xlsx e .png.
"""
import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.stats import chi2_contingency
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao4_Eixo_x_Mediacao.xlsx"
PNG = "Solicitacao4_Eixo_x_Mediacao.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
EIXO, MED = "Eixo Temático", "Mediação da Informação"
MEDS = ["Factual", "Publicitária", "De Entretenimento", "Institucional", "Opinativa"]

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
for c in (EIXO, MED):
    assert not df[c].astype(str).str.contains("; ").any()
assert set(df[MED]) == set(MEDS)

# 10 eixos mais frequentes do total geral (empate desfeito por ordem alfabética; não há empate no corte)
freq = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
assert freq["n"].iloc[9] > freq["n"].iloc[10]  # sem empate na 10ª posição
TOP = freq["e"].iloc[:10].tolist()
OUTROS = sorted(freq["e"].iloc[10:].tolist())


def tabela(d):
    ct = pd.crosstab(d[EIXO], d[MED]).reindex(columns=MEDS, fill_value=0)
    ct = ct.reindex(TOP + OUTROS, fill_value=0)
    linhas = ct.loc[TOP].copy()
    linhas.loc["Outros eixos"] = ct.loc[OUTROS].sum()
    linhas.loc["Total"] = ct.sum()
    out = pd.DataFrame(index=linhas.index)
    out["N do eixo"] = linhas.sum(axis=1)
    for m in MEDS:
        out[f"{m} (N)"] = linhas[m]
        out[f"{m} (%)"] = np.where(out["N do eixo"] > 0, linhas[m] / out["N do eixo"].replace(0, np.nan), np.nan)
    rot = []
    for e in out.index:
        s = e
        if e not in ("Total",) and out.loc[e, "N do eixo"] < 10:
            s += " *"
        rot.append(s)
    out.index = rot
    out = out.reset_index().rename(columns={"index": "Eixo"})
    # conferências
    assert out.iloc[:-1]["N do eixo"].sum() == out.iloc[-1]["N do eixo"] == len(d)
    for m in MEDS:
        assert out.iloc[:-1][f"{m} (N)"].sum() == out.iloc[-1][f"{m} (N)"]
    pcts = out[[f"{m} (%)" for m in MEDS]].sum(axis=1)
    assert all(abs(p - 1) < 1e-9 for p, n in zip(pcts, out["N do eixo"]) if n > 0)
    return out, linhas.loc[TOP]


tabs = {"Total": tabela(df)}
for p in PERFIS:
    tabs[p.replace(" ", "_")] = tabela(df[df["Perfil"] == p])

# ---- Teste qui-quadrado (aba Total, 10 eixos × 5 mediações, sem "Outros eixos")
obs = tabs["Total"][1].values
chi2, pval, gl, esp = chi2_contingency(obs, correction=False)
n_teste = obs.sum()
k = min(obs.shape) - 1
v = np.sqrt(chi2 / (n_teste * k))
cel_menor5 = (esp < 5).sum()
prop_menor5 = cel_menor5 / esp.size
cel_menor1 = (esp < 1).sum()
confiavel = prop_menor5 <= 0.20 and cel_menor1 == 0
esperadas = pd.DataFrame(esp, index=TOP, columns=MEDS).reset_index().rename(columns={"index": "Eixo"})

teste = pd.DataFrame([
    ["Tabela testada", "10 eixos mais frequentes × 5 categorias de Mediação (aba Total, sem \"Outros eixos\")"],
    ["N do teste", int(n_teste)],
    ["Qui-quadrado (χ²)", chi2],
    ["Graus de liberdade", int(gl)],
    ["p-valor", pval if pval > 0 else "p < 0,001 (valor menor que a precisão numérica do cálculo; o programa retorna 0)"],
    ["V de Cramér", v],
    ["Células com frequência esperada < 5", f"{cel_menor5} de {esp.size}"],
    ["Proporção de células com esperada < 5", prop_menor5],
    ["Células com frequência esperada < 1", f"{cel_menor1} de {esp.size}"],
    ["Avaliação", "Teste NÃO confiável: mais de 20% das células com frequência esperada < 5." if prop_menor5 > 0.20
     else ("Atenção: há células com esperada < 1." if cel_menor1 else "Pressuposto de frequências esperadas atendido.")],
    ["Observação", "Teste de Pearson sem correção de continuidade (scipy.stats.chi2_contingency, correction=False). V de Cramér = √(χ² / (N × (min(linhas, colunas) − 1))). A referência metodológica precisa ser verificada pela autora."],
], columns=["Item", "Valor"])

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma nas abas de distribuição. No teste, \"Outros eixos\" foi excluído (N do teste informado na aba Teste)."],
    ["Fórmula do %", "Percentual de linha: % = N da mediação no eixo ÷ N do eixo (dentro da aba) × 100. Cada linha soma 100%."],
    ["10 eixos (mesmos em todas as abas)", "; ".join(TOP) + ". Definidos pelo total geral (4.266 posts)."],
    ["Outros eixos (24)", "; ".join(OUTROS)],
    ["Células vazias de %", "Quando o eixo tem N = 0 no perfil, o % fica em branco (divisão por zero)."],
    ["*", "Eixo com menos de 10 posts dentro da aba (o % dessa linha se apoia em poucos casos)."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a Eixo Temático ou Mediação. Lista completa: Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix."],
    ["Teste", "Ver aba Teste; frequências esperadas na aba Esperadas_Teste."],
    ["Script", "solicitacao4.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    for nome in ["Total", "Ego", "Roma", "Imperatriz_Online"]:
        tabs[nome][0].to_excel(w, sheet_name=nome, index=False)
    teste.to_excel(w, sheet_name="Teste", index=False)
    esperadas.to_excel(w, sheet_name="Esperadas_Teste", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
for nome in ["Total", "Ego", "Roma", "Imperatriz_Online"]:
    ws = wb[nome]
    for c in ws[1] + ws[ws.max_row]:
        c.font = Font(bold=True)
    for col in ws.columns:
        h = str(col[0].value)
        for c in col[1:]:
            if "(%)" in h:
                c.number_format = "0.0%"
            elif "(N)" in h or h == "N do eixo":
                c.number_format = "#,##0"
        ws.column_dimensions[col[0].column_letter].width = max(12, len(h) + 2, max(len(str(c.value or "")) for c in col[:1]) + 2)
    ws.column_dimensions["A"].width = 22
ws = wb["Teste"]
fmt = {"Qui-quadrado (χ²)": "#,##0.00", "p-valor": "0.0E+00", "V de Cramér": "0.000",
       "Proporção de células com esperada < 5": "0.0%", "N do teste": "#,##0"}
for r in ws.iter_rows(min_row=2):
    if r[0].value in fmt:
        r[1].number_format = fmt[r[0].value]
    r[1].alignment = Alignment(wrap_text=True, vertical="top")
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 90
ws = wb["Esperadas_Teste"]
for row in ws.iter_rows(min_row=2, min_col=2):
    for c in row:
        c.number_format = "0.0"
ws.column_dimensions["A"].width = 22
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico (aba Total): barras horizontais empilhadas a 100%
cores = ["#0072B2", "#E69F00", "#009E73", "#CC79A7", "#56B4E9"]
hach = ["", "//", "", "..", "xx"]
t = tabs["Total"][0].iloc[:10].iloc[::-1]
fig, ax = plt.subplots(figsize=(11, 6.5))
esq = np.zeros(len(t))
br = lambda v: f"{v:.1f}%".replace(".", ",")
for k, m in enumerate(MEDS):
    vals = t[f"{m} (%)"].values * 100
    ax.barh(t["Eixo"], vals, left=esq, color=cores[k], hatch=hach[k], edgecolor="white", linewidth=0.8, label=m)
    for y, (v_, l_) in enumerate(zip(vals, esq)):
        if v_ >= 6:
            ax.text(l_ + v_ / 2, y, br(v_), ha="center", va="center", fontsize=8.5,
                    color="white" if k in (0, 2) else "black",
                    bbox=None if k in (0, 2) else dict(facecolor="white", edgecolor="none", alpha=0.8, pad=1))
    esq += vals
# segmentos < 6%: valores listados à direita da barra, com a inicial da categoria
sig = {"Factual": "F", "Publicitária": "P", "De Entretenimento": "E", "Institucional": "I", "Opinativa": "O"}
for y, (_, r) in enumerate(t.iterrows()):
    peq = [f"{sig[m]} {br(r[f'{m} (%)'] * 100)}" for m in MEDS if 0 < r[f"{m} (N)"] and r[f"{m} (%)"] * 100 < 6]
    if peq:
        ax.text(101.5, y, " · ".join(peq), va="center", fontsize=8)
ax.set_yticks(range(len(t)))
ax.set_yticklabels([f"{e} (N = {n:,})".replace(",", ".") for e, n in zip(t["Eixo"], t["N do eixo"])])
ax.set_xlim(0, 100)
ax.set_xlabel("% dos posts do eixo")
ax.xaxis.set_major_formatter(lambda v_, _: f"{v_:.0f}%")
ax.spines[["top", "right"]].set_visible(False)
ax.legend(title="Mediação da Informação", frameon=False, ncol=5, loc="upper center",
          bbox_to_anchor=(0.5, -0.1))
fig.text(0.99, 0.01, "À direita: segmentos < 6% (F Factual · P Publicitária · E De Entretenimento · I Institucional · O Opinativa)",
         ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")

print(TOP, OUTROS, sep="\n")
for nome in ["Total", "Ego", "Roma", "Imperatriz_Online"]:
    print(nome); print(tabs[nome][0].to_string())
print(teste.to_string())
