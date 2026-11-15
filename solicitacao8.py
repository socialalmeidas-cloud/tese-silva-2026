"""Solicitação 8 — Impacto Visual por Recurso Audiovisual e por Perfil + teste χ² (Impacto × Perfil).
Uso: python3 solicitacao8.py <planilha.xlsx>
Gera Solicitacao8_Impacto_visual.xlsx e .png.
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
SAIDA = "Solicitacao8_Impacto_visual.xlsx"
PNG = "Solicitacao8_Impacto_visual.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
IMP, REC = "Impacto Visual", "Recursos Audiovisuais"
NIV = ["Alto", "Médio", "Baixo"]
EMERG = {"Vídeo sem som", "Desenho"}
MIN = 30

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
for c in (IMP, REC):
    assert not df[c].astype(str).str.contains("; ").any()
assert set(df[IMP]) == set(NIV), set(df[IMP])

freq = df[REC].value_counts().rename_axis("r").reset_index(name="n").sort_values(["n", "r"], ascending=[False, True])
MANTER = freq.loc[freq["n"] >= MIN, "r"].tolist()
AGRUP = freq.loc[freq["n"] < MIN, ["r", "n"]].values.tolist()
df["rec_g"] = df[REC].where(df[REC].isin(MANTER), "Outros recursos")


def dist(grupo, ordem, total_label="Total"):
    ct = pd.crosstab(df[grupo], df[IMP]).reindex(index=ordem, columns=NIV, fill_value=0)
    ct.loc[total_label] = ct.sum()
    out = pd.DataFrame(index=ct.index)
    out["N"] = ct.sum(axis=1)
    for v in NIV:
        out[f"{v} (N)"] = ct[v]
        out[f"{v} (%)"] = ct[v] / out["N"]
    assert out.iloc[-1]["N"] == len(df)
    assert out.iloc[:-1]["N"].sum() == len(df)
    assert np.allclose(out[[f"{v} (%)" for v in NIV]].sum(axis=1), 1)
    return out, ct


t_rec, ct_rec = dist("rec_g", MANTER + ["Outros recursos"])
t_rec = t_rec.rename(columns={"N": "N do recurso"}).reset_index().rename(columns={"rec_g": "Recurso", "index": "Recurso"})
t_per, ct_per = dist("Perfil", PERFIS)
t_per = t_per.rename(columns={"N": "N do perfil"}).reset_index().rename(columns={"Perfil": "Perfil", "index": "Perfil"})

# ---- Teste: Impacto Visual × Perfil (3 × 3)
obs = ct_per.loc[PERFIS, NIV].values
chi2, p, gl, esp = chi2_contingency(obs, correction=False)
n = obs.sum()
v = np.sqrt(chi2 / (n * (min(obs.shape) - 1)))
teste = pd.DataFrame([
    ["Tabela testada", "Impacto Visual (Alto, Médio, Baixo) × Perfil (Ego, Roma, Imperatriz Online)"],
    ["N do teste", int(n)],
    ["Qui-quadrado (χ²)", chi2],
    ["Graus de liberdade", int(gl)],
    ["p-valor", p if p > 0 else "p < 0,001 (menor que a precisão numérica; o programa retorna 0)"],
    ["V de Cramér", v],
    ["Células com frequência esperada < 5", f"{(esp < 5).sum()} de {esp.size}"],
    ["Proporção de células com esperada < 5", (esp < 5).mean()],
    ["Menor frequência esperada", esp.min()],
    ["Avaliação", "Pressuposto atendido (nenhuma célula com esperada < 5)." if (esp < 5).mean() == 0
     else ("Teste confiável (≤ 20% das células com esperada < 5)." if (esp < 5).mean() <= 0.2 else "Teste NÃO confiável (> 20% das células com esperada < 5).")],
    ["Observação", "Pearson sem correção de continuidade (scipy.stats.chi2_contingency, correction=False). V = √(χ² / (N × (min(linhas, colunas) − 1))). A referência metodológica precisa ser verificada pela autora."],
], columns=["Item", "Valor"])
esperadas = pd.DataFrame(esp, index=PERFIS, columns=NIV).reset_index().rename(columns={"index": "Perfil"})

agrup_txt = "; ".join(f"{r}{' †' if r in EMERG else ''} ({k})" for r, k in AGRUP)
notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma."],
    ["Variáveis", "Impacto Visual e Recursos Audiovisuais: rótulo único por post. Impacto Visual tem 3 níveis no corpus (Alto, Médio, Baixo); \"Não há\" não ocorre."],
    ["Fórmula do %", "Percentual de linha: % = N do nível ÷ N do recurso (ou do perfil) × 100. Cada linha soma 100%."],
    ["Recursos mantidos (≥ 30 posts)", "; ".join(MANTER)],
    ["Agrupados em \"Outros recursos\"", agrup_txt + f". Total: {sum(k for _, k in AGRUP)} posts."],
    ["† Categorias emergentes", "Vídeo sem som † e Desenho † estão dentro de \"Outros recursos\". Surgiram por \"Outro (especificar)\" e foram aplicadas apenas a partir do momento em que surgiram; resultados que dependem delas são indicativos."],
    ["Lista completa de emergentes", "Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix."],
    ["Teste", "Ver abas Teste e Esperadas_Teste."],
    ["Script", "solicitacao8.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t_rec.to_excel(w, sheet_name="Por_recurso", index=False)
    t_per.to_excel(w, sheet_name="Por_perfil", index=False)
    teste.to_excel(w, sheet_name="Teste", index=False)
    esperadas.to_excel(w, sheet_name="Esperadas_Teste", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
for nome in ("Por_recurso", "Por_perfil"):
    ws = wb[nome]
    for c in ws[1] + ws[ws.max_row]:
        c.font = Font(bold=True)
    for col in ws.columns:
        h = str(col[0].value)
        for c in col[1:]:
            c.number_format = "0.0%" if "(%)" in h else "#,##0" if ("(N)" in h or h.startswith("N do")) else c.number_format
        ws.column_dimensions[col[0].column_letter].width = max(12, len(h) + 2)
    ws.column_dimensions["A"].width = 24
ws = wb["Teste"]
fmt = {"N do teste": "#,##0", "Qui-quadrado (χ²)": "#,##0.00", "V de Cramér": "0.000",
       "Proporção de células com esperada < 5": "0.0%", "Menor frequência esperada": "#,##0.0"}
for r in ws.iter_rows(min_row=2):
    if r[0].value in fmt:
        r[1].number_format = fmt[r[0].value]
    if r[0].value == "p-valor" and isinstance(r[1].value, float):
        r[1].number_format = "0.0E+00"
    r[1].alignment = Alignment(wrap_text=True, vertical="top")
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 90
ws = wb["Esperadas_Teste"]
for row in ws.iter_rows(min_row=2, min_col=2):
    for c in row:
        c.number_format = "#,##0.0"
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico: dois painéis, barras horizontais empilhadas a 100%
cores = {"Alto": "#D55E00", "Médio": "#F0E442", "Baixo": "#56B4E9"}
hach = {"Alto": "", "Médio": "//", "Baixo": ".."}
br = lambda x: f"{x:.1f}%".replace(".", ",")
fig, axes = plt.subplots(1, 2, figsize=(15, 6.5), gridspec_kw={"width_ratios": [1.1, 1]})
for ax, (t, col, ncol) in zip(axes, [(t_rec, "Recurso", "N do recurso"), (t_per, "Perfil", "N do perfil")]):
    d = t.iloc[::-1]
    esq = np.zeros(len(d))
    for nv in NIV:
        vals = d[f"{nv} (%)"].values * 100
        ax.barh(range(len(d)), vals, left=esq, color=cores[nv], hatch=hach[nv], edgecolor="white", linewidth=0.6, label=nv)
        for y, (vv, l) in enumerate(zip(vals, esq)):
            if vv >= 5:
                ax.text(l + vv / 2, y, br(vv), ha="center", va="center", fontsize=8,
                        bbox=dict(facecolor="white", edgecolor="none", alpha=0.75, pad=0.8))
        esq += vals
    # segmentos < 5% listados à direita
    for y, (_, r) in enumerate(d.iterrows()):
        peq = [f"{nv[0]} {br(r[f'{nv} (%)'] * 100)}" for nv in NIV if 0 < r[f"{nv} (N)"] and r[f"{nv} (%)"] * 100 < 5]
        if peq:
            ax.text(101.5, y, " · ".join(peq), va="center", fontsize=7.5)
    ax.set_yticks(range(len(d)))
    ax.set_yticklabels([f"{e} (N = {k:,})".replace(",", ".") for e, k in zip(d[col], d[ncol])], fontsize=9)
    ax.set_xlim(0, 100)
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:.0f}%")
    ax.set_xlabel(f"% dos posts do {col.lower()}")
    ax.set_title(f"Por {col.lower()}", fontsize=10)  # identificação do painel
    ax.spines[["top", "right"]].set_visible(False)
h, l = axes[0].get_legend_handles_labels()
fig.legend(h, l, title="Impacto Visual", frameon=False, ncol=3, loc="lower center", bbox_to_anchor=(0.5, -0.04))
fig.text(0.99, -0.04, "À direita: segmentos < 5% (A Alto · M Médio · B Baixo)", ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(w_pad=6)
fig.savefig(PNG, dpi=300, bbox_inches="tight")

pd.set_option("display.width", 250)
print(MANTER, AGRUP)
print(t_rec.to_string()); print(t_per.to_string()); print(teste.to_string()); print(esperadas.round(1).to_string())
