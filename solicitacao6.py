"""Solicitação 6 — Regime de Visibilidade × Eixo Temático.
Conta posts com cada rótulo de Regime dentro de cada eixo (multirrótulo, separador "; ").
Uso: python3 solicitacao6.py <planilha.xlsx>
Gera Solicitacao6_Regime_x_Eixo.xlsx e .png.
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
SAIDA = "Solicitacao6_Regime_x_Eixo.xlsx"
PNG = "Solicitacao6_Regime_x_Eixo.png"
EIXO, REG = "Eixo Temático", "Regime de Visibilidade"
EMERG = {"Criminoso", "Animal", "Bens materiais", "Candidatos", "Outro perfil"}

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
assert df["Perfil"].value_counts().to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
assert not df[EIXO].astype(str).str.contains("; ").any()

# eixos: top 10 do total geral + Criminalidade e Segurança Pública se faltarem
freq = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
EIXOS = freq["e"].iloc[:10].tolist()
extras = [e for e in ("Criminalidade", "Segurança Pública") if e not in EIXOS]
EIXOS += extras

df["rot"] = df[REG].astype(str).str.split(";").apply(lambda l: sorted({x.strip() for x in l if x.strip()}))
long = df[["url ", EIXO, "rot"]].explode("rot").reset_index(drop=True)
# ordem das colunas = ordem do total geral (igual à Solicitação 5)
cont = long["rot"].value_counts().rename_axis("r").reset_index(name="n").sort_values(["n", "r"], ascending=[False, True])
REGS = cont["r"].tolist()
assert len(REGS) == 13
rot = {r: r + (" †" if r in EMERG else "") for r in REGS}

ct = pd.crosstab(long[EIXO], long["rot"]).reindex(index=EIXOS, columns=REGS, fill_value=0)
n_eixo = df[EIXO].value_counts().reindex(EIXOS)
assert (ct.values <= n_eixo.values[:, None]).all()

t_n = ct.copy()
t_n.columns = [rot[r] for r in REGS]
t_n.insert(0, "N do eixo", n_eixo.values)
t_n.index.name = "Eixo"
t_n = t_n.reset_index()

t_p = ct.div(n_eixo, axis=0)
t_p.columns = [rot[r] for r in REGS]
t_p.insert(0, "N do eixo", n_eixo.values)
t_p.index.name = "Eixo"
t_p = t_p.reset_index()
t_p["Soma dos % da linha"] = ct.sum(axis=1).values / n_eixo.values

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", f"{int(n_eixo.sum()):,} posts dos {len(EIXOS)} eixos (de 4.266 no corpus)".replace(",", ".")],
    ["Exclusões", "Os 24 eixos fora do top 10 (588 posts). Criminalidade e Segurança Pública já estão no top 10; nenhum eixo foi acrescentado."],
    ["Eixos incluídos", "; ".join(EIXOS)],
    ["Contagem", "Regime é multirrótulo: a célula foi separada por \"; \" e cada post conta uma vez para cada rótulo que tem (posts, não células)."],
    ["Fórmula do %", "% = nº de posts do eixo com o rótulo ÷ N do eixo × 100. Como um post pode ter vários rótulos, a linha pode somar mais de 100% (ver coluna \"Soma dos % da linha\")."],
    ["Ordem das colunas", "Rótulos de Regime em ordem decrescente do total geral (mesma ordem da Solicitação 5)."],
    ["† CATEGORIAS EMERGENTES", "Criminoso †, Animal †, Bens materiais †, Candidatos †, Outro perfil † (colunas destacadas em amarelo nas abas N e Percentual). Surgiram por \"Outro (especificar)\" e foram aplicadas apenas a partir do momento em que surgiram, de forma desigual entre lotes. Resultados indicativos."],
    ["RESSALVA — CRIMINOSO", "\"Criminoso\" foi aplicado de forma desigual entre lotes (0% a 71% dos posts de Criminalidade e Segurança Pública), resultado indicativo."],
    ["Lista completa de emergentes", "Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix."],
    ["Script", "solicitacao6.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t_n.to_excel(w, sheet_name="N", index=False)
    t_p.to_excel(w, sheet_name="Percentual", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

amarelo = PatternFill("solid", fgColor="FFF2CC")
wb = load_workbook(SAIDA)
for nome in ("N", "Percentual"):
    ws = wb[nome]
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[1].height = 45
    for col in ws.columns:
        h = str(col[0].value)
        for c in col[1:]:
            if h == "N do eixo" or nome == "N":
                c.number_format = "#,##0"
            else:
                c.number_format = "0.0%"
            if "†" in h:
                c.fill = amarelo
        if "†" in h:
            col[0].fill = amarelo
            col[0].font = Font(bold=True, color="9C0006")
        ws.column_dimensions[col[0].column_letter].width = 14
    ws.column_dimensions["A"].width = 20
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
    if row[0].value in ("† CATEGORIAS EMERGENTES", "RESSALVA — CRIMINOSO"):
        for c in row:
            c.fill = amarelo
            c.font = Font(bold=True)
wb.save(SAIDA)

# ---- Mapa de calor (cividis: sequencial e legível com daltonismo)
M = (ct.div(n_eixo, axis=0) * 100).values
fig, ax = plt.subplots(figsize=(14, 7))
im = ax.imshow(M, cmap="cividis_r", vmin=0, vmax=100, aspect="auto")
for i in range(M.shape[0]):
    for j in range(M.shape[1]):
        v = M[i, j]
        ax.text(j, i, f"{v:.1f}".replace(".", ","), ha="center", va="center", fontsize=7.5,
                color="white" if v > 55 else "black")
ax.set_xticks(range(len(REGS)))
ax.set_xticklabels([rot[r] for r in REGS], rotation=40, ha="right", fontsize=8.5)
for lab in ax.get_xticklabels():
    if "†" in lab.get_text():
        lab.set_fontweight("bold")
ax.set_yticks(range(len(EIXOS)))
ax.set_yticklabels([f"{e} (N = {n:,})".replace(",", ".") for e, n in zip(EIXOS, n_eixo)], fontsize=9)
cb = fig.colorbar(im, ax=ax, fraction=0.025, pad=0.01)
cb.set_label("% dos posts do eixo")
cb.ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
ax.set_xlabel("Regime de Visibilidade (valores em %; um post pode ter mais de um rótulo)")
fig.text(0.01, 0.005, "† categoria emergente (em negrito): aplicada de forma desigual entre lotes; resultado indicativo.",
         fontsize=7.5, color="dimgray")
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")

pd.set_option("display.width", 300)
print(EIXOS)
print(t_n.to_string())
print((t_p.set_index("Eixo").drop(columns="N do eixo") * 100).round(1).to_string())
