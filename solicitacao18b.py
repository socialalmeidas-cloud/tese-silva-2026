"""Solicitação 18B — Gráficos de dispersão a partir da Solicitação 18.
(1) Eixos no plano: Análise de Coordenadas Principais (PCoA, MDS clássico) da matriz de distâncias do cosseno
    da Solicitação 18; cor = grupo da solução k = 7 (melhor silhueta); forma = grupo da solução k = 3.
    Método determinístico (decomposição em autovalores): não há semente.
(2) Pares eixo-temática (≥ 5 posts): x = % no eixo, y = lift (escala log), tamanho = N de posts, cor = grupo k = 7.
Lê "Solicitacao18_Eixo_Tematica.xlsx" (gerado por solicitacao18.py) — rode aquele script antes.
Uso: python3 -I solicitacao18b.py
"""
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from adjustText import adjust_text          # só reposiciona rótulos para não se sobreporem (não altera dados)
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ENT = "Solicitacao18_Eixo_Tematica.xlsx"
SAIDA = "Solicitacao18B_Dispersao.xlsx"
PNG_E, PNG_P = "Solicitacao18B_Dispersao_eixos.png", "Solicitacao18B_Dispersao_pares.png"
br = lambda v, d=1: f"{v:.{d}f}".replace(".", ",")
br_int = lambda v: f"{int(v):,}".replace(",", ".")

# ---- leitura
D = pd.read_excel(ENT, sheet_name="Distancias_cosseno").set_index("Distância do cosseno")
A = pd.read_excel(ENT, sheet_name="Associacao")
ag = pd.read_excel(ENT, sheet_name="Agrupamento", header=None)
ini = ag.index[ag[0] == "Eixo"][0]                    # tabela "Grupo de cada eixo"
G = ag.iloc[ini + 1:, :].copy(); G.columns = ag.iloc[ini].tolist()
G = G.dropna(subset=["Eixo"]).set_index("Eixo")
EIXOS = D.index.tolist()
assert set(G.index) == set(EIXOS)
g7 = G["Grupo (k = 7)"].astype(int).reindex(EIXOS)
g3 = G["Grupo (k = 3)"].astype(int).reindex(EIXOS)
NE = G["N do eixo"].astype(int).reindex(EIXOS)

# ---- PCoA (MDS clássico)
Dm = D.values.astype(float)
n = len(Dm)
J = np.eye(n) - np.ones((n, n)) / n
B = -0.5 * J @ (Dm ** 2) @ J
val, vec = np.linalg.eigh(B)
o = np.argsort(val)[::-1]
val, vec = val[o], vec[:, o]
pos = val[val > 1e-12]
coord = vec[:, :2] * np.sqrt(val[:2])
for j in range(2):                                    # sinal fixo (o maior |valor| positivo) para reprodutibilidade
    if coord[np.argmax(np.abs(coord[:, j])), j] < 0:
        coord[:, j] *= -1
expl = val[:2] / pos.sum()                            # proporção da variação (autovalores positivos)
neg = val[val < -1e-12]

# ---- cores e formas
CORES = ["#E69F00", "#56B4E9", "#009E73", "#CC79A7", "#0072B2", "#D55E00", "#000000"]   # Okabe-Ito (sem o amarelo claro)
FORMAS = {1: "s", 2: "^", 3: "o"}
membros7 = {g: [e for e in EIXOS if g7[e] == g] for g in sorted(set(g7))}
membros3 = {g: [e for e in EIXOS if g3[e] == g] for g in sorted(set(g3))}

# ---- (1) dispersão dos eixos
fig, ax = plt.subplots(figsize=(11, 8.5))
for e, (x, y) in zip(EIXOS, coord):
    ax.scatter(x, y, s=60 + 900 * np.sqrt(NE[e] / NE.max()), color=CORES[g7[e] - 1], marker=FORMAS[g3[e]],
               edgecolor="white", linewidth=1.2, alpha=0.9, zorder=3)
textos = [ax.text(x, y, f"{e} ({br_int(NE[e])})", fontsize=8.5, zorder=4) for e, (x, y) in zip(EIXOS, coord)]
adjust_text(textos, x=coord[:, 0], y=coord[:, 1], ax=ax, expand=(1.3, 1.6),
            arrowprops=dict(arrowstyle="-", color="gray", lw=0.5))
ax.axhline(0, color="#dddddd", lw=0.8, zorder=1); ax.axvline(0, color="#dddddd", lw=0.8, zorder=1)
ax.set_xlabel(f"Coordenada principal 1 ({br(expl[0] * 100)}% da variação)")
ax.set_ylabel(f"Coordenada principal 2 ({br(expl[1] * 100)}% da variação)")
fmt = lambda v_, _: br(v_, 1)
ax.xaxis.set_major_formatter(fmt); ax.yaxis.set_major_formatter(fmt)
ax.spines[["top", "right"]].set_visible(False)
h1 = [plt.Line2D([], [], marker="o", ls="", color=CORES[g - 1], markersize=9, label=f"Grupo {g}") for g in membros7]
h2 = [plt.Line2D([], [], marker=FORMAS[g], ls="", color="gray", markersize=8, label=f"Grupo {g}") for g in membros3]
l1 = ax.legend(handles=h1, title="Cor: solução k = 7", frameon=False, loc="upper left", bbox_to_anchor=(1.0, 1.0), fontsize=8.5)
ax.add_artist(l1)
ax.legend(handles=h2, title="Forma: solução k = 3", frameon=False, loc="upper left", bbox_to_anchor=(1.0, 0.55), fontsize=8.5)
fig.text(0.99, 0.005, f"PCoA da distância do cosseno entre eixos (vetor = % de posts do eixo com cada temática) · tamanho = N do eixo · "
         f"os dois eixos somam {br(expl.sum() * 100)}% da variação · Esporte perto do centro NÃO indica perfil médio: está longe de todos nas outras dimensões · exploratório", ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(PNG_E, dpi=300, bbox_inches="tight"); plt.close(fig)

# ---- (2) dispersão dos pares
A["Grupo (k = 7)"] = A["Eixo"].map(g7)
fig, ax = plt.subplots(figsize=(12, 8))
for g, d in A.groupby("Grupo (k = 7)"):
    ax.scatter(d["% no eixo"] * 100, d["Lift"], s=12 + 2.2 * d["N de posts"], color=CORES[g - 1], alpha=0.75,
               edgecolor="white", linewidth=0.7, label=f"Grupo {g}: " + ", ".join(membros7[g]), zorder=3)
ax.set_yscale("log")
ax.axhline(1, color="gray", lw=0.9, ls=":", zorder=1)
ax.axhline(2, color="gray", lw=0.9, ls="--", zorder=1)
xmax = A["% no eixo"].max() * 100
ax.text(xmax, 1.04, "lift = 1 (independência)", fontsize=7.5, color="dimgray", ha="right")
ax.text(xmax, 2.08, "lift = 2", fontsize=7.5, color="dimgray", ha="right")
# rótulos: 12 pares de maior N com lift > 2 e ≥ 10 posts FORA da Publicidade, + os 3 de maior N da Publicidade,
# + os 5 de maior % no eixo (a Publicidade tem muitos pares sobrepostos perto de lift 2)
ok = A[(A["Lift"] > 2) & (A["N de posts"] >= 10)]
rot = pd.concat([ok[ok["Eixo"] != "Publicidade"].nlargest(12, "N de posts"), A[A["Eixo"] == "Publicidade"].nlargest(3, "N de posts"),
                 A.nlargest(5, "% no eixo")]).drop_duplicates(["Eixo", "Temática"])
textos = [ax.text(r._4 * 100, r.Lift, f"{r.Temática} · {r.Eixo}", fontsize=7.2, zorder=4) for r in rot.itertuples()]
adjust_text(textos, x=(A["% no eixo"] * 100).values, y=A["Lift"].values, ax=ax, expand=(1.2, 1.5),
            arrowprops=dict(arrowstyle="-", color="gray", lw=0.5))
ax.set_xlabel("% dos posts do eixo com a temática")
ax.set_ylabel("Lift (escala logarítmica)")
ax.xaxis.set_major_formatter(lambda v_, _: f"{v_:.0f}%")
ax.yaxis.set_major_formatter(lambda v_, _: br(v_, 0) if v_ >= 1 else br(v_, 1))
ax.spines[["top", "right"]].set_visible(False)
hs = [plt.Line2D([], [], marker="o", ls="", color=CORES[g - 1], markersize=8, label=f"Grupo {g}: " + ", ".join(membros7[g])) for g in membros7]
hs += [plt.scatter([], [], s=12 + 2.2 * n, color="lightgray", edgecolor="gray", label=f"{n} posts") for n in (10, 50, 150)]
ax.legend(handles=hs, title="Cor = grupo do eixo (k = 7) · tamanho = N de posts do par", frameon=False, fontsize=7.5,
          loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=2)
fig.text(0.99, 0.003, f"{len(A)} pares eixo-temática com ≥ 5 posts (eixos com ≥ 30 posts; N = 4.036) · rótulos: 12 pares de maior N com lift > 2 e ≥ 10 posts fora da Publicidade, 3 maiores da Publicidade e 5 de maior % · exploratório",
         ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(PNG_P, dpi=300, bbox_inches="tight"); plt.close(fig)

# ---- planilha
T = pd.DataFrame({"Eixo": EIXOS, "N do eixo": NE.values, "Coordenada 1": coord[:, 0], "Coordenada 2": coord[:, 1],
                  "Grupo (k = 7)": g7.values, "Grupo (k = 3)": g3.values})
V = pd.DataFrame({"Eixo principal": [f"Coordenada {i}" for i in range(1, 6)], "Autovalor": val[:5],
                  "% da variação (autovalores positivos)": val[:5] / pos.sum()})
notas = pd.DataFrame([
    ["Base", "Solicitação 18 (planilha congelada de 04/11/2026, trabalhada em cópia; variantes aprovadas em 07/11/2026): 19 eixos com ≥ 30 posts, 4.036 posts, 258 temáticas com ≥ 5 posts."],
    ["Gráfico 1", "Análise de Coordenadas Principais (PCoA, MDS clássico) da matriz de distâncias do cosseno entre eixos da Solicitação 18. "
     f"As duas primeiras coordenadas somam {br(expl.sum() * 100)}% da variação (sobre os autovalores positivos)"
     + (f"; há {len(neg)} autovalores negativos, porque a distância do cosseno não é euclidiana" if len(neg) else "") + ". "
     "Proximidade no plano = perfis de temáticas parecidos, mas o plano mostra só parte da informação (Esporte aparece perto do centro porque está longe de todos os eixos nas dimensões seguintes, não por ter perfil médio): a distância exata está na aba Distancias_cosseno da S18. "
     "Cor = grupo na solução k = 7; forma = grupo na solução k = 3; tamanho = N do eixo. Método determinístico: não há semente."],
    ["Gráfico 2", f"{len(A)} pares eixo-temática com ≥ 5 posts (aba Associacao da S18). x = % dos posts do eixo com a temática; y = lift (escala log); "
     "tamanho = N de posts do par; cor = grupo do eixo (k = 7). Linhas: lift 1 (independência) e lift 2 (limite usado na rede da S18). "
     "Pares com lift alto e % baixa costumam ter poucos posts: lift instável."],
    ["RESSALVA", "Exploratório: PCoA, agrupamento e lift dependem dos filtros, da codificação livre e das variantes aprovadas. Não são testes de hipótese."],
    ["Script", "solicitacao18b.py (lê Solicitacao18_Eixo_Tematica.xlsx)"],
], columns=["Item", "Descrição"])
with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    T.to_excel(w, sheet_name="Coordenadas_eixos", index=False)
    V.to_excel(w, sheet_name="Autovalores", index=False)
    A.to_excel(w, sheet_name="Pares", index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)
wb = load_workbook(SAIDA)
for ws in wb.worksheets:
    cab = {c.column: str(c.value) for c in ws[1]}
    for c in ws[1]:
        c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 24
    for row in ws.iter_rows(min_row=2):
        for c in row:
            h = cab.get(c.column, "")
            if isinstance(c.value, float):
                c.number_format = "0.0%" if h.startswith("%") else ("0.00" if h == "Lift" else "0.000")
            elif isinstance(c.value, int):
                c.number_format = "#,##0"
ws = wb["Notas"]; ws.column_dimensions["B"].width = 130
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)
print("expl", expl, "neg", len(neg)); print(T.round(3).to_string())
