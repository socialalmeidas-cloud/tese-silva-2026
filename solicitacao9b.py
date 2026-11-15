"""Solicitação 9B — Gráficos de Referência Local.
Lê as tabelas geradas por solicitacao9.py (com os agrupamentos aprovados) e gera:
 (a) Solicitacao9B_Top15_locais.png — 15 locais mais citados por perfil (versão B: sem Não identificado, Imperatriz, Maranhão)
 (b) Solicitacao9B_Nao_identificado.png — % de posts com "Não identificado" por perfil
Uso: python3 solicitacao9.py <planilha.xlsx> && python3 solicitacao9b.py
"""
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator

TAB = "Solicitacao9_Referencia_local.xlsx"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
N = {"Ego": 414, "Roma": 1131, "Imperatriz Online": 2721}
CORES = {"Ego": "#E69F00", "Roma": "#0072B2", "Imperatriz Online": "#009E73"}
HACH = {"Ego": "", "Roma": "//", "Imperatriz Online": ".."}
TOP = 15
br = lambda v: f"{v:.1f}%".replace(".", ",")
milhar = lambda n: f"{int(n):,}".replace(",", ".")

# ---- (a) Top 15 por perfil, versão B
blocos = {}
for p in PERFIS:
    y = pd.read_excel(TAB, sheet_name="Top30_" + p.replace(" ", "_"), header=1)
    b = y.iloc[:, 5:9].dropna()
    b.columns = ["Posição", "Local", "N", "%"]
    b["pos"] = b["Posição"].astype(str).str.rstrip("=").astype(int)
    b = b[b["pos"] <= TOP].copy()  # empates na 15ª posição são todos incluídos
    assert (b["N"].astype(int) <= N[p]).all()
    blocos[p] = b
    print(p, "linhas:", len(b), "| última posição:", b["Posição"].iloc[-1])

altura = max(len(b) for b in blocos.values())
xmax = max(b["%"].max() for b in blocos.values()) * 100 * 1.45
fig, axes = plt.subplots(1, 3, figsize=(16, 0.38 * altura + 1.8))
for ax, p in zip(axes, PERFIS):
    b = blocos[p].iloc[::-1]
    v = b["%"].values * 100
    ax.barh(range(len(b)), v, color=CORES[p], hatch=HACH[p], edgecolor="white", linewidth=0.5)
    for yy, (vv, n) in enumerate(zip(v, b["N"])):
        ax.text(vv + 0.15, yy, f"{br(vv)} ({milhar(n)})", va="center", fontsize=7.5)
    ax.set_yticks(range(len(b)))
    ax.set_yticklabels([f"{pos} {loc}" for pos, loc in zip(b["Posição"], b["Local"])], fontsize=8.5)
    ax.set_xlim(0, xmax)
    ax.xaxis.set_major_locator(MultipleLocator(2))  # marcas em 0, 2, 4... (evita rótulo arredondado)
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:.0f}%")
    ax.set_xlabel(f"% dos posts do perfil (N = {milhar(N[p])})")
    ax.set_title(p, fontsize=10)  # identificação do painel, não título do gráfico
    ax.spines[["top", "right"]].set_visible(False)
fig.text(0.99, 0.005, "Sem \"Não identificado\", \"Imperatriz\" e \"Maranhão\". Um post pode citar vários locais. "
         "\"=\" empate; * menos de 10 posts no perfil.", ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(w_pad=3)
fig.savefig("Solicitacao9B_Top15_locais.png", dpi=300, bbox_inches="tight")

# ---- (b) Não identificado por perfil
ni = pd.read_excel(TAB, sheet_name="Nao_identificado").set_index("Perfil")
fig, ax = plt.subplots(figsize=(7, 5))
vals = [ni.loc[p, "Não identificado (%)"] * 100 for p in PERFIS]
ns = [ni.loc[p, "Não identificado (N)"] for p in PERFIS]
for i, p in enumerate(PERFIS):
    ax.bar(i, vals[i], width=0.6, color=CORES[p], hatch=HACH[p], edgecolor="white", linewidth=0.5)
    ax.text(i, vals[i] + 0.25, f"{br(vals[i])}\n({milhar(ns[i])} de {milhar(N[p])})", ha="center", va="bottom", fontsize=9)
ax.set_xticks(range(len(PERFIS)))
ax.set_xticklabels(PERFIS)
ax.set_ylim(0, max(vals) * 1.35)
ax.set_ylabel("% dos posts do perfil com \"Não identificado\"")
ax.yaxis.set_major_formatter(lambda x, _: f"{x:.0f}%")
ax.spines[["top", "right"]].set_visible(False)
fig.tight_layout()
fig.savefig("Solicitacao9B_Nao_identificado.png", dpi=300, bbox_inches="tight")
print(dict(zip(PERFIS, ns)), [round(v, 1) for v in vals])
