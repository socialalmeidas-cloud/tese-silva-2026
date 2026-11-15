"""Solicitação 18 — Eixo Temático × Temática Secundária: associação (lift) e agrupamento de eixos.

Uso: python3 -I solicitacao18.py <planilha.xlsx>

Passos:
  1. Lê a cópia da planilha (a original não é alterada). Separa "Temática Secundária " (nome real, com espaço
     no fim) por ";" e aplica as variantes de grafia aprovadas pela autora em 07/11/2026 (variantes_aprovadas.py).
     Conta POSTS: cada rótulo no máximo uma vez por post.
  2. Filtros: temáticas com ≥ 5 posts (no corpus inteiro, 4.266 posts); eixos com ≥ 30 posts.
  3. Matriz eixo × temática (N e % no eixo), associação (lift) e top por eixo.
  4. Agrupamento hierárquico dos eixos: vetor = % de posts do eixo com cada temática; distância do cosseno;
     ligação média (UPGMA). Silhueta para 2 a 8 grupos.
     O método é determinístico: não há sorteio, portanto não há semente.
  5. Gráficos: mapa de calor (30 temáticas), dendrograma e rede bipartida (lift > 2 e ≥ 10 posts).
"""
import os
import sys

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import linkage, fcluster, dendrogram
from scipy.spatial.distance import pdist, squareform
from sklearn.metrics import silhouette_score
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))   # pasta do script (python -I não a inclui)
from variantes_aprovadas import rotulos_do_post, MAPA, DATA_DECISAO

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao18_Eixo_Tematica.xlsx"
PNG_CALOR = "Solicitacao18_Mapa_de_calor_Eixo_x_Tematica.png"
PNG_DENDRO = "Solicitacao18_Dendrograma_eixos.png"
PNG_REDE = "Solicitacao18_Rede_Eixo_Tematica.png"
COL, EIXO = "Temática Secundária ", "Eixo Temático"
MIN_TEM, MIN_EIXO = 5, 30          # filtros pedidos
MIN_PAR, POUCOS = 5, 10            # pares na aba Associacao; sinal "poucos posts"
LIFT_REDE, N_REDE = 2, 10          # arestas da rede
KS = range(2, 9)                   # nº de grupos testados na silhueta
br_int = lambda v: f"{int(v):,}".replace(",", ".")
br_dec = lambda v, d=1: f"{v:.{d}f}".replace(".", ",")

# ------------------------------------------------------------------ 1. leitura
df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
df["Linha"] = df.index + 2
df[EIXO] = df[EIXO].astype(str).str.strip()
df["rot"] = df[COL].apply(rotulos_do_post)
assert df["rot"].apply(len).min() >= 1                       # todo post tem ao menos um rótulo

# ------------------------------------------------------------------ 2. filtros
freq_tem = df["rot"].explode().value_counts()                 # posts por temática no corpus inteiro
TEM_OK = set(freq_tem[freq_tem >= MIN_TEM].index)
tem_fora = freq_tem[freq_tem < MIN_TEM]
posts_sem_tem = int(df["rot"].apply(lambda l: not any(x in TEM_OK for x in l)).sum())

freq_eixo = df[EIXO].value_counts()
EIXOS = freq_eixo[freq_eixo >= MIN_EIXO].sort_values(ascending=False).index.tolist()
eixos_fora = freq_eixo[freq_eixo < MIN_EIXO].sort_values(ascending=False)

base = df[df[EIXO].isin(EIXOS)].copy()                        # posts usados
base["rot_ok"] = base["rot"].apply(lambda l: [x for x in l if x in TEM_OK])
N = len(base)
NE = base[EIXO].value_counts().reindex(EIXOS)
assert N + eixos_fora.sum() == 4266

# ------------------------------------------------------------------ 3. matriz eixo × temática
ex = base[[EIXO, "rot_ok"]].explode("rot_ok").dropna().reset_index(drop=True)
M = pd.crosstab(ex[EIXO], ex["rot_ok"]).reindex(index=EIXOS, fill_value=0)
NT = M.sum(axis=0)                                            # posts com a temática dentro da base
M = M.loc[:, NT.sort_values(ascending=False, kind="mergesort").index]
M = M.loc[:, M.sum() > 0]
NT = M.sum(axis=0)
# ordem das temáticas: pelo total (empates em ordem alfabética)
ordem_t = sorted(M.columns, key=lambda t: (-NT[t], t))
M, NT = M[ordem_t], NT[ordem_t]
P = M.div(NE, axis=0)                                         # % no eixo
assert (M.values <= NE.values[:, None]).all()

tem_fora_base = sorted(TEM_OK - set(M.columns))               # passaram no filtro, mas só ocorrem em eixos excluídos

# ------------------------------------------------------------------ 4. associação (lift)
pares = []
for e in EIXOS:
    for t in M.columns:
        n = int(M.loc[e, t])
        if n >= MIN_PAR:
            pares.append({"Eixo": e, "Temática": t, "N de posts": n,
                          "% no eixo": n / NE[e],
                          "% da temática que está no eixo": n / NT[t],
                          "Lift": n * N / (NE[e] * NT[t]),
                          "N do eixo": int(NE[e]), "N da temática (base)": int(NT[t]),
                          "Poucos posts (< 10)": "*" if n < POUCOS else ""})
A = pd.DataFrame(pares).sort_values(["Lift", "N de posts"], ascending=[False, False]).reset_index(drop=True)

# top por eixo: 10 por % no eixo; 10 por lift (pares com ≥ 10 posts)
tops = []
for e in EIXOS:
    p = P.loc[e].sort_values(ascending=False, kind="mergesort")
    p = p[p > 0]
    p = p.loc[sorted(p.index, key=lambda t: (-p[t], t))].head(10)
    for i, (t, v) in enumerate(p.items(), 1):
        lift = M.loc[e, t] * N / (NE[e] * NT[t])
        tops.append({"Eixo": e, "Critério": "Maior % no eixo", "Posição": i, "Temática": t,
                     "N de posts": int(M.loc[e, t]), "% no eixo": v, "Lift": lift,
                     "Poucos posts (< 10)": "*" if M.loc[e, t] < POUCOS else ""})
    a = A[(A["Eixo"] == e) & (A["N de posts"] >= POUCOS)].head(10)
    for i, r in enumerate(a.itertuples(), 1):
        tops.append({"Eixo": e, "Critério": "Maior lift (≥ 10 posts)", "Posição": i, "Temática": r.Temática,
                     "N de posts": r._3, "% no eixo": r._4, "Lift": r.Lift, "Poucos posts (< 10)": ""})
    if len(a) == 0:
        tops.append({"Eixo": e, "Critério": "Maior lift (≥ 10 posts)", "Posição": "", "Temática": "(nenhum par com ≥ 10 posts)",
                     "N de posts": np.nan, "% no eixo": np.nan, "Lift": np.nan, "Poucos posts (< 10)": ""})
T = pd.DataFrame(tops)

# ------------------------------------------------------------------ 5. agrupamento hierárquico dos eixos
X = P.values                                                  # vetor de % por eixo
D = pdist(X, metric="cosine")
Z = linkage(D, method="average")
DQ = squareform(D)
sil = []
for k in KS:
    lab = fcluster(Z, t=k, criterion="maxclust")
    nk = len(set(lab))
    s = silhouette_score(DQ, lab, metric="precomputed") if 1 < nk < len(EIXOS) else np.nan
    sil.append({"Nº de grupos (k)": k, "Grupos obtidos": nk, "Silhueta média": s,
                "Eixos sozinhos num grupo": int(sum((lab == g).sum() == 1 for g in set(lab)))})
S = pd.DataFrame(sil)
K_MELHOR = int(S.loc[S["Silhueta média"].idxmax(), "Nº de grupos (k)"])
SOLUCOES = sorted({K_MELHOR, 3})


def renumerar(lab):
    """Numera os grupos na ordem em que aparecem no dendrograma (de cima para baixo), para ficar estável."""
    folhas = dendrogram(Z, no_plot=True)["leaves"]
    novo, mapa = [], {}
    for i in folhas:
        mapa.setdefault(lab[i], len(mapa) + 1)
    return np.array([mapa[g] for g in lab])


grupos = {k: renumerar(fcluster(Z, t=k, criterion="maxclust")) for k in SOLUCOES}
sil_eixo = {}
from sklearn.metrics import silhouette_samples
for k in SOLUCOES:
    sil_eixo[k] = silhouette_samples(DQ, grupos[k], metric="precomputed")

linhas_g, carac = [], []
for k in SOLUCOES:
    for g in sorted(set(grupos[k])):
        membros = [e for e, l in zip(EIXOS, grupos[k]) if l == g]
        linhas_g.append({"Solução": f"k = {k}" + (" (melhor silhueta)" if k == K_MELHOR else "") + (" (3 grupos, pedida)" if k == 3 and k != K_MELHOR else ""),
                         "Grupo": g, "Nº de eixos": len(membros),
                         "Eixos (N de posts)": "; ".join(f"{e} ({br_int(NE[e])})" for e in membros),
                         "N de posts do grupo": int(NE[membros].sum()),
                         "Silhueta média do grupo": float(np.mean([s for s, l in zip(sil_eixo[k], grupos[k]) if l == g]))})
        media = P.loc[membros].mean().sort_values(ascending=False)
        media = media.loc[sorted(media.index, key=lambda t: (-media[t], t))].head(10)
        for i, (t, v) in enumerate(media.items(), 1):
            carac.append({"Solução": f"k = {k}", "Grupo": g, "Eixos do grupo": ", ".join(membros), "Posição": i,
                          "Temática": t, "% médio nos eixos do grupo": v,
                          "Nº de eixos do grupo em que aparece": int((M.loc[membros, t] > 0).sum()),
                          "N de posts (soma no grupo)": int(M.loc[membros, t].sum())})
G, C = pd.DataFrame(linhas_g), pd.DataFrame(carac)
eixo_grupos = pd.DataFrame({"Eixo": EIXOS, "N do eixo": NE.values.astype(int),
                            **{f"Grupo (k = {k})": grupos[k] for k in SOLUCOES},
                            **{f"Silhueta do eixo (k = {k})": sil_eixo[k] for k in SOLUCOES}})
Dist = pd.DataFrame(DQ, index=EIXOS, columns=EIXOS).reset_index().rename(columns={"index": "Distância do cosseno"})

# ------------------------------------------------------------------ 6. abas de matriz
def com_total(m, pct):
    out = m.copy()
    tot = (m.sum(axis=0) if not pct else M.sum(axis=0) / N)
    out.loc["Total"] = tot
    out.insert(0, "N do eixo", list(NE.values.astype(int)) + [N])
    out.index.name = "Eixo"
    return out.reset_index()

mat_n = com_total(M, False)
mat_p = com_total(P, True)

notas = pd.DataFrame([
    ["Arquivo", f"{NOME_ORIGINAL} (versão congelada de 04/11/2026; trabalhado em cópia; a planilha original não foi alterada)"],
    ["Coluna", 'Temática Secundária (nome real: "Temática Secundária ", com espaço no fim), separada por ";".'],
    ["Variantes", f"Variantes de grafia aprovadas pela autora em {DATA_DECISAO} aplicadas ({len(MAPA)} grafias unificadas; ver aba Variantes_aplicadas da Solicitação 17). "
     "Não aprovados (mantidos separados): Cantor/Cantora, Campeã/Campeão, Feira/Feirão. Raio/Raio-X unificados por decisão da autora."],
    ["Contagem", "Posts, não células: cada temática conta no máximo uma vez por post. % no eixo pode somar mais de 100% na linha."],
    ["Filtro de temáticas", f"Temáticas com ≥ {MIN_TEM} posts no corpus inteiro (4.266 posts): {len(TEM_OK)} de {len(freq_tem)} rótulos distintos. "
     f"Ficaram de fora {len(tem_fora)} rótulos ({int(tem_fora.sum())} ocorrências). Posts que ficaram sem nenhuma temática após o filtro: {posts_sem_tem} "
     "(continuam no N do eixo, que é o denominador)."
     + (f" {len(tem_fora_base)} temáticas passaram no filtro, mas só ocorrem em eixos excluídos: {', '.join(tem_fora_base)}." if tem_fora_base else "")],
    ["Filtro de eixos", f"Eixos com ≥ {MIN_EIXO} posts: {len(EIXOS)} eixos, {br_int(N)} posts. Ficaram de fora {len(eixos_fora)} eixos ({int(eixos_fora.sum())} posts): "
     + "; ".join(f"{e} ({n})" for e, n in eixos_fora.items()) + "."],
    ["Matriz", f"{len(EIXOS)} eixos × {M.shape[1]} temáticas. Eixos e temáticas ordenados pelo total (empates em ordem alfabética). Linha Total: posts da base com a temática (N) ou % sobre os {br_int(N)} posts da base."],
    ["Lift", f"Lift = (N do par ÷ N da base) ÷ [(N do eixo ÷ N da base) × (N da temática na base ÷ N da base)] = N do par × {br_int(N)} ÷ (N do eixo × N da temática). "
     "Lift 1 = o par ocorre o que se esperaria se eixo e temática fossem independentes; > 1 = mais que o esperado. "
     "% da temática que está no eixo = N do par ÷ N de posts da base com a temática."],
    ["Associacao", f"Pares com ≥ {MIN_PAR} posts ({len(A)} pares), ordenados por lift e por N. * = menos de {POUCOS} posts: lift instável."],
    ["Top_por_eixo", "Para cada eixo: 10 temáticas de maior % no eixo (empates em ordem alfabética) e 10 de maior lift entre pares com ≥ 10 posts."],
    ["Agrupamento", "Vetor de cada eixo = % de posts do eixo com cada temática (todas as temáticas da matriz). Distância do cosseno (1 − cosseno). "
     "Agrupamento hierárquico aglomerativo, ligação média (UPGMA; scipy.cluster.hierarchy.linkage, method='average'). "
     "Corte em k grupos com fcluster(criterion='maxclust'). Silhueta média com a mesma matriz de distâncias (sklearn, metric='precomputed'), k de 2 a 8. "
     "Método determinístico: não há semente. Grupos numerados na ordem do dendrograma."],
    ["Soluções apresentadas", f"Melhor silhueta: k = {K_MELHOR} ({br_dec(S['Silhueta média'].max(), 3)}). Solução de 3 grupos pedida pela autora: silhueta {br_dec(S.set_index('Nº de grupos (k)').loc[3, 'Silhueta média'], 3)}. Escolha a cargo da autora."],
    ["Temáticas características", "Para cada grupo: 10 temáticas de maior % médio nos eixos do grupo (média simples entre os eixos, sem ponderar pelo N)."],
    ["Gráficos", f"(1) Mapa de calor: {len(EIXOS)} eixos × 30 temáticas mais frequentes na base, % no eixo, célula em branco = 0. "
     "(2) Dendrograma com as linhas de corte das soluções apresentadas. "
     f"(3) Rede bipartida: arestas para pares com lift > {LIFT_REDE} e ≥ {N_REDE} posts; espessura proporcional ao N de posts."],
    ["RESSALVA", "Agrupamento e lift são EXPLORATÓRIOS: dependem dos filtros, da codificação livre das temáticas e das variantes aprovadas; "
     "eixos pequenos e pares com poucos posts produzem valores instáveis. Não são testes de hipótese."],
    ["Script", "solicitacao18.py (a Solicitação 18 anterior, Ambiente midiático, passou a usar solicitacao18_ambiente.py)"],
], columns=["Item", "Descrição"])

fora_t = tem_fora.rename_axis("Temática excluída (< 5 posts)").reset_index(name="N de posts")
fora_t = fora_t.sort_values(["N de posts", "Temática excluída (< 5 posts)"], ascending=[False, True])
fora_e = eixos_fora.rename_axis("Eixo excluído (< 30 posts)").reset_index(name="N de posts")

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    mat_n.to_excel(w, sheet_name="Matriz_Eixo_Tematica", index=False)
    mat_p.to_excel(w, sheet_name="Percentual_no_eixo", index=False)
    A.to_excel(w, sheet_name="Associacao", index=False)
    T.to_excel(w, sheet_name="Top_por_eixo", index=False)
    S.to_excel(w, sheet_name="Agrupamento", index=False)
    G.to_excel(w, sheet_name="Agrupamento", index=False, startrow=len(S) + 3)
    eixo_grupos.to_excel(w, sheet_name="Agrupamento", index=False, startrow=len(S) + len(G) + 6)
    C.to_excel(w, sheet_name="Tematicas_por_grupo", index=False)
    Dist.to_excel(w, sheet_name="Distancias_cosseno", index=False)
    fora_e.to_excel(w, sheet_name="Filtros", index=False)
    fora_t.to_excel(w, sheet_name="Filtros", index=False, startcol=4)
    notas.to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
for ws in wb.worksheets:
    for c in ws[1]:
        c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="top")
    if ws.title == "Notas":
        ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 130
        for row in ws.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")
        continue
    hdr = {}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.font.bold is not True and c.row > 1:
                pass
        for c in row:
            if isinstance(c.value, str):
                hdr[c.column] = c.value if (c.row == 1 or c.value in ("Solução", "Eixo", "Nº de grupos (k)")) else hdr.get(c.column, "")
    ws.column_dimensions["A"].width = 26
pct_cols = ("%", "Silhueta", "Distância", "Lift")


def formatar(ws, pct_sheet=False):
    cab = {}
    for row in ws.iter_rows():
        if any(isinstance(c.value, str) and c.value in ("Solução", "Eixo", "Nº de grupos (k)", "Distância do cosseno", "Eixo excluído (< 30 posts)") for c in row[:1]):
            cab = {c.column: str(c.value) for c in row}
            for c in row:
                c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="top")
            continue
        for c in row:
            h = cab.get(c.column, "")
            if isinstance(c.value, float):
                if pct_sheet or h.startswith("%"):
                    c.number_format = "0.0%"
                elif "Lift" in h or "Silhueta" in h or ws.title == "Distancias_cosseno":
                    c.number_format = "0.000" if ("Silhueta" in h or ws.title == "Distancias_cosseno") else "0.00"
                else:
                    c.number_format = "#,##0"
            elif isinstance(c.value, int) and not isinstance(c.value, bool):
                c.number_format = "#,##0"


for nome in wb.sheetnames:
    if nome != "Notas":
        formatar(wb[nome], pct_sheet=(nome == "Percentual_no_eixo"))
for nome in ["Matriz_Eixo_Tematica", "Percentual_no_eixo"]:
    ws = wb[nome]
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    for c in ws[2:ws.max_row]:
        pass
    if nome == "Percentual_no_eixo":
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
            row[0].number_format = "#,##0"
ws = wb["Agrupamento"]
ws.column_dimensions["D"].width = 90
ws.cell(len(S) + 3, 1, "Soluções apresentadas (eixos de cada grupo)").font = Font(bold=True)
ws.cell(len(S) + len(G) + 6, 1, "Grupo de cada eixo e silhueta individual").font = Font(bold=True)
wb["Associacao"].column_dimensions["B"].width = 28
wb["Top_por_eixo"].column_dimensions["B"].width = 24; wb["Top_por_eixo"].column_dimensions["D"].width = 28
wb["Tematicas_por_grupo"].column_dimensions["C"].width = 60; wb["Tematicas_por_grupo"].column_dimensions["E"].width = 26
wb["Filtros"].column_dimensions["E"].width = 34
wb.save(SAIDA)

# ------------------------------------------------------------------ 7. gráficos
# (1) mapa de calor: eixos × 30 temáticas mais frequentes na base
t30 = list(M.columns[:30])
H = P[t30].values * 100
fig, ax = plt.subplots(figsize=(17, 0.42 * len(EIXOS) + 3.2))
cmap = matplotlib.colormaps["cividis_r"].copy(); cmap.set_bad("white")
im = ax.imshow(np.ma.masked_equal(H, 0), cmap=cmap, aspect="auto", vmin=0, vmax=H.max())
for i in range(H.shape[0]):
    for j in range(H.shape[1]):
        if H[i, j] > 0:
            ax.text(j, i, br_dec(H[i, j]), ha="center", va="center", fontsize=6.5,
                    color="white" if H[i, j] > H.max() * 0.55 else "black")
for i in range(H.shape[0] + 1):
    ax.axhline(i - 0.5, color="#e6e6e6", linewidth=0.5)
for j in range(H.shape[1] + 1):
    ax.axvline(j - 0.5, color="#e6e6e6", linewidth=0.5)
ax.set_yticks(range(len(EIXOS))); ax.set_yticklabels([f"{e} (N = {br_int(NE[e])})" for e in EIXOS], fontsize=8.5)
ax.set_xticks(range(len(t30))); ax.set_xticklabels([f"{t} ({br_int(NT[t])})" for t in t30], fontsize=8, rotation=55, ha="right")
for s in ax.spines.values():
    s.set_visible(False)
cb = fig.colorbar(im, ax=ax, fraction=0.02, pad=0.01); cb.set_label("% dos posts do eixo com a temática")
cb.ax.yaxis.set_major_formatter(lambda v_, _: f"{v_:.0f}%")
fig.text(0.99, 0.005, f"Temáticas: 30 mais frequentes na base (N de posts entre parênteses) · eixos com ≥ 30 posts (N = {br_int(N)}) · "
         "célula em branco = 0 · um post pode ter várias temáticas · variantes aprovadas aplicadas", ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(PNG_CALOR, dpi=300, bbox_inches="tight"); plt.close(fig)

# (2) dendrograma
fig, ax = plt.subplots(figsize=(10, 7.5))
dendrogram(Z, labels=[f"{e} ({br_int(NE[e])})" for e in EIXOS], orientation="right", ax=ax,
           color_threshold=0, above_threshold_color="#0072B2", leaf_font_size=9)
estilos = {K_MELHOR: ("#D55E00", "--"), 3: ("#009E73", ":")}
for k in SOLUCOES:
    hs = sorted(Z[:, 2])
    corte = (hs[-k] + hs[-(k - 1)]) / 2                       # altura entre a fusão que deixa k e a que deixa k−1 grupos
    cor, ls = estilos[k] if k in estilos else ("gray", "--")
    rot = f"k = {k}" + (" (melhor silhueta)" if k == K_MELHOR else "")
    ax.axvline(corte, color=cor, linestyle=ls, linewidth=1.4, label=rot)
ax.set_xlabel("Distância do cosseno (ligação média)")
ax.spines[["top", "right"]].set_visible(False)
ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.09), ncol=2, fontsize=9)
ax.xaxis.set_major_formatter(lambda v_, _: f"{v_:.1f}".replace(".", ","))
fig.text(0.99, 0.005, "Eixos com ≥ 30 posts (N entre parênteses) · vetor = % de posts do eixo com cada temática secundária · exploratório",
         ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(PNG_DENDRO, dpi=300, bbox_inches="tight"); plt.close(fig)

# (3) rede bipartida: eixos à esquerda, temáticas à direita
R = A[(A["Lift"] > LIFT_REDE) & (A["N de posts"] >= N_REDE)].copy()
eixos_r = [e for e in EIXOS if e in set(R["Eixo"])]
# temáticas ordenadas pelo eixo de maior N com que se ligam (reduz cruzamentos)
dono = R.sort_values("N de posts", ascending=False).drop_duplicates("Temática").set_index("Temática")["Eixo"]
tems_r = sorted(dono.index, key=lambda t: (eixos_r.index(dono[t]), -R[R["Temática"] == t]["N de posts"].max(), t))
ye = {e: y for e, y in zip(eixos_r, np.linspace(1, 0, len(eixos_r)))}
yt = {t: y for t, y in zip(tems_r, np.linspace(1, 0, len(tems_r)))}
fig, ax = plt.subplots(figsize=(11, max(7, 0.24 * len(tems_r) + 1.5)))
nmax = R["N de posts"].max()
for r in R.itertuples():
    ax.plot([0, 1], [ye[r.Eixo], yt[r.Temática]], color="#0072B2", alpha=0.55,
            linewidth=0.6 + 6 * r._3 / nmax, solid_capstyle="round", zorder=1)
ax.scatter([0] * len(eixos_r), list(ye.values()), s=90, color="#D55E00", zorder=3, edgecolor="white", linewidth=1.5)
ax.scatter([1] * len(tems_r), list(yt.values()), s=40, color="#009E73", zorder=3, edgecolor="white", linewidth=1)
for e, y in ye.items():
    ax.text(-0.03, y, f"{e} ({br_int(NE[e])})", ha="right", va="center", fontsize=9, fontweight="bold")
for t, y in yt.items():
    ax.text(1.03, y, t, ha="left", va="center", fontsize=8)
for n in [10, 50, 100]:
    if n <= nmax:
        ax.plot([], [], color="#0072B2", alpha=0.55, linewidth=0.6 + 6 * n / nmax, label=f"{n} posts")
ax.legend(title="Espessura = N de posts do par", frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.01), ncol=3, fontsize=8)
ax.text(0, 1.04, "Eixo Temático", ha="center", fontsize=9, color="dimgray", transform=ax.transData)
ax.text(1, 1.04, "Temática Secundária", ha="center", fontsize=9, color="dimgray")
ax.set_xlim(-0.6, 1.5); ax.set_ylim(-0.03, 1.07); ax.axis("off")
fig.text(0.99, 0.003, f"Arestas: pares com lift > {LIFT_REDE} e ≥ {N_REDE} posts ({len(R)} pares) · eixos com ≥ 30 posts · exploratório",
         ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(PNG_REDE, dpi=300, bbox_inches="tight"); plt.close(fig)

# ------------------------------------------------------------------ 8. conferência
pd.set_option("display.width", 220); pd.set_option("display.max_colwidth", 140)
print("N base", N, "eixos", len(EIXOS), "temáticas", M.shape[1], "TEM_OK", len(TEM_OK), "distintos", len(freq_tem),
      "fora", len(tem_fora), int(tem_fora.sum()), "posts sem tem", posts_sem_tem, "tem_fora_base", tem_fora_base)
print(eixos_fora.to_string())
print(mat_n.iloc[:10, :8].to_string())
print(S.round(3).to_string())
print(G.round(3).to_string())
print("pares", len(A), "rede", len(R), "eixos na rede", len(eixos_r), "tems", len(tems_r))
print(A.head(15).round(2).to_string())
