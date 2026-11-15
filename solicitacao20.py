"""
Solicitação 20 — Nuvem de palavras da Linguagem Regional.

Uso:  python3 -I solicitacao20.py copia.xlsx

Regras (definidas pela pesquisadora):
- Unidade = cada termo tal como codificado entre "; " (não se divide por espaços).
- Conta-se o número de POSTS que têm o termo (no máximo uma vez por post).
- Grafia mantida como codificada; só se aplicam as uniões aprovadas (UNIOES).
- "Não há" é excluído.
- Tamanho da palavra proporcional à RAIZ QUADRADA do número de posts.
- Posicionamento das palavras: semente fixa SEMENTE (reprodutível).

Saídas: Solicitacao20_Linguagem_regional.xlsx e quatro PNG (300 dpi).
A planilha original não é alterada; o script lê a cópia e aplica as
correções aprovadas da I18 (correcoes_planilha.aplicar).
"""
import os
import sys
import math
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from wordcloud import WordCloud
import correcoes_planilha

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
COL = "Linguagem Regional"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
SEMENTE = 20          # semente do posicionamento das palavras (random_state do WordCloud)
FONTE = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"

# Uniões aprovadas pela autora em 08/11/2026 (todas as demais: manter separado).
# A forma final foi escolhida pela análise (termo contido na expressão) e pode ser trocada aqui.
UNIOES = {
    "Óia só": "Oia",
    "E ai galera": "E aí",
}

# Paleta Okabe-Ito (segura para daltonismo)
COR = {"Ego": "#0072B2", "Roma": "#D55E00", "Imperatriz Online": "#009E73",
       "Mais de um perfil": "#000000"}


def termos_do_post(celula):
    """Divide a célula em "; ", aplica as uniões aprovadas e remove repetições no post."""
    if not isinstance(celula, str) or celula.strip() in ("", "Não há"):
        return []
    ts = [t.strip() for t in celula.split(";") if t.strip()]
    ts = [UNIOES.get(t, t) for t in ts]
    return list(dict.fromkeys(ts))      # cada termo no máximo uma vez por post


# ---------------------------------------------------------------- dados
df = correcoes_planilha.aplicar(pd.read_excel(ARQ))
assert len(df) == 4266
df["termos"] = df[COL].apply(termos_do_post)
N_PERFIL = df["Perfil"].value_counts().to_dict()
com_lr = df[df["termos"].str.len() > 0]

cont = defaultdict(Counter)            # cont[perfil][termo] = N de posts
for _, r in com_lr.iterrows():
    for t in r["termos"]:
        cont[r["Perfil"]][t] += 1
total = Counter()
for p in PERFIS:
    total.update(cont[p])

# ---------------------------------------------------------------- tabela Termos
ter = pd.DataFrame([{"Termo": t,
                     "Ego (N)": cont["Ego"][t], "Roma (N)": cont["Roma"][t],
                     "Imperatriz Online (N)": cont["Imperatriz Online"][t],
                     "Total (N de posts)": n} for t, n in total.items()])
ter["Perfis em que aparece"] = (ter[["Ego (N)", "Roma (N)", "Imperatriz Online (N)"]] > 0).sum(axis=1)
ter["Aparece uma única vez"] = ter["Total (N de posts)"].eq(1).map({True: "sim", False: ""})
ter = ter.sort_values(["Total (N de posts)", "Termo"], ascending=[False, True],
                      key=lambda c: c.str.casefold() if c.dtype == object else c).reset_index(drop=True)
assert ter[["Ego (N)", "Roma (N)", "Imperatriz Online (N)"]].sum().sum() == ter["Total (N de posts)"].sum()
ter_tot = pd.DataFrame([{"Termo": "Total", **{c: int(ter[c].sum()) for c in
                         ["Ego (N)", "Roma (N)", "Imperatriz Online (N)", "Total (N de posts)"]},
                         "Perfis em que aparece": None, "Aparece uma única vez": int(ter["Aparece uma única vez"].eq("sim").sum())}])

# ---------------------------------------------------------------- tabela Resumo
def resumo(nome, sub, c):
    top = sorted(c.items(), key=lambda x: (-x[1], x[0].casefold()))[:5]
    return {"Perfil": nome, "Posts do perfil": len(sub) if nome != "Total" else 4266,
            "Posts com linguagem regional (N)": int((sub["termos"].str.len() > 0).sum()),
            "Posts com linguagem regional (% do perfil)": (sub["termos"].str.len() > 0).mean(),
            "Termos distintos": len(c),
            "Termos que aparecem uma só vez": sum(1 for v in c.values() if v == 1),
            "5 termos mais frequentes (N)": "; ".join(f"{t} ({n})" for t, n in top)}

res = pd.DataFrame([resumo(p, df[df["Perfil"] == p], cont[p]) for p in PERFIS] +
                   [resumo("Total", df, total)])
n_multi = int((ter["Perfis em que aparece"] > 1).sum())

# ---------------------------------------------------------------- nuvens
def origem(t):
    ps = [p for p in PERFIS if cont[p][t] > 0]
    return ps[0] if len(ps) == 1 else "Mais de um perfil"


def nuvem(freq, arq, cor_fn, rodape, legenda=None):
    """Tamanho ∝ √(N de posts): passa-se √N ao WordCloud com relative_scaling=1."""
    pesos = {t: math.sqrt(n) for t, n in freq.items()}
    # Se uma palavra não cabe, o WordCloud reduz a fonte dela e quebra a proporção.
    # Por isso o tamanho máximo é reduzido aos poucos até o desvio da proporção ficar < 3%.
    mfs = 260
    while True:
        wc = WordCloud(width=3000, height=1700, background_color="white", font_path=FONTE,
                       prefer_horizontal=1.0, relative_scaling=1.0, max_font_size=mfs,
                       min_font_size=8, max_words=len(pesos), margin=14,
                       collocations=False, random_state=SEMENTE)
        wc.generate_from_frequencies(pesos)
        colocados = {w for (w, _), *_ in wc.layout_}
        razoes = [fs / math.sqrt(freq[w]) for (w, _), fs, *_ in wc.layout_]
        desvio = (max(razoes) - min(razoes)) / max(razoes)
        if colocados == set(pesos) and desvio < 0.03:
            break
        mfs = int(mfs * 0.92)
    wc.recolor(color_func=lambda word, **k: cor_fn(word), random_state=SEMENTE)
    fig = plt.figure(figsize=(10, 6.2))
    ax = fig.add_axes([0, 0.12, 1, 0.88]); ax.imshow(wc.to_array(), interpolation="bilinear"); ax.axis("off")
    fig.text(0.5, 0.018, rodape, fontsize=8.5, color="#333333", ha="center")
    if legenda:
        fig.legend(handles=legenda, loc="lower center", bbox_to_anchor=(0.5, 0.045), ncol=4,
                   fontsize=8, frameon=False, handlelength=1.2)
    fig.savefig(arq, dpi=300); plt.close(fig)
    return desvio


desvios = {}
leg = [Patch(color=COR[p], label=f"{p} ({sum(1 for t in total if origem(t) == p)})") for p in PERFIS] + \
      [Patch(color=COR["Mais de um perfil"], label=f"Mais de um perfil ({n_multi})")]
desvios["Total"] = nuvem(total, "Solicitacao20_Nuvem_total.png", lambda w: COR[origem(w)],
    f"Total: {len(com_lr)} posts com linguagem regional (de 4.266); {len(total)} termos distintos. "
    f"Tamanho ∝ √(N de posts). Cor = perfil de origem; entre parênteses, N de termos.", leg)
for p, arq in zip(PERFIS, ["Solicitacao20_Nuvem_Ego.png", "Solicitacao20_Nuvem_Roma.png",
                           "Solicitacao20_Nuvem_Imperatriz_Online.png"]):
    r = res[res["Perfil"] == p].iloc[0]
    desvios[p] = nuvem(dict(cont[p]), arq, lambda w, p=p: COR[p],
        f"{p}: {r['Posts com linguagem regional (N)']} posts com linguagem regional "
        f"(de {r['Posts do perfil']:,}".replace(",", ".") + f" posts do perfil); {r['Termos distintos']} termos distintos. "
        "Tamanho ∝ √(N de posts).")

# ---------------------------------------------------------------- Excel
var = pd.read_excel("Solicitacao20_Variantes_para_aprovacao.xlsx", sheet_name="Variantes_para_aprovacao")
aprov = {("Oia", "Óia só"), ("E aí", "E ai galera")}
var["Decisão da autora"] = [("unir" if (a, b) in aprov else "manter separado") for a, b in zip(var["Termo A"], var["Termo B"])]

notas = pd.DataFrame([
    ["Base", "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx (cópia; original não alterada), com as correções aprovadas na I18. 4.266 posts."],
    ["Exclusões", "\"Não há\": 4.027 posts. Base da nuvem: 239 posts com linguagem regional."],
    ["Unidade", "Cada termo entre \"; \" é uma unidade, sem divisão por espaços; grafia mantida como codificada."],
    ["Contagem", "N de posts que têm o termo (no máximo uma vez por post). Denominador do % no Resumo: posts do perfil."],
    ["Uniões aprovadas (08/11/2026)", "\"Óia só\" → \"Oia\"; \"E ai galera\" → \"E aí\". As outras 38 variantes listadas: manter separado (aba Variantes_decididas)."],
    ["Forma final das uniões", "Escolhida pela análise: o termo contido na expressão (Oia; E aí). Pode ser trocada no dicionário UNIOES do script."],
    ["Efeito das uniões", f"Oia: {total['Oia']} posts; E aí: {total['E aí']} posts. Termos distintos: 190 antes, {len(total)} depois."],
    ["Escala", "Tamanho da fonte proporcional à raiz quadrada do N de posts (WordCloud com pesos √N e relative_scaling = 1). "
               "Ex.: Valamedeus (68) fica √68/√1 ≈ 8,2 vezes maior que um termo com 1 post (e não 68 vezes). "
               "Tamanho máximo reduzido até todos os termos caberem sem perder a proporção. Desvio máximo da proporcionalidade (arredondamento do tamanho da fonte): " +
               "; ".join(f"{k} {v*100:.1f}%".replace(".", ",") for k, v in desvios.items()) + "."],
    ["Comparação entre nuvens", "Cada nuvem é dimensionada para ocupar o espaço: a proporção √N vale dentro de cada nuvem, mas o tamanho absoluto de um termo não é comparável entre nuvens."],
    ["Posicionamento", f"Aleatório com semente fixa (random_state = {SEMENTE}); todas as palavras na horizontal; todos os termos foram posicionados (verificado no script)."],
    ["Cores", "Paleta Okabe-Ito: Ego azul, Roma vermelhão, Imperatriz Online verde, termo em mais de um perfil preto. "
              f"Termos em mais de um perfil: {n_multi}."],
    ["Sinalização", "\"Aparece uma única vez\" = total de 1 post no corpus. No Resumo, \"Termos que aparecem uma só vez\" = termos com 1 post no perfil. Empates no top 5: ordem alfabética."],
    ["Pontos a conferir", "\"duVídeodó\" parece efeito de substituição automática de texto; \"Alimentação\", \"Sim\", \"Pressão\", \"Varejo\", \"top\" e \"0800\" estão codificados como linguagem regional. Mantidos como estão."],
    ["Script", "solicitacao20.py (o script anterior com este número, do gráfico Eixo × Mediação sem Publicidade, foi renomeado para solicitacao20_mediacao.py)."],
], columns=["Item", "Descrição"])

saida = "Solicitacao20_Linguagem_regional.xlsx"
with pd.ExcelWriter(saida, engine="openpyxl") as w:
    pd.concat([ter, ter_tot]).to_excel(w, "Termos", index=False)
    res.to_excel(w, "Resumo", index=False)
    var.to_excel(w, "Variantes_decididas", index=False)
    notas.to_excel(w, "Notas", index=False)

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
wb = load_workbook(saida)
larg = {"Termos": [40, 10, 10, 20, 18, 20, 20], "Resumo": [20, 14, 18, 20, 14, 18, 70],
        "Variantes_decididas": [7, 22, 10, 30, 10, 48, 34, 40, 20], "Notas": [28, 130]}
for ws in wb:
    for i, l in enumerate(larg[ws.title]):
        ws.column_dimensions[chr(65 + i)].width = l
    for c in ws[1]:
        c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9E2F3")
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
wb["Termos"].cell(wb["Termos"].max_row, 1).font = Font(bold=True)
for r in range(2, wb["Resumo"].max_row + 1):
    wb["Resumo"].cell(r, 4).number_format = "0.0%"
wb.save(saida)

print(ter.head(10).to_string(index=False))
print(res.to_string(index=False))
print("termos em mais de um perfil:", n_multi, "| desvios:", desvios)
