"""Solicitação 18 — Os perfis no ambiente midiático de Imperatriz (menções a outros perfis).
Pergunta: com quais veículos e perfis informativos cada perfil se relaciona, pelas menções (@) que faz?
Base: planilha congelada + menções da Solicitação 13 v2 (Ego e Roma: colunas mentions/ da planilha;
Imperatriz Online: Mencoes_Imperatriz_Online.xlsx, aba "Para colar", enviada pela autora).
A TIPOLOGIA DAS CONTAS É PROPOSTA DA ASSISTENTE, feita só pelo nome da conta, e depende de aprovação da autora.
Uso: python3 solicitacao18_ambiente.py copia.xlsx
Gera Solicitacao18_Ambiente_midiatico.xlsx e .png.
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
MENC = "io_mencoes/mencoes_io.xlsx"
SAIDA = "Solicitacao18_Ambiente_midiatico.xlsx"
PNG = "Solicitacao18_Ambiente_midiatico.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
M = [f"mentions/{i}" for i in range(14)]

# ---- Tipologia PROPOSTA (por nome da conta; a confirmar pela autora)
REDE = {  # contas dos próprios perfis analisados
    "egoimperatriz": "Ego",
    "romab2.2": "Roma (a confirmar: perfil do apresentador)",
    "romab2.2reserva": "Roma (a confirmar: \"página reserva\" citada nas descrições)",
    "imperatrizonline": "Imperatriz Online",
    "maismaranhao": "Imperatriz Online (a confirmar: \"outro perfil do Imperatriz Online\", segundo a descrição da linha 2746)",
}
MIDIA_LOCAL = ["secreta.noticias", "secreta.noticia", "jegueonlineitz", "jegueonline", "jegueonlineiz", "manimnoticiasitz",
               "manimnoticias", "blog_do_xibiu", "blog.noticia.da.foto", "tvdifusoraitz", "tvdifusoraitz_", "difusoratv",
               "difusoraon", "portaldifusoranews", "bandeira2.itz", "tvnativaimp", "ideci_vieira_noticias", "itznofoco",
               "itznofoco2", "imperatriznoticias", "imperatrizeregiao", "noticiasnewsitz", "nasruasdeitz", "seliganaitz_",
               "tretas.itz", "portaldoguigui1", "imperaitz"]
MIDIA_REGIONAL = ["deolhoemribeiraozinho", "ribeiraozinho_news_", "joaolisboaonline_", "nasruasdecdm", "davinopolis.noticias",
                  "jornal_de_davinopolis", "portofranco.online", "saomiguelonline_", "senadorlarocque_news",
                  "senadorlarocqueonline", "amarcosnoticias.com.br", "brejao_online", "tvbananal", "bananal_online",
                  "altamiraonline", "axixato_noticias", "buriticupu.online", "estreitonline", "sitio_novo.online",
                  "sitionovo.online", "portal_ulianopolisoficial", "noticiasbomjesusoficial", "slzonline", "blogmaranhao",
                  "sensacionalistaslz", "tv_liberdade", "diegocostatv"]
MIDIA_NACIONAL = ["portalg1", "tvglobo", "sbt", "uoloficial"]
ORGAO = ["14batalhao", "3bbm.itz", "delegaciadamulherimperatriz", "delegaciadehomicidiosimp", "denarcimperatriz",
         "policiacivildeimperatriz", "pmmaoficial", "mpmaoficial", "caemaoficial", "edu.ma.gov.br", "imperatriz.oficial"]
A_CONFIRMAR = {"imperaitz", "diegocostatv", "sensacionalistaslz", "bandeira2.itz"}
CAT = {**{k: "Contas dos três perfis analisados" for k in REDE},
       **{k: "Mídia e perfis informativos de Imperatriz" for k in MIDIA_LOCAL},
       **{k: "Mídia e perfis informativos de outras localidades" for k in MIDIA_REGIONAL},
       **{k: "Mídia nacional" for k in MIDIA_NACIONAL},
       **{k: "Órgão público" for k in ORGAO}}
CATS = ["Mídia e perfis informativos de Imperatriz", "Mídia e perfis informativos de outras localidades", "Mídia nacional",
        "Órgão público", "Contas dos três perfis analisados", "Demais contas (empresas, pessoas etc.; não classificadas)"]

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
assert df["Perfil"].value_counts().to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
nv = pd.read_excel(MENC, sheet_name="Para colar")
assert (nv["url"].astype(str).str.strip().values == df["url "].astype(str).str.strip().values).all()
df["m"] = nv[M].apply(lambda r: sorted({str(x).strip() for x in r if pd.notna(x) and str(x).strip()}), axis=1)
todas = set(x for l in df["m"] for x in l)
for k in CAT:
    assert k in todas, f"conta da tipologia não aparece nas menções: {k}"
cat_post = df["m"].apply(lambda l: {CAT.get(x, CATS[-1]) for x in l})

# ---- 1. Posts que mencionam cada categoria (não exclusivo)
rows = []
for p in PERFIS + ["Total"]:
    k = df["Perfil"] == p if p != "Total" else df["Perfil"].notna()
    n = int(k.sum())
    com = int((df.loc[k, "m"].str.len() > 0).sum())
    r = [p, n, com, com / n]
    for c in CATS:
        x = int(cat_post[k].apply(lambda s: c in s).sum())
        r += [x, x / n]
    rows.append(r)
t_cat = pd.DataFrame(rows, columns=["Perfil", "N de posts", "Posts com alguma menção (N)", "Posts com alguma menção (%)"]
                     + [x for c in CATS for x in (f"{c} (N)", f"{c} (%)")])

# ---- 2. Menções cruzadas entre os três perfis (e às próprias contas)
alvo = {"egoimperatriz": "Ego", "romab2.2": "Roma", "romab2.2reserva": "Roma", "imperatrizonline": "Imperatriz Online",
        "maismaranhao": "Imperatriz Online"}
rows = []
for p in PERFIS:
    b = df[df["Perfil"] == p]
    r = [p, len(b)]
    for q in PERFIS:
        x = int(b["m"].apply(lambda l: any(alvo.get(a) == q for a in l)).sum())
        r += [x, x / len(b)]
    rows.append(r)
t_cruz = pd.DataFrame(rows, columns=["Perfil que menciona", "N de posts"] + [x for q in PERFIS for x in (f"Menciona {q} (N)", f"Menciona {q} (%)")])
det = []
for p in PERFIS:
    b = df[df["Perfil"] == p]
    for a in alvo:
        x = int(b["m"].apply(lambda l: a in l).sum())
        if x:
            det.append([p, a, alvo[a], x, x / len(b)])
t_cruz_det = pd.DataFrame(det, columns=["Perfil que menciona", "Conta mencionada", "Perfil a que pertence (proposta)", "N de posts", "% dos posts do perfil"])

# ---- 3. Contas de mídia mais mencionadas por perfil (Imperatriz + outras localidades + nacional)
MIDIA = set(MIDIA_LOCAL) | set(MIDIA_REGIONAL) | set(MIDIA_NACIONAL)
rows = []
for p in PERFIS:
    b = df[df["Perfil"] == p]
    c = b["m"].explode().dropna()
    c = c[c.isin(MIDIA)].value_counts().rename_axis("a").reset_index(name="n").sort_values(["n", "a"], ascending=[False, True])
    for a, n in zip(c["a"], c["n"]):
        rows.append([p, a + (" *" if n < 10 else ""), CAT[a], int(n), n / len(b)])
t_midia = pd.DataFrame(rows, columns=["Perfil", "Conta mencionada", "Categoria (proposta)", "N de posts", "% dos posts do perfil"])

# ---- 4. Eixos dos posts que mencionam mídia de Imperatriz (por perfil)
EIXO = "Eixo Temático"
rows = []
for p in PERFIS:
    b = df[df["Perfil"] == p]
    k = cat_post[b.index].apply(lambda s: CATS[0] in s)
    e = b.loc[k, EIXO].value_counts()
    for i, (nome, n) in enumerate(e.head(5).items()):
        rows.append([p, int(k.sum()), i + 1, nome, int(n), n / k.sum()])
t_eixo = pd.DataFrame(rows, columns=["Perfil", "Posts que mencionam mídia de Imperatriz (N)", "Ordem", "Eixo", "N", "% desses posts"])

# ---- 5. Indicadores já confirmados, recalculados aqui para o mesmo quadro (Collab e Mediação)
rows = []
for p in PERFIS + ["Total"]:
    b = df if p == "Total" else df[df["Perfil"] == p]
    rows.append([p, len(b), int((b["Collab"].astype(str).str.strip() == "Sim").sum()),
                 (b["Collab"].astype(str).str.strip() == "Sim").mean(),
                 int((b["Mediação da Informação"].astype(str).str.strip() == "Factual").sum()),
                 (b["Mediação da Informação"].astype(str).str.strip() == "Factual").mean(),
                 int((b["Mediação da Informação"].astype(str).str.strip() == "Publicitária").sum()),
                 (b["Mediação da Informação"].astype(str).str.strip() == "Publicitária").mean()])
t_ind = pd.DataFrame(rows, columns=["Perfil", "N de posts", "Collab Sim (N)", "Collab Sim (%)", "Mediação Factual (N)",
                                    "Mediação Factual (%)", "Mediação Publicitária (N)", "Mediação Publicitária (%)"])

# ---- Tipologia para aprovação
tip = pd.DataFrame([[a, CAT[a], REDE.get(a, ""), "Sim" if a in A_CONFIRMAR or "a confirmar" in REDE.get(a, "") else "",
                     *[int(df[df["Perfil"] == p]["m"].apply(lambda l: a in l).sum()) for p in PERFIS], ""]
                    for a in CAT], columns=["Conta", "Categoria (proposta da assistente)", "Observação", "Dúvida da assistente",
                                            "Ego (N)", "Roma (N)", "Imperatriz Online (N)", "Decisão da autora"])

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia). Menções: as mesmas da Solicitação 13 v2 (Ego e Roma da planilha; Imperatriz Online do arquivo Mencoes_Imperatriz_Online.xlsx, aba \"Para colar\", enviado pela autora; o script confere que as URLs batem linha a linha)."],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721). Exclusões: nenhuma. Cada conta conta uma vez por post."],
    ["TIPOLOGIA = PROPOSTA", "As categorias das contas foram atribuídas PELA ASSISTENTE, só pelo nome da conta, sem abrir os perfis. Apenas contas com nome de veículo, perfil de notícias, emissora, órgão público ou das próprias contas analisadas foram classificadas; as demais (empresas, pessoas, artistas) ficaram em \"Demais contas\". Há contas marcadas com dúvida. Tudo depende de aprovação (aba Tipologia_para_aprovacao)."],
    ["Contas dos três perfis", "A categoria reúne as contas de Ego, Roma e Imperatriz Online, sem distinguir se o perfil menciona a si mesmo ou a outro; essa distinção está na aba Mencoes_cruzadas (diagonal = menção à própria conta). romab2.2 e romab2.2reserva foram atribuídas ao Roma, e maismaranhao ao Imperatriz Online, por inferência (descrições da Solicitação 16 citam a \"página reserva\" do Roma e o \"outro perfil do Imperatriz Online, o Mais Maranhão\"). A confirmar."],
    ["Fórmula do %", "% = posts do perfil que mencionam pelo menos uma conta da categoria ÷ N de posts do perfil × 100. Categorias não são exclusivas: um post pode mencionar contas de várias categorias, por isso as linhas não somam 100%."],
    ["Menções ≠ relação", "Menção (@) é um indício de relação (crédito de imagem, parceria, repost, citação), não a prova do tipo de relação. Collab é outra variável (aba Indicadores)."],
    ["Indicadores", "Collab e Mediação da Informação recalculados da planilha; devem coincidir com as Solicitações 12 e 3, já confirmadas."],
    ["*", "Menos de 10 posts."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a menções, Collab ou Mediação."],
    ["Script", "solicitacao18_ambiente.py (renomeado em 07/11/2026, quando solicitacao18.py passou a ser Eixo × Temática Secundária)"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t_cat.to_excel(w, sheet_name="Mencoes_por_categoria", index=False)
    t_cruz.to_excel(w, sheet_name="Mencoes_cruzadas", index=False)
    t_cruz_det.to_excel(w, sheet_name="Mencoes_cruzadas", index=False, startrow=len(t_cruz) + 3)
    t_midia.to_excel(w, sheet_name="Contas_de_midia_por_perfil", index=False)
    t_eixo.to_excel(w, sheet_name="Eixos_mencao_midia_local", index=False)
    t_ind.to_excel(w, sheet_name="Indicadores", index=False)
    tip.to_excel(w, sheet_name="Tipologia_para_aprovacao", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)
wb = load_workbook(SAIDA)
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            h = str(ws.cell(1, c.column).value) if c.row > 1 else ""
            if isinstance(c.value, float) and ("%" in h or (ws.title == "Mencoes_cruzadas")):
                c.number_format = "0.0%"
            elif isinstance(c.value, (int, np.integer)):
                c.number_format = "#,##0"
    for c in ws[1]:
        c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="top")
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18
ws = wb["Tipologia_para_aprovacao"]
ws.column_dimensions["A"].width = 26; ws.column_dimensions["B"].width = 44; ws.column_dimensions["C"].width = 50
for r in range(2, ws.max_row + 1):
    ws.cell(r, 8).fill = PatternFill("solid", fgColor="FFF2CC")
ws = wb["Notas"]; ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico
cats_g = CATS[:5]
cores = ["#0072B2", "#56B4E9", "#CC79A7", "#009E73", "#E69F00"]
fig, ax = plt.subplots(figsize=(11, 4.2))
x = np.arange(len(PERFIS)); wd = 0.16
for i, (c, cor) in enumerate(zip(cats_g, cores)):
    v = t_cat.set_index("Perfil").loc[PERFIS, f"{c} (%)"].values * 100
    bars = ax.bar(x + (i - 2) * wd, v, wd, color=cor, label=c)
    for bx, val in zip(bars, v):
        ax.text(bx.get_x() + bx.get_width() / 2, val + 0.8, f"{val:.1f}".replace(".", ","), ha="center", fontsize=7)
ax.set_xticks(x, [f"{p}\n(N = {n:,})".replace(",", ".") for p, n in zip(PERFIS, t_cat.set_index("Perfil").loc[PERFIS, "N de posts"])])
ax.set_ylabel("% dos posts do perfil que mencionam\npelo menos uma conta da categoria")
ax.legend(fontsize=8, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=3)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
fig.text(0.01, 0.005, "Tipologia das contas proposta pela assistente, pelo nome da conta, a confirmar pela autora. Categorias não exclusivas.", fontsize=7, color="#555")
fig.tight_layout(); fig.savefig(PNG, dpi=300, bbox_inches="tight")

pd.set_option("display.width", 250)
print(t_cat.iloc[:, [0, 1, 3, 5, 7, 9, 11, 13, 15]].to_string()); print(t_cruz.to_string()); print(t_cruz_det.to_string())
print(t_midia.groupby("Perfil").head(8).to_string()); print(t_eixo.to_string()); print(t_ind.to_string())
