"""Solicitação 19 — Identidade visual de cada perfil (Parte 1: perfil visual em números; gráficos).
A Parte 2 (leitura das 1.605 descrições) é feita por leitura integral, perfil a perfil; as evidências ficam em
s19_evidencias.json e são conferidas aqui (cada trecho tem de ser substring literal da descrição corrigida).

Uso: python3 -I solicitacao19.py <planilha.xlsx>

Regras: cópia da planilha (original não alterada); correções aprovadas aplicadas via correcoes_planilha.aplicar;
conta posts; Regime de Visibilidade é multirrótulo (separado por ";", cada rótulo uma vez por post);
denominador = N de posts do perfil (ou do eixo/mediação dentro do perfil).
O nome antigo "Solicitação 19" (memorando, memorando_s19.js) não usa este script.
"""
import json
import os
import sys

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)                         # python -I não inclui a pasta do script
from correcoes_planilha import aplicar

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao19_Identidade_visual.xlsx"
EVID_JSON, EVID_XLSX = "s19_evidencias.json", "Solicitacao19_Evidencias.xlsx"
PNG_A, PNG_B, PNG_C = ("Solicitacao19a_Recursos_por_perfil.png", "Solicitacao19b_Impacto_visual_por_perfil.png",
                       "Solicitacao19c_Recurso_x_Eixo_por_perfil.png")
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
REC, IMP, TIPO, REG, MED, EIXO, DESC = ("Recursos Audiovisuais", "Impacto Visual", "Tipo de Conteúdo",
                                        "Regime de Visibilidade", "Mediação da Informação", "Eixo Temático", "Descrição da Imagem")
IMPS = ["Alto", "Médio", "Baixo"]
TIPOS = ["Vídeo", "Image", "Sidecar"]
ROT_TIPO = {"Vídeo": "Vídeo", "Image": "Image (imagem única)", "Sidecar": "Sidecar (carrossel)"}
EMERG_REC = {"Vídeo sem som", "Desenho"}
EMERG_REG = {"Criminoso", "Animal", "Bens materiais", "Candidatos", "Outro perfil"}
br_int = lambda v: f"{int(v):,}".replace(",", ".")
br_pct = lambda v: f"{v:.1f}%".replace(".", ",")

# ------------------------------------------------------------------ leitura
df = aplicar(pd.read_excel(ARQ))
assert df.shape == (4266, 65)
df["Linha"] = df.index + 2
for c in [REC, IMP, TIPO, MED, EIXO, DESC]:
    df[c] = df[c].astype(str).str.strip()
assert set(df[IMP]) == set(IMPS) and set(df[TIPO]) == set(TIPOS)
assert not df[REC].str.contains(";").any()
df["reg"] = df[REG].astype(str).str.split(";").apply(lambda l: list(dict.fromkeys(x.strip() for x in l if x.strip())))
NP = df["Perfil"].value_counts().reindex(PERFIS)
assert NP.tolist() == [414, 1131, 2721]
REC_ORD = df[REC].value_counts().index.tolist()
fe = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
EIXOS10 = fe["e"].iloc[:10].tolist()
MEDS = ["Factual", "Publicitária", "De Entretenimento", "Institucional", "Opinativa"]
marca_rec = lambda r: r + (" †" if r in EMERG_REC else "")


def dist(col, ordem, marca=lambda x: x):
    """N e % por perfil (denominador = posts do perfil) + Total."""
    out = pd.DataFrame({"Categoria": [marca(o) for o in ordem]})
    for p in PERFIS + ["Total"]:
        d = df if p == "Total" else df[df["Perfil"] == p]
        n = d[col].value_counts().reindex(ordem, fill_value=0)
        out[f"{p} (N)"] = n.values
        out[f"{p} (%)"] = n.values / len(d)
    tot = {"Categoria": "Total"}
    for p in PERFIS + ["Total"]:
        tot[f"{p} (N)"] = int(out[f"{p} (N)"].sum()); tot[f"{p} (%)"] = out[f"{p} (%)"].sum()
    out = pd.concat([out, pd.DataFrame([tot])], ignore_index=True)
    for p in PERFIS:     # * para categorias com menos de 10 posts no perfil (marcadas na própria célula de N? não: coluna de aviso)
        out[f"{p} (< 10)"] = ["*" if (0 < n < 10 and c != "Total") else "" for n, c in zip(out[f"{p} (N)"], out["Categoria"])]
    assert out.loc[out.index[-1], "Total (N)"] == 4266
    return out


t_rec = dist(REC, REC_ORD, marca_rec)
t_tipo = dist(TIPO, TIPOS, lambda x: ROT_TIPO[x])
t_imp = dist(IMP, IMPS)

# Recurso × Impacto: 10 combinações mais frequentes por perfil
comb = []
for p in PERFIS:
    d = df[df["Perfil"] == p]
    c = d.groupby([REC, IMP]).size().reset_index(name="N de posts")
    c = c.sort_values(["N de posts", REC, IMP], ascending=[False, True, True])
    corte = c["N de posts"].iloc[9] if len(c) >= 10 else 0
    top = c.head(10).copy()
    top.insert(0, "Posição", range(1, len(top) + 1)); top.insert(0, "Perfil", p)
    top[REC] = top[REC].map(marca_rec)
    top["% dos posts do perfil"] = top["N de posts"] / len(d)
    top["< 10 posts"] = np.where(top["N de posts"] < 10, "*", "")
    top["Empatados com o 10º fora da lista"] = int((c["N de posts"] == corte).sum() - (top["N de posts"] == corte).sum())
    comb.append(top)
t_comb = pd.concat(comb, ignore_index=True)


def rec_por(col, cats):
    """Recursos dentro de cada categoria (eixo ou mediação), por perfil: N e % do N da categoria no perfil."""
    linhas = []
    for p in PERFIS:
        d = df[df["Perfil"] == p]
        for c in cats:
            dc = d[d[col] == c]
            n = dc[REC].value_counts().reindex(REC_ORD, fill_value=0)
            lin = {"Perfil": p, col: c + (" *" if 0 < len(dc) < 10 else ""), "N da categoria no perfil": len(dc)}
            for r in REC_ORD:
                lin[f"{marca_rec(r)} (N)"] = int(n[r])
            for r in REC_ORD:
                lin[f"{marca_rec(r)} (%)"] = n[r] / len(dc) if len(dc) else np.nan
            linhas.append(lin)
    return pd.DataFrame(linhas)


t_eixo = rec_por(EIXO, EIXOS10)
t_med = rec_por(MED, MEDS)

# Regime de Visibilidade (multirrótulo): % de posts do perfil com cada rótulo
ex = df[["Perfil", "reg"]].explode("reg").dropna()
REGS = ex["reg"].value_counts().index.tolist()
t_reg = pd.DataFrame({"Regime de Visibilidade": [r + (" †" if r in EMERG_REG else "") for r in REGS]})
for p in PERFIS + ["Total"]:
    e = ex if p == "Total" else ex[ex["Perfil"] == p]
    n = e["reg"].value_counts().reindex(REGS, fill_value=0)
    den = 4266 if p == "Total" else NP[p]
    t_reg[f"{p} (N)"] = n.values
    t_reg[f"{p} (% dos posts)"] = n.values / den
nrot = df.groupby("Perfil")["reg"].apply(lambda s: s.apply(len).mean()).reindex(PERFIS)

# Resumo
res = []
for p in PERFIS:
    d = df[df["Perfil"] == p]
    def top1(col):
        v = d[col].value_counts()
        return v.index[0], int(v.iloc[0]), v.iloc[0] / len(d)
    r1, t1, i1 = top1(REC), top1(TIPO), top1(IMP)
    res.append({"Perfil": p, "N": len(d),
                "Recurso predominante": marca_rec(r1[0]), "Recurso (N)": r1[1], "Recurso (%)": r1[2],
                "Tipo de conteúdo predominante": ROT_TIPO[t1[0]], "Tipo (N)": t1[1], "Tipo (%)": t1[2],
                "Impacto visual predominante": i1[0], "Impacto (N)": i1[1], "Impacto (%)": i1[2],
                "Média de rótulos de Regime por post": nrot[p]})
t_res = pd.DataFrame(res)

# ------------------------------------------------------------------ evidências (Parte 2)
desc = df.set_index("Linha")[DESC]
evid = json.load(open(EVID_JSON, encoding="utf8")) if os.path.exists(EVID_JSON) else []
for e in evid:
    assert e["Trecho"] in desc[e["Linha"]], (e["Linha"], e["Trecho"])          # trecho literal
    assert df.loc[df["Linha"] == e["Linha"], "Perfil"].item() == e["Perfil"]
t_ev = pd.DataFrame(evid, columns=["Linha", "Perfil", "Código", "Padrão", "Trecho"])
recorte = df[desc.reindex(df["Linha"]).str.lower().ne("não analisado").values]
NREC = recorte["Perfil"].value_counts().reindex(PERFIS)
assert NREC.tolist() == [298, 551, 756]

# ------------------------------------------------------------------ planilha
lidos = sorted(set(t_ev["Perfil"]), key=PERFIS.index) if len(t_ev) else []
notas = pd.DataFrame([
    ["Arquivo", f"{NOME_ORIGINAL} (versão congelada de 04/11/2026; trabalhado em cópia; correções aprovadas aplicadas só no script)"],
    ["N", "Parte 1: 4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721). Denominador de cada % = N de posts do perfil "
     "(ou do eixo / da mediação dentro do perfil). Parte 2: recorte de 1.605 descrições (Ego 298; Roma 551; Imperatriz Online 756), lidas integralmente."],
    ["Exclusões", "Nenhuma na Parte 1 (as variáveis estão preenchidas em todos os posts). Na Parte 2, os 2.661 posts com descrição \"Não analisado\"."],
    ["Recursos Audiovisuais", "Rótulo único por post. † Vídeo sem som e Desenho são categorias emergentes (resultado indicativo)."],
    ["Tipo de Conteúdo", "Metadado do programa de coleta: Vídeo, Image (imagem única), Sidecar (carrossel)."],
    ["Regime de Visibilidade", "Multirrótulo (separado por \";\"): % = posts do perfil com o rótulo; a coluna soma mais de 100%. "
     "† categorias emergentes (Criminoso, Animal, Bens materiais, Candidatos, Outro perfil): aplicadas de forma desigual entre lotes; resultado indicativo."],
    ["Recursos por eixo / mediação", "10 eixos mais frequentes do corpus (os mesmos das Solicitações 4 e 6) e as 5 mediações. % = posts da categoria no perfil com o recurso. "
     "* = categoria com menos de 10 posts no perfil: percentuais instáveis."],
    ["Asterisco (*)", "Categorias/combinações com menos de 10 posts no perfil."],
    ["Limites", "A pesquisadora (e a assistente) não vê as imagens nesta análise: cor, tipografia e layout só entram na Parte 2 quando as descrições os mencionam. "
     "As descrições foram escritas pela codificação da pesquisa e variam em estilo entre lotes."],
    ["Parte 2 — perfis lidos", ", ".join(lidos) if lidos else "nenhum"],
    ["Script", "solicitacao19.py"],
], columns=["Item", "Descrição"])

abas = {"Recursos": t_rec, "Tipo_de_conteudo": t_tipo, "Impacto_visual": t_imp, "Recurso_x_Impacto_top10": t_comb,
        "Recurso_por_Eixo": t_eixo, "Recurso_por_Mediacao": t_med, "Regime_de_visibilidade": t_reg, "Resumo": t_res, "Notas": notas}
with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    for k, v in abas.items():
        v.to_excel(w, sheet_name=k, index=False)
wb = load_workbook(SAIDA)
for ws in wb.worksheets:
    cab = {c.column: str(c.value) for c in ws[1]}
    for c in ws[1]:
        c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 26
    if ws.title == "Notas":
        ws.column_dimensions["B"].width = 130
        for row in ws.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")
        continue
    for row in ws.iter_rows(min_row=2):
        for c in row:
            h = cab.get(c.column, "")
            if isinstance(c.value, float):
                c.number_format = "0.0%" if "%" in h else "0.00"
            elif isinstance(c.value, int):
                c.number_format = "#,##0"
    if ws.cell(ws.max_row, 1).value == "Total":
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
wb.save(SAIDA)

if len(t_ev):
    with pd.ExcelWriter(EVID_XLSX, engine="openpyxl") as w:
        t_ev.to_excel(w, sheet_name="Evidencias", index=False)
        pd.DataFrame([["Regra", "Cada trecho é cópia literal da descrição (conferida pelo script). Padrões recorrentes de composição e organização visual, "
                       "com casos que fogem ao padrão marcados no código com o sufixo X."],
                      ["Perfis lidos", ", ".join(lidos)],
                      ["N lido", "; ".join(f"{p} {NREC[p]} descrições" for p in lidos)]], columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)
    wb = load_workbook(EVID_XLSX); ws = wb["Evidencias"]
    for c in ws[1]:
        c.font = Font(bold=True)
    for col, wd in zip("ABCDE", [8, 18, 8, 50, 110]):
        ws.column_dimensions[col].width = wd
    for row in ws.iter_rows(min_row=2):
        row[4].alignment = Alignment(wrap_text=True, vertical="top"); row[3].alignment = Alignment(wrap_text=True, vertical="top")
    wb["Notas"].column_dimensions["B"].width = 120
    wb.save(EVID_XLSX)

# ------------------------------------------------------------------ gráficos
OKABE = ["#0072B2", "#E69F00", "#009E73", "#CC79A7", "#56B4E9", "#D55E00", "#F0E442", "#000000"]
HACH = ["", "//", "..", "xx", "\\\\", "--", "++", "oo"]


def barras100(tab, cats, nome, legenda, png):
    """Barras horizontais 100% por perfil; valor nos segmentos ≥ 5%."""
    fig, ax = plt.subplots(figsize=(11, 4.2))
    y = np.arange(len(PERFIS))[::-1]
    esq = np.zeros(len(PERFIS))
    for k, c in enumerate(cats):
        lin = tab[tab["Categoria"] == c]
        v = np.array([lin[f"{p} (%)"].item() * 100 for p in PERFIS])
        ax.barh(y, v, left=esq, color=OKABE[k % 8], hatch=HACH[k % 8], edgecolor="white", linewidth=0.8, label=c)
        for yy, vv, ll in zip(y, v, esq):
            if vv >= 5:
                ax.text(ll + vv / 2, yy, br_pct(vv).replace("%", ""), ha="center", va="center", fontsize=8, color="black",
                        bbox=dict(facecolor="white", edgecolor="none", alpha=0.85, pad=1.2))
        esq += v
    ax.set_yticks(y); ax.set_yticklabels([f"{p} (N = {br_int(NP[p])})" for p in PERFIS])
    ax.set_xlim(0, 100); ax.xaxis.set_major_formatter(lambda v_, _: f"{v_:.0f}%")
    ax.set_xlabel("% dos posts do perfil")
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(title=legenda, frameon=False, ncol=4 if len(cats) > 4 else len(cats), fontsize=8,
              loc="upper center", bbox_to_anchor=(0.5, -0.2))
    fig.text(0.99, 0.003, "Valores (%) escritos nos segmentos ≥ 5%; os demais estão na planilha · " + nome,
             ha="right", fontsize=7.5, color="dimgray")
    fig.tight_layout(rect=(0, 0.02, 1, 1)); fig.savefig(png, dpi=300, bbox_inches="tight"); plt.close(fig)


# (a) recursos: 7 maiores + "Outros recursos" só para o gráfico (na planilha, todos separados)
cats_a = [marca_rec(r) for r in REC_ORD[:7]]
outros = [marca_rec(r) for r in REC_ORD[7:]]
ta = t_rec.copy()
lin_o = {"Categoria": "Outros recursos"}
for p in PERFIS:
    lin_o[f"{p} (%)"] = ta[ta["Categoria"].isin(outros)][f"{p} (%)"].sum()
ta = pd.concat([ta, pd.DataFrame([lin_o])], ignore_index=True)
barras100(ta, cats_a + ["Outros recursos"], "\"Outros recursos\" = " + ", ".join(outros) + " (só no gráfico)", "Recurso Audiovisual", PNG_A)
barras100(t_imp, IMPS, "Impacto Visual: rótulo único por post", "Impacto Visual", PNG_B)

# (c) mapa de calor Recurso × Eixo por perfil (% dos posts do eixo no perfil)
fig, axs = plt.subplots(1, 3, figsize=(21, 8.5), gridspec_kw={"wspace": 0.75})
recs_c = REC_ORD[:8]
cmap = matplotlib.colormaps["cividis_r"].copy(); cmap.set_bad("white")
for ax, p in zip(axs, PERFIS):
    d = t_eixo[t_eixo["Perfil"] == p]
    H = np.array([[r[f"{marca_rec(rc)} (%)"] * 100 if r["N da categoria no perfil"] else np.nan for rc in recs_c] for _, r in d.iterrows()])
    im = ax.imshow(np.ma.masked_invalid(np.where(H == 0, np.nan, H)), cmap=cmap, vmin=0, vmax=100, aspect="auto")
    for i in range(H.shape[0]):
        for j in range(H.shape[1]):
            if not np.isnan(H[i, j]) and H[i, j] > 0:
                ax.text(j, i, f"{H[i, j]:.0f}" if H[i, j] >= 1 else "<1", ha="center", va="center", fontsize=7.5, color="white" if H[i, j] > 55 else "black")
    for i in range(H.shape[0] + 1):
        ax.axhline(i - 0.5, color="#e6e6e6", lw=0.5)
    for j in range(H.shape[1] + 1):
        ax.axvline(j - 0.5, color="#e6e6e6", lw=0.5)
    ax.set_xticks(range(len(recs_c))); ax.set_xticklabels([marca_rec(r) for r in recs_c], rotation=50, ha="right", fontsize=8)
    ax.set_yticks(range(len(d)))
    ax.set_yticklabels([f"{e} (N = {n})" for e, n in zip(d[EIXO], d["N da categoria no perfil"])], fontsize=8.5)
    ax.set_title(f"{p} (N = {br_int(NP[p])})", loc="left", fontsize=11)
    for s in ax.spines.values():
        s.set_visible(False)
cb = fig.colorbar(im, ax=axs, fraction=0.015, pad=0.01); cb.set_label("% dos posts do eixo no perfil")
fig.text(0.99, -0.11, "Células: % dos posts do eixo (no perfil) com o recurso; em branco = 0 · * eixo com menos de 10 posts no perfil (instável) · "
         "8 recursos mais frequentes (Printscreen, Fotografia com som, Vídeo sem som † e Desenho † só na planilha)", ha="right", fontsize=7.5, color="dimgray")
fig.savefig(PNG_C, dpi=300, bbox_inches="tight"); plt.close(fig)

# ------------------------------------------------------------------ conferência
pd.set_option("display.width", 250); pd.set_option("display.max_columns", 30)
for nome, t in [("REC", t_rec), ("TIPO", t_tipo), ("IMP", t_imp)]:
    print(nome); print(t[["Categoria"] + [f"{p} (N)" for p in PERFIS] + [f"{p} (%)" for p in PERFIS]].round(3).to_string())
print(t_comb[t_comb["Perfil"] == "Ego"].round(3).to_string())
print(t_reg[["Regime de Visibilidade"] + [f"{p} (% dos posts)" for p in PERFIS]].round(3).to_string())
print(t_res.round(3).to_string())
e = t_eixo[t_eixo["Perfil"] == "Ego"]
print(e[[EIXO, "N da categoria no perfil"] + [f"{marca_rec(r)} (%)" for r in REC_ORD[:6]]].round(2).to_string())
m = t_med[t_med["Perfil"] == "Ego"]
print(m[[MED, "N da categoria no perfil"] + [f"{marca_rec(r)} (%)" for r in REC_ORD[:6]]].round(2).to_string())
print("evidências:", len(t_ev))

# ------------------------------------------------------------------ tabelas em JSON (para o .docx)
def _js(t):
    return json.loads(t.to_json(orient="records", force_ascii=False))
json.dump({"rec": _js(t_rec), "tipo": _js(t_tipo), "imp": _js(t_imp), "comb": _js(t_comb), "eixo": _js(t_eixo),
           "med": _js(t_med), "reg": _js(t_reg), "res": _js(t_res), "eixos10": EIXOS10, "rec_ord": [marca_rec(r) for r in REC_ORD],
           "nrec": {p: int(NREC[p]) for p in PERFIS}},
          open("s19_tabelas.json", "w", encoding="utf8"), ensure_ascii=False)
