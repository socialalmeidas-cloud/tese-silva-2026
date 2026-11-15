"""Solicitação 17 — Os "rostos" da cidade × Referência Local.
Pergunta: em que escala territorial cada perfil fala da cidade e quais eixos temáticos acompanham os bairros?
Base: planilha congelada (Perfil, Eixo Temático) + Solicitacao9_Referencia_local_agrupada.xlsx (grafias agrupadas
aprovadas) + tipologia dos 30 locais mais citados aprovada pela autora em 05/11/2026 (mesma da Solicitação 9C).
Uso: python3 solicitacao17_rostos.py copia.xlsx
Gera Solicitacao17_Rostos_x_Referencia_local.xlsx e .png.
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
AGRUP = "Solicitacao9_Referencia_local_agrupada.xlsx"
SAIDA = "Solicitacao17_Rostos_x_Referencia_local.xlsx"
PNG = "Solicitacao17_Rostos_x_Referencia_local.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
EIXO = "Eixo Temático"

# Tipologia aprovada (05/11/2026) — idêntica à solicitacao9c.py
TIPO = {
    "Centro": "Bairro", "Entroncamento": "Bairro", "Nova Imperatriz": "Bairro", "Bacuri": "Bairro",
    "Bom Sucesso": "Bairro", "Juçara": "Bairro", "Mercadinho": "Bairro", "Santa Rita": "Bairro",
    "Vila Cafeteira": "Bairro", "Vila Lobão": "Bairro", "Jardim Oriental": "Bairro",
    "Parque Santa Lúcia": "Bairro", "Parque Alvorada I": "Bairro", "Maranhão Novo": "Bairro",
    "Parque Alvorada II": "Bairro", "Parque São José": "Bairro", "Vilinha": "Bairro",
    "Jardim São Luís": "Bairro", "Vila João Castelo": "Bairro", "Santa Inês": "Bairro",
    "Sebastião Régis": "Bairro", "Vila Nova": "Bairro",
    "Tocantins": "Estado", "Pará": "Estado",
    "Açailândia": "Cidade", "João Lisboa": "Cidade", "São Luís": "Cidade", "Davinópolis": "Cidade",
    "BR-010": "Rodovia", "Beira Rio": "Avenida",
}
BAIRROS = [l for l, t in TIPO.items() if t == "Bairro"]
assert len(BAIRROS) == 22
GERAIS = {"Imperatriz", "Maranhão"}

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
assert df["Perfil"].value_counts().to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
ag = pd.read_excel(AGRUP, sheet_name="Dados")
d = df[["url ", "Perfil", EIXO]].merge(ag[["url ", "Referência Local (agrupada)"]], on="url ", validate="1:1")
d["loc"] = d["Referência Local (agrupada)"].astype(str).str.split(";").apply(lambda l: {x.strip() for x in l if x.strip()})


# ---- Escala territorial: um rótulo por post, pela ordem abaixo (o primeiro que se aplica)
def escala(locs):
    if locs & set(BAIRROS):
        return "1 Cita bairro (22 bairros)"
    if any(l in TIPO for l in locs):
        return "2 Cita cidade, estado, rodovia ou avenida (sem bairro)"
    if locs - GERAIS - {"Não identificado"}:
        return "3 Cita outro local (fora dos 30 mais citados; tipo não classificado)"
    if locs & GERAIS:
        return "4 Só \"Imperatriz\" e/ou \"Maranhão\""
    return "5 Não identificado"


d["escala"] = d["loc"].apply(escala)
ESC = sorted(d["escala"].unique())
assert len(ESC) == 5


def tab_pct(col, cats, base=None):
    base = d if base is None else base
    rows = []
    for p in PERFIS + ["Total"]:
        b = base if p == "Total" else base[base["Perfil"] == p]
        vc = b[col].value_counts().reindex(cats, fill_value=0)
        rows.append([p, len(b)] + [x for c in cats for x in (int(vc[c]), vc[c] / len(b))])
    cols = ["Perfil", "N de posts"] + [x for c in cats for x in (f"{c[2:]} (N)", f"{c[2:]} (%)")]
    t = pd.DataFrame(rows, columns=cols)
    assert (t[[c for c in cols if c.endswith("(N)")]].sum(axis=1) == t["N de posts"]).all()
    return t


t_esc = tab_pct("escala", ESC)
SP = d[d[EIXO] != "Publicidade"]
assert len(SP) == 4266 - 1742
t_esc_sp = tab_pct("escala", ESC, SP)

# ---- Eixos dos posts com bairro × sem bairro, por perfil (10 eixos mais frequentes do total + Outros)
freq = d[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
EIXOS = freq["e"].iloc[:10].tolist()
d["eixo_g"] = d[EIXO].where(d[EIXO].isin(EIXOS), "Outros eixos")
COLS = EIXOS + ["Outros eixos"]
d["tem_bairro"] = d["escala"].str.startswith("1").map({True: "Com bairro", False: "Sem bairro"})
rows = []
for p in PERFIS + ["Total"]:
    for g in ["Com bairro", "Sem bairro"]:
        b = d if p == "Total" else d[d["Perfil"] == p]
        b = b[b["tem_bairro"] == g]
        vc = b["eixo_g"].value_counts().reindex(COLS, fill_value=0)
        rows.append([p, g, len(b)] + list(vc / len(b) if len(b) else vc * np.nan))
t_eixo = pd.DataFrame(rows, columns=["Perfil", "Posts", "N"] + [f"{e} (%)" for e in COLS])

# ---- Bairros mais citados por perfil, com o eixo mais frequente (com e sem Publicidade)
long = d[["Perfil", EIXO, "loc"]].explode("loc")
long = long[long["loc"].isin(BAIRROS)]
rows = []
for p in PERFIS:
    L = long[long["Perfil"] == p]
    cont = L["loc"].value_counts()
    cont = cont.rename_axis("b").reset_index(name="n").sort_values(["n", "b"], ascending=[False, True]).head(5)
    for b, n in zip(cont["b"], cont["n"]):
        e = L[L["loc"] == b][EIXO].value_counts()
        e2 = L[(L["loc"] == b) & (L[EIXO] != "Publicidade")][EIXO].value_counts()
        rows.append([p, b + (" *" if n < 10 else ""), int(n), e.index[0], e.iloc[0] / n,
                     int(e2.sum()), (e2.index[0] if len(e2) else "—"), (e2.iloc[0] / e2.sum() if len(e2) else np.nan)])
rows_sp = []
for p in PERFIS:
    L = long[(long["Perfil"] == p) & (long[EIXO] != "Publicidade")]
    cont = L["loc"].value_counts().rename_axis("b").reset_index(name="n").sort_values(["n", "b"], ascending=[False, True]).head(5)
    for b, n in zip(cont["b"], cont["n"]):
        e = L[L["loc"] == b][EIXO].value_counts()
        rows_sp.append([p, b + (" *" if n < 10 else ""), int(n), e.index[0], e.iloc[0] / n,
                        (e.index[1] if len(e) > 1 else "—"), (e.iloc[1] / n if len(e) > 1 else np.nan)])
t_top_sp = pd.DataFrame(rows_sp, columns=["Perfil", "Bairro", "N de posts (sem Publicidade)", "1º eixo", "% do 1º eixo", "2º eixo", "% do 2º eixo"])
t_top = pd.DataFrame(rows, columns=["Perfil", "Bairro", "N de posts que citam o bairro", "Eixo mais frequente",
                                    "% do eixo", "N sem Publicidade", "Eixo mais frequente sem Publicidade",
                                    "% do eixo (sem Publicidade)"])

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia) + " + AGRUP + " (grafias agrupadas aprovadas)."],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721). Exclusões: nenhuma."],
    ["Tipologia", "Tipologia dos 30 locais mais citados aprovada pela autora em 05/11/2026 (22 bairros; Beira Rio = avenida). Locais fora dos 30 mais citados não têm tipo classificado."],
    ["Escala territorial", "Um rótulo por post, atribuído na ordem: (1) cita pelo menos um dos 22 bairros; (2) não cita bairro, mas cita cidade, estado, rodovia ou avenida da tipologia; (3) cita outro local, fora dos 30 mais citados (pode ser bairro, cidade ou outro tipo; não classificado); (4) cita só \"Imperatriz\" e/ou \"Maranhão\"; (5) \"Não identificado\". A ordem é PROPOSTA DA ASSISTENTE: prioriza a escala mais próxima (bairro). Cada perfil soma 100%."],
    ["Eixos", "Os 10 eixos mais frequentes do total geral (mesmos das Solicitações 4 e 9C) + \"Outros eixos\". Percentual de linha: % = posts do grupo no eixo ÷ N do grupo × 100."],
    ["Bairros_por_perfil", "5 bairros mais citados em cada perfil. \"Eixo mais frequente\" com e sem Publicidade. Um post que cita dois bairros conta em cada um."],
    ["Sem Publicidade", "Abas Escala_sem_Publicidade e Bairros_por_perfil_sem_Pub: excluídos os 1.742 posts do eixo Publicidade (base 2.524), porque nos anúncios o bairro costuma ser o endereço do anunciante."],
    ["*", "Menos de 10 posts (percentual instável)."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a Eixo Temático ou Referência Local."],
    ["Interpretação", "Os \"rostos\" da cidade (Solicitação 16) são leitura interpretativa; esta solicitação só oferece o cruzamento quantitativo. A leitura conjunta está no memorando e depende da confirmação destas tabelas pela autora."],
    ["Script", "solicitacao17_rostos.py (renomeado em 05/11/2026, quando solicitacao17.py passou a ser a Temática Secundária)"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t_esc.to_excel(w, sheet_name="Escala_territorial", index=False)
    t_eixo.to_excel(w, sheet_name="Eixos_com_sem_bairro", index=False)
    t_esc_sp.to_excel(w, sheet_name="Escala_sem_Publicidade", index=False)
    t_top.to_excel(w, sheet_name="Bairros_por_perfil", index=False)
    t_top_sp.to_excel(w, sheet_name="Bairros_por_perfil_sem_Pub", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)
wb = load_workbook(SAIDA)
for ws in wb.worksheets:
    for c in ws[1]:
        c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            h = str(ws.cell(1, c.column).value)
            if h.endswith("(%)") or h.startswith("% do"):
                c.number_format = "0.0%"
            elif isinstance(c.value, (int, np.integer)):
                c.number_format = "#,##0"
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 16
ws = wb["Notas"]; ws.column_dimensions["A"].width = 22; ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico: escala territorial por perfil (barras 100% empilhadas)
cores = ["#0072B2", "#56B4E9", "#009E73", "#E69F00", "#BBBBBB"]
fig, ax = plt.subplots(figsize=(11, 3.6))
lab = [f"{p}\n(N = {n:,})".replace(",", ".") for p, n in zip(t_esc["Perfil"], t_esc["N de posts"])]
esq = np.zeros(len(t_esc))
for c, cor in zip(ESC, cores):
    v = t_esc[f"{c[2:]} (%)"].values * 100
    ax.barh(lab, v, left=esq, color=cor, label=c[2:], edgecolor="white")
    for i, (x, w_) in enumerate(zip(esq, v)):
        if w_ >= 4:
            ax.text(x + w_ / 2, i, f"{w_:.1f}".replace(".", ","), ha="center", va="center", fontsize=8,
                    color="white" if cor in ("#0072B2", "#009E73") else "black")
    esq += v
ax.invert_yaxis(); ax.set_xlim(0, 100); ax.set_xlabel("% dos posts do perfil")
ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.28), ncol=2, fontsize=8, frameon=False)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
fig.text(0.01, 0.005, "Um rótulo por post (escala mais próxima citada). Valores abaixo de 4% sem rótulo; ver tabela.", fontsize=7, color="#555")
fig.tight_layout(); fig.savefig(PNG, dpi=300, bbox_inches="tight")

pd.set_option("display.width", 250)
print(t_esc.to_string()); print(t_eixo.round(3).to_string()); print(t_top.to_string()); print(t_esc_sp.iloc[:, [0,1,3,5,7,9,11]].to_string()); print(t_top_sp.to_string())
