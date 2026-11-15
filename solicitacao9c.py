"""Solicitação 9C — Bairro × Eixo Temático.
Base: Solicitacao9_Referencia_local_agrupada.xlsx (agrupamentos de grafia aprovados) + planilha congelada (Eixo).
Tipologia dos locais aprovada pela autora em 05/11/2026 (Santa Inês, Sebastião Régis e Vila Nova = bairros de Imperatriz; Beira Rio = avenida).
Uso: python3 solicitacao9.py <planilha.xlsx> && python3 solicitacao9c.py <planilha.xlsx>
"""
import sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
SEM_PUB = "--sem-publicidade" in sys.argv  # exclui os posts do eixo Publicidade do cruzamento
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
AGRUP = "Solicitacao9_Referencia_local_agrupada.xlsx"
SAIDA = "Solicitacao9C_Bairro_x_Eixo_sem_Publicidade.xlsx" if SEM_PUB else "Solicitacao9C_Bairro_x_Eixo.xlsx"
EIXO = "Eixo Temático"
EXCL = {"Não identificado", "Imperatriz", "Maranhão"}

# Tipologia PROPOSTA (a confirmar pela autora) para os 30 locais mais citados (versão B, total)
TIPO = {
    "Centro": "Bairro", "Entroncamento": "Bairro", "Nova Imperatriz": "Bairro", "Bacuri": "Bairro",
    "Bom Sucesso": "Bairro", "Juçara": "Bairro", "Mercadinho": "Bairro", "Santa Rita": "Bairro",
    "Vila Cafeteira": "Bairro", "Vila Lobão": "Bairro", "Jardim Oriental": "Bairro",
    "Parque Santa Lúcia": "Bairro", "Parque Alvorada I": "Bairro", "Maranhão Novo": "Bairro",
    "Parque Alvorada II": "Bairro", "Parque São José": "Bairro", "Vilinha": "Bairro",
    "Jardim São Luís": "Bairro", "Vila João Castelo": "Bairro",
    "Tocantins": "Estado", "Pará": "Estado",
    "Açailândia": "Cidade", "João Lisboa": "Cidade", "São Luís": "Cidade", "Davinópolis": "Cidade",
    "Santa Inês": "Bairro",
    "BR-010": "Rodovia",
    "Beira Rio": "Avenida",
    "Sebastião Régis": "Bairro",
    "Vila Nova": "Bairro",
}

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
ag = pd.read_excel(AGRUP, sheet_name="Dados")
assert len(ag) == 4266 and ag["url "].is_unique
d = df[["url ", "Perfil", EIXO]].merge(ag[["url ", "Referência Local (agrupada)"]], on="url ", validate="1:1")
d["loc"] = d["Referência Local (agrupada)"].astype(str).str.split(";").apply(lambda l: sorted({x.strip() for x in l if x.strip()}))

# top 30 locais (versão B, total) — mesmos da Solicitação 9
cont = d["loc"].explode().loc[lambda s: ~s.isin(EXCL)].value_counts().rename_axis("l").reset_index(name="n")
cont = cont.sort_values(["n", "l"], ascending=[False, True])
cont["pos"] = cont["n"].rank(method="min", ascending=False).astype(int)
TOP30 = cont[cont["pos"] <= 30]["l"].tolist()
assert set(TOP30) == set(TIPO), set(TOP30) ^ set(TIPO)
BAIRROS = [l for l in TOP30 if TIPO[l] == "Bairro"]

freq = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
EIXOS = freq["e"].iloc[:10].tolist()
d["eixo_g"] = d[EIXO].where(d[EIXO].isin(EIXOS), "Outros eixos")
COLS = EIXOS + ["Outros eixos"]
N_ANTES = len(d)
if SEM_PUB:
    d = d[d[EIXO] != "Publicidade"].copy()
    COLS = [c for c in COLS if c != "Publicidade"]
    assert len(d) == N_ANTES - 1742
long = d[["url ", "Perfil", "eixo_g", "loc"]].explode("loc").reset_index(drop=True)


def tabela(locais, sub=None):
    L = long[long["loc"].isin(locais)]
    if sub:
        L = L[L["Perfil"] == sub]
    ct = pd.crosstab(L["loc"], L["eixo_g"]).reindex(index=locais, columns=COLS, fill_value=0)
    out = pd.DataFrame(index=ct.index)
    out["Tipo"] = [TIPO[l] for l in ct.index]
    out["N de posts do local"] = ct.sum(axis=1)
    for e in COLS:
        out[f"{e} (N)"] = ct[e]
    for e in COLS:
        out[f"{e} (%)"] = np.where(out["N de posts do local"] > 0, ct[e] / out["N de posts do local"].replace(0, np.nan), np.nan)
    out = out.sort_values("N de posts do local", ascending=False, kind="stable")
    out.index = [l + (" *" if n < 10 else "") for l, n in zip(out.index, out["N de posts do local"])]
    # linha de referência: posts que citam pelo menos um dos locais (cada post uma vez)
    base = d if sub is None else d[d["Perfil"] == sub]
    m = base["loc"].apply(lambda l: any(x in locais for x in l))
    ref = base.loc[m, "eixo_g"].value_counts().reindex(COLS, fill_value=0)
    nref = int(m.sum())
    out.loc["Posts que citam pelo menos um destes locais"] = ["—", nref] + list(ref.values) + list(ref.values / nref)
    out = out.reset_index().rename(columns={"index": "Local"})
    for c in [x for x in out.columns if "(N)" in x or x == "N de posts do local"]:
        out[c] = out[c].astype(int)
    corpo = out.iloc[:-1]
    assert (corpo[[f"{e} (N)" for e in COLS]].sum(axis=1) == corpo["N de posts do local"]).all()
    return out


abas = {"Bairros_Total": tabela(BAIRROS)}
for p in ["Ego", "Roma", "Imperatriz Online"]:
    abas["Bairros_" + p.replace(" ", "_")] = tabela(BAIRROS, p)
abas["Top30_locais_todos_tipos"] = tabela(TOP30)
tip = pd.DataFrame({"Local": TOP30, "N de posts (total)": [int(cont.set_index("l").loc[l, "n"]) for l in TOP30],
                    "Tipo (aprovado)": [TIPO[l] for l in TOP30],
                    "Decisão da autora": ["Aprovado (05/11/2026)" for _ in TOP30]})

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia) + Solicitacao9_Referencia_local_agrupada.xlsx (agrupamentos de grafia aprovados)."],
    ["VERSÃO", ("SEM PUBLICIDADE: os 1.742 posts do eixo Publicidade foram excluídos antes do cruzamento, a pedido da autora (05/11/2026), para que o peso da publicidade não domine a distribuição temática dos bairros. Base: 2.524 posts. Os 19 bairros e os 30 locais são os mesmos da versão com Publicidade (seleção feita sobre o total geral), para permitir comparação."
                if SEM_PUB else "COM PUBLICIDADE: todos os 4.266 posts.")],
    ["N analisado", ("2.524 posts (4.266 − 1.742 de Publicidade)" if SEM_PUB else "4.266 posts") + ". Em cada aba entram só os posts que citam os locais da linha."],
    ["Locais", f"Os 30 locais mais citados no total (versão B da Solicitação 9: sem \"Não identificado\", \"Imperatriz\" e \"Maranhão\"). Abas Bairros_*: os {len(BAIRROS)} classificados como \"Bairro\" na tipologia aprovada pela autora."],
    ["TIPOLOGIA", "Classificação em Bairro/Cidade/Estado/Rodovia/Avenida proposta pela assistente e APROVADA pela autora em 05/11/2026, com estas decisões: Santa Inês = bairro de Imperatriz (antes proposto como cidade); Sebastião Régis e Vila Nova = bairros de Imperatriz; Beira Rio = avenida (fora das abas Bairros_*). Versão anterior: 19 bairros; esta: 22."],
    ["Eixos", "Os 10 eixos mais frequentes do total geral (mesmos da Solicitação 4) + \"Outros eixos\" (24 eixos)."],
    ["Fórmula do %", "Percentual de linha: % = nº de posts do local no eixo ÷ N de posts do local (dentro da aba) × 100. Cada linha soma 100% (Eixo tem um rótulo por post)."],
    ["Sobreposição", "Um post pode citar mais de um local; nesse caso conta na linha de cada local. Por isso as linhas não se somam. A última linha conta cada post uma vez (posts que citam pelo menos um dos locais da aba)."],
    ["*", "Local com menos de 10 posts dentro da aba (percentuais instáveis)."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a Eixo Temático ou Referência Local."],
    ["Script", "solicitacao9c.py" + (" --sem-publicidade" if SEM_PUB else "")],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    for nome, t in abas.items():
        t.to_excel(w, sheet_name=nome, index=False)
    tip.to_excel(w, sheet_name="Tipologia_para_aprovacao", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

amarelo = PatternFill("solid", fgColor="FFF2CC")
wb = load_workbook(SAIDA)
for nome in abas:
    ws = wb[nome]
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[1].height = 45
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    for col in ws.columns:
        h = str(col[0].value)
        for c in col[1:]:
            if "(%)" in h:
                c.number_format = "0.0%"
            elif "(N)" in h or h.startswith("N de"):
                c.number_format = "#,##0"
        ws.column_dimensions[col[0].column_letter].width = 11
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    for row in ws.iter_rows(min_row=2):
        if str(row[1].value).startswith("A confirmar"):
            for c in row:
                c.fill = amarelo
    ws.freeze_panes = "C2"
ws = wb["Tipologia_para_aprovacao"]
for c in ws[1]:
    c.font = Font(bold=True)
for L, wd in zip("ABCD", (24, 18, 50, 30)):
    ws.column_dimensions[L].width = wd
for row in ws.iter_rows(min_row=2):
    row[1].number_format = "#,##0"
    if str(row[2].value).startswith("A confirmar"):
        for c in row:
            c.fill = amarelo
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
    if row[0].value == "TIPOLOGIA = PROPOSTA":
        for c in row:
            c.fill = amarelo
            c.font = Font(bold=True)
wb.save(SAIDA)

pd.set_option("display.width", 320)
t = abas["Bairros_Total"]
show = t[["Local", "N de posts do local"] + [f"{e} (%)" for e in COLS]].copy()
for e in COLS:
    show[f"{e} (%)"] = (show[f"{e} (%)"] * 100).round(1)
show.columns = ["Local", "N"] + [e[:8] for e in COLS]
print(show.to_string(index=False))
for nome in ["Bairros_Ego", "Bairros_Roma", "Bairros_Imperatriz_Online"]:
    t = abas[nome]
    print(nome, t.iloc[-1]["N de posts do local"], (t.iloc[:-1]["N de posts do local"] < 10).sum())

# ---- Mapa de calor: Bairros_Total (bairros nas linhas, eixos nas colunas; % de linha)
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

PNG = SAIDA.replace(".xlsx", ".png")
t = abas["Bairros_Total"].iloc[:-1]
M = t[[f"{e} (%)" for e in COLS]].values * 100
fig, ax = plt.subplots(figsize=(13, 9))
im = ax.imshow(M, cmap="cividis_r", vmin=0, vmax=100, aspect="auto")
for i in range(M.shape[0]):
    for j in range(M.shape[1]):
        v = M[i, j]
        ax.text(j, i, f"{v:.1f}".replace(".", ","), ha="center", va="center", fontsize=7.5,
                color="white" if v > 50 else "black")
ax.set_xticks(range(len(COLS)))
ax.set_xticklabels(COLS, rotation=35, ha="right", fontsize=9)
ax.set_yticks(range(len(t)))
ax.set_yticklabels([f"{l} (N = {n:,})".replace(",", ".") for l, n in zip(t["Local"], t["N de posts do local"])], fontsize=9)
cb = fig.colorbar(im, ax=ax, fraction=0.03, pad=0.01)
cb.set_label("% dos posts que citam o bairro")
cb.ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
ax.set_xlabel("Eixo Temático (valores em %; cada linha soma 100%)")
fig.text(0.01, 0.005, ("Sem os posts do eixo Publicidade. " if SEM_PUB else "") + "Bairros segundo tipologia aprovada pela autora (05/11/2026). Um post que cita mais de um bairro conta em cada linha.",
         fontsize=7.5, color="dimgray")
fig.tight_layout()
fig.savefig(PNG, dpi=300, bbox_inches="tight")
print("png ok")
