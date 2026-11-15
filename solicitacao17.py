"""Solicitação 17 — Temática Secundária (rótulos livres, multirrótulo).

Uso: python3 -I solicitacao17.py <planilha.xlsx>

O que o script faz (sem alterar a planilha original, que é lida de uma cópia):
  1. Lê a coluna "Temática Secundária " (o nome real tem um espaço no fim).
  2. Separa cada célula por ";" e tira espaços das pontas de cada rótulo.
     v1: nenhuma variante agrupada. v2: só as variantes aprovadas pela autora (variantes_aprovadas.py).
  3. Conta POSTS, não células: cada rótulo conta no máximo uma vez por post.
  4. Gera as abas pedidas, a aba de variantes (com a decisão da autora) e o gráfico.
  v2 (07/11/2026): aplica as variantes aprovadas (módulo variantes_aprovadas.py).
"""
import sys
import re
import unicodedata
from difflib import SequenceMatcher
from itertools import combinations

import numpy as np
import pandas as pd
import os as _os; sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))  # pasta do script (python -I não a inclui)
from variantes_aprovadas import MAPA, DATA_DECISAO, rotulos_do_post, tabela_correspondencia
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao17_Tematica_secundaria.xlsx"
PNG = "Solicitacao17_Tematica_secundaria.png"
COL = "Temática Secundária "          # nome real, com espaço no fim
EIXO = "Eixo Temático"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
ABA = {"Ego": "Ego", "Roma": "Roma", "Imperatriz Online": "Imperatriz_Online", "Total": "Total"}
SEM = "Não há"                         # valor de "sem temática secundária" pedido pela autora
LIMIAR_TOP = 5                         # * para rótulos com menos de 5 posts
br_int = lambda v: f"{int(v):,}".replace(",", ".")
br_pct = lambda v: f"{v:.1f}%".replace(".", ",")

# ---------------------------------------------------------------- leitura e conferências
df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
assert df["Perfil"].value_counts().to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
assert COL in df.columns
assert df[COL].notna().all()
df["Linha"] = df.index + 2                       # linha na planilha (cabeçalho = linha 1)
df[EIXO] = df[EIXO].astype(str).str.strip()

# Rótulos por post: separa por ";" e tira espaços das pontas; descarta vazios
df["rot"] = df[COL].astype(str).str.split(";").apply(lambda l: [x.strip() for x in l if x.strip()])
assert not df["rot"].apply(lambda l: len(l) != len(set(l))).any()   # nenhum rótulo repetido no mesmo post
df["rot_bruto"] = df["rot"].apply(lambda l: [x for x in l if x != SEM])          # grafia original (para a aba de variantes)
df["rot_validos"] = df[COL].apply(rotulos_do_post)                                # variantes aprovadas aplicadas (v2, 07/11/2026)
df["n_rot"] = df["rot_validos"].apply(len)

GRUPOS = [(p, df[df["Perfil"] == p]) for p in PERFIS] + [("Total", df)]

# ---------------------------------------------------------------- aba Rotulos_por_post
linhas = []
for nome, d in GRUPOS:
    n = len(d)
    sem = int((d["n_rot"] == 0).sum())
    com = d[d["n_rot"] > 0]
    cont = {"1 rótulo": (d["n_rot"] == 1).sum(), "2 rótulos": (d["n_rot"] == 2).sum(),
            "3 rótulos": (d["n_rot"] == 3).sum(), "Mais de 3 rótulos": (d["n_rot"] > 3).sum()}
    lin = {"Perfil": nome, "N de posts": n, f'Sem temática secundária ("{SEM}")': sem}
    for k, v in cont.items():
        lin[f"{k} (N)"] = int(v)
        lin[f"{k} (%)"] = v / n
    lin["Média de rótulos por post (todos os posts)"] = d["n_rot"].mean()
    lin["Média de rótulos por post (só posts com temática)"] = com["n_rot"].mean()
    lin["Total de rótulos atribuídos"] = int(d["n_rot"].sum())
    assert sem + sum(cont.values()) == n
    linhas.append(lin)
t_rot = pd.DataFrame(linhas)

mais3 = df[df["n_rot"] > 3].sort_values("Linha")
t_mais3 = pd.DataFrame({"Linha": mais3["Linha"], "Perfil": mais3["Perfil"], "Eixo": mais3[EIXO],
                        "Nº de rótulos": mais3["n_rot"], "Rótulos": mais3["rot_validos"].apply("; ".join)})

# ---------------------------------------------------------------- posts por rótulo
def contar(d):
    """N de posts com cada rótulo (cada rótulo uma vez por post), sem "Não há".
    Empates: ordem alfabética, só para fixar a posição."""
    s = d["rot_validos"].explode().dropna()
    c = s.value_counts().rename_axis("Temática").reset_index(name="N de posts")
    return c.sort_values(["N de posts", "Temática"], ascending=[False, True]).reset_index(drop=True)


def top(d, k, denom):
    c = contar(d)
    corte = c["N de posts"].iloc[k - 1] if len(c) >= k else None
    t = c.head(k).copy()
    t.insert(0, "Posição", range(1, len(t) + 1))
    t["% dos posts"] = t["N de posts"] / denom
    t["Temática"] = [f"{x} *" if n < LIMIAR_TOP else x for x, n in zip(t["Temática"], t["N de posts"])]
    # empate na fronteira do top: quantos rótulos com o mesmo N ficaram de fora
    fora = int(((c["N de posts"] == corte).sum() - (t["N de posts"] == corte).sum())) if corte is not None else 0
    return t, fora


tops, empates = {}, {}
for nome, d in GRUPOS:
    t, fora = top(d, 30, len(d))
    t = t.rename(columns={"% dos posts": "% dos posts do perfil" if nome != "Total" else "% dos posts (total)"})
    tops[nome], empates[nome] = t, fora

# ---------------------------------------------------------------- aba Frequencia_dos_rotulos
freq = []
for nome, d in GRUPOS:
    c = contar(d)
    distintos, unicos = len(c), int((c["N de posts"] == 1).sum())
    freq.append({"Perfil": nome, "Rótulos distintos (N)": distintos,
                 "Rótulos que aparecem em um só post (N)": unicos,
                 "Rótulos que aparecem em um só post (% dos distintos)": unicos / distintos,
                 "Rótulos com menos de 5 posts (N)": int((c["N de posts"] < 5).sum()),
                 "Rótulos com menos de 5 posts (% dos distintos)": (c["N de posts"] < 5).mean()})
t_freq = pd.DataFrame(freq)

# ---------------------------------------------------------------- aba Por_eixo
fe = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
TOP10 = fe["e"].iloc[:10].tolist()                 # mesmos 10 eixos das Solicitações 4 e 6
blocos = []
for e in TOP10:
    d = df[df[EIXO] == e]
    t, fora = top(d, 10, len(d))
    t.insert(0, "Eixo", f"{e} (N = {br_int(len(d))})")
    t = t.rename(columns={"% dos posts": "% dos posts do eixo"})
    t["Empates fora do top 10"] = fora
    blocos.append(t)
t_eixo = pd.concat(blocos, ignore_index=True)

# ---------------------------------------------------------------- aba Variantes_para_aprovacao
def sem_acento(s):
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")


def chave(s):                                      # minúsculas, sem acento, espaços normalizados
    return " ".join(sem_acento(s).lower().replace(".", " ").split())


def singular(s):                                   # plural simples do português, só para comparar
    pal = []
    for w in s.split():
        for suf, rep in (("oes", "ao"), ("aes", "ao"), ("ais", "al"), ("eis", "el"), ("res", "r"), ("zes", "z"), ("ns", "m"), ("s", "")):
            if len(w) > 3 and w.endswith(suf):
                w = w[: -len(suf)] + rep
                break
        pal.append(w)
    return " ".join(pal)


# A triagem de variantes usa a GRAFIA ORIGINAL (antes da unificação)
ct_total = df["rot_bruto"].explode().dropna().value_counts()
perfis_do = df.explode("rot_bruto").groupby("rot_bruto")["Perfil"].agg(lambda s: sorted(set(s), key=PERFIS.index))


def decisao(a, b=None):
    if b is None:
        return f"Unificado em \"{MAPA[a]}\" ({DATA_DECISAO})" if a in MAPA else "Mantido"
    if {a, b} in ({"Feira", "Feirão"}, {"Cantor", "Cantora"}, {"Campeã", "Campeão"}):
        return f"Não aprovado pela autora ({DATA_DECISAO}): mantidos separados"
    fa, fb = MAPA.get(a, a), MAPA.get(b, b)
    return f"Unificados em \"{fa}\" ({DATA_DECISAO})" if fa == fb else "Não unificar"
rotulos = sorted(ct_total.index)
pares = []
for a, b in combinations(rotulos, 2):
    ka, kb = chave(a), chave(b)
    tipo = None
    if ka == kb:
        tipo = "Acento, maiúscula ou pontuação"
    elif singular(ka) == singular(kb):
        tipo = "Singular/plural"
    else:
        curto, longo = sorted([ka, kb], key=len)
        # abreviação: rótulo curto (≥ 3 letras) que é início de uma única palavra do outro rótulo
        if len(curto) >= 3 and " " not in curto and " " not in longo and longo.startswith(curto) and len(longo) - len(curto) >= 2 \
                and (a.endswith(".") or b.endswith(".") or len(curto) <= 5):
            tipo = "Possível abreviação"
        elif min(len(ka), len(kb)) >= 5 and SequenceMatcher(None, ka, kb).ratio() >= 0.88:
            tipo = "Grafia muito parecida (possível erro de digitação)"
    if tipo:
        pa, pb = perfis_do[a], perfis_do[b]
        pares.append({"Variante A": a, "N A": int(ct_total[a]), "Variante B": b, "N B": int(ct_total[b]),
                      "Tipo de semelhança": tipo,
                      "Perfil onde aparecem": f"A: {', '.join(pa)} | B: {', '.join(pb)}",
                      "Decisão da autora": decisao(a, b)})
t_var = pd.DataFrame(pares)

# Avaliação prévia da assistente (NÃO aplicada; só ajuda a triagem). Pares fora do dicionário:
# "Provavelmente palavras diferentes" (a regra automática pegou só a semelhança de letras).
PROVAVEL = "Provável variante da mesma temática"
ERRO = "Provável erro de digitação"
DECIDIR = "A decidir (palavras relacionadas, mas não idênticas)"
AVALIACAO = {
    ("Alimento", "Alimentos"): PROVAVEL, ("Cantor", "Cantores"): PROVAVEL, ("Dica", "Dicas"): PROVAVEL,
    ("Equipamento", "Equipamentos"): PROVAVEL, ("Funcionário", "Funcionários"): PROVAVEL,
    ("Materiais", "Material"): PROVAVEL, ("Matrícula", "Matrículas"): PROVAVEL, ("Motorista", "Motoristas"): PROVAVEL,
    ("Perfume", "Perfumes"): PROVAVEL, ("Visita", "Visitas"): PROVAVEL, ("LGBT", "LGBTQIAPN+"): PROVAVEL,
    ("Câmara de Vereadores", "Câmara dos vereadores"): PROVAVEL,
    ("Corpus Christ", "Corpus Christi"): ERRO, ("Gentileza", "Gentiliza"): ERRO,
    ("Rastramento Veicular", "Rastreamento Veicular"): ERRO, ("Transplanete Capilar", "Transplante Capilar"): ERRO,
    ("Arraia", "Arraial"): ERRO + " (a conferir: \"Arraia\" também é o nome de um peixe)",
    ("Cantor", "Cantora"): DECIDIR + ": gênero", ("Campeã", "Campeão"): DECIDIR + ": gênero",
    ("Feira", "Feirão"): DECIDIR, ("Policial", "Polícia"): DECIDIR, ("Humor", "Humoristas"): DECIDIR,
    ("Pizza", "Pizzaria"): DECIDIR, ("Raio", "Raio-X"): DECIDIR + " (a conferir: \"Raio\" pode ser descarga elétrica)",
}
t_var["Avaliação prévia da assistente (não aplicada)"] = [
    AVALIACAO.get((a_, b_), "Provavelmente palavras diferentes") for a_, b_ in zip(t_var["Variante A"], t_var["Variante B"])]

# Rótulos isolados com grafia irregular (maiúscula no meio da palavra ou inicial minúscula): só informados
irreg = [x for x in ct_total.index if re.search(r"[a-zà-ú][A-ZÀ-Ú]", x) or x[:1].islower()]
t_irreg = pd.DataFrame({"Rótulo": irreg, "N de posts": [int(ct_total[x]) for x in irreg],
                        "Perfil onde aparece": [", ".join(perfis_do[x]) for x in irreg],
                        "Observação": ["Maiúscula no meio da palavra" if re.search(r"[a-zà-ú][A-ZÀ-Ú]", x) else "Inicial minúscula" for x in irreg],
                        "Decisão da autora": [decisao(x) for x in irreg]})
ordem_tipo = ["Acento, maiúscula ou pontuação", "Singular/plural", "Possível abreviação", "Grafia muito parecida (possível erro de digitação)"]
t_var["_o"] = t_var["Tipo de semelhança"].map(ordem_tipo.index)
t_var = t_var.sort_values(["_o", "Variante A", "Variante B"]).drop(columns="_o").rename(columns={"N A": "N", "N B": "N "})

# ---------------------------------------------------------------- Notas
nao_ident = df[df["rot_validos"].apply(lambda l: "Não identificado" in l)]
notas = pd.DataFrame([
    ["Arquivo", f"{NOME_ORIGINAL} (versão congelada de 04/11/2026; trabalhado em cópia; a planilha original não foi alterada)"],
    ["Coluna", 'Nome real: "Temática Secundária " (com espaço no fim). Preenchida em todos os 4.266 posts.'],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721). Nenhum post excluído."],
    ["Separação dos rótulos", 'Célula separada por ";" e espaços das pontas retirados. Nenhum rótulo repetido dentro do mesmo post na grafia original. Variantes aprovadas aplicadas (ver linha \"Variantes (v2)\").'],
    ["Contagem", "Posts, não células: cada rótulo conta no máximo uma vez por post."],
    ['"Não há"', f'O valor "{SEM}" NÃO ocorre na coluna (0 posts): todos os posts têm ao menos um rótulo. Se ocorresse, seria excluído dos tops.'],
    ['"Não identificado"', f'Ocorre em {len(nao_ident)} posts (linhas {", ".join(map(str, nao_ident["Linha"]))}). Foi mantido como rótulo, porque a instrução manda excluir só "Não há". Decisão a confirmar pela autora.'],
    ["% nos tops", "% = N de posts com o rótulo ÷ N de posts do perfil (ou do total, ou do eixo) × 100. Como o post pode ter vários rótulos, as % não somam 100%."],
    ["Asterisco (*)", "Rótulos com menos de 5 posts no perfil, no total ou no eixo da aba."],
    ["Empates", "Rótulos com o mesmo N ficam em ordem alfabética só para fixar a posição. "
     + "Empates na posição 30 deixados de fora: " + "; ".join(f"{k} {v}" for k, v in empates.items())
     + ". Na aba Por_eixo, a coluna \"Empates fora do top 10\" informa quantos rótulos com o mesmo N da 10ª posição ficaram de fora."],
    ["Por_eixo", "10 eixos mais frequentes do total (os mesmos das Solicitações 4 e 6). % = N de posts do eixo com o rótulo ÷ N de posts do eixo."],
    ["Variantes_para_aprovacao", "Pares detectados automaticamente (mesma grafia sem acento/maiúscula/ponto; singular/plural; rótulo curto que é início do outro; semelhança de grafia ≥ 0,88). "
     "É uma triagem: traz falsos positivos (ex.: palavras diferentes com início comum) e não detecta sinônimos. A coluna \"Avaliação prévia da assistente\" é só uma sugestão de triagem e não foi aplicada. "
     "Nenhum par por diferença só de acento ou maiúscula foi encontrado. Abaixo dos pares, a lista de rótulos isolados com grafia irregular (ex.: \"MóVéis\"). A coluna \"Decisão da autora\" registra a decisão de 07/11/2026."],
    ["Mais de 3 rótulos", f"O livro de códigos prevê no máximo 3 rótulos. {len(t_mais3)} posts têm mais de 3 (lista na aba Rotulos_por_post)."],
    ["Gráfico", "15 temáticas mais frequentes de cada perfil, um painel por perfil, % dos posts do perfil, rótulos como estão na planilha."],
    ["Variantes (v2)", f"Versão 2 ({DATA_DECISAO}): contagens com as variantes aprovadas pela autora aplicadas por tabela de correspondência (aba Variantes_aplicadas, módulo variantes_aprovadas.py). Não aprovados (mantidos separados): Cantor/Cantora, Campeã/Campeão, Feira/Feirão. Raio/Raio-X unificados por decisão da autora; pares do bloco D e Desaparecido/Desaparecimento, Pavimentação/Pavimentação pública não unificados. Se um post tinha duas grafias do mesmo grupo (ex.: Cantor; Cantores), conta uma vez só. A aba Variantes_para_aprovacao mostra a grafia original. Versão 1 salva como *_v1_sem_variantes."],
    ["Script", "solicitacao17.py (a Solicitação 17 anterior, Rostos × Referência Local, passou a usar solicitacao17_rostos.py)"],
], columns=["Item", "Descrição"])

# ---------------------------------------------------------------- gravação
with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t_rot.to_excel(w, sheet_name="Rotulos_por_post", index=False)
    t_mais3.to_excel(w, sheet_name="Rotulos_por_post", index=False, startrow=len(t_rot) + 3)
    for nome in PERFIS + ["Total"]:
        tops[nome].to_excel(w, sheet_name=f"Top30_{ABA[nome]}", index=False)
    t_freq.to_excel(w, sheet_name="Frequencia_dos_rotulos", index=False)
    t_eixo.to_excel(w, sheet_name="Por_eixo", index=False)
    t_var.to_excel(w, sheet_name="Variantes_para_aprovacao", index=False)
    tabela_correspondencia().to_excel(w, sheet_name="Variantes_aplicadas", index=False)
    t_irreg.to_excel(w, sheet_name="Variantes_para_aprovacao", index=False, startrow=len(t_var) + 3)
    notas.to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
for ws in wb.worksheets:
    for c in ws[1]:
        c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 24
    if ws.title == "Notas":
        ws.column_dimensions["B"].width = 120
        for row in ws.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")
        continue
    cab = {c.column: str(c.value) for c in ws[1]}
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, float):
                c.number_format = "0.0%" if "%" in cab.get(c.column, "") else "0.00"
            elif isinstance(c.value, int):
                c.number_format = "#,##0"
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 16
ws = wb["Rotulos_por_post"]
for c in ws[len(t_rot) + 1]:
    c.font = Font(bold=True)                                   # linha Total
r0 = len(t_rot) + 3
ws.cell(r0, 1, "Posts com mais de 3 rótulos (o livro prevê no máximo 3)").font = Font(bold=True)
for c in ws[r0 + 1]:
    c.font = Font(bold=True)
for row in ws.iter_rows(min_row=r0 + 2):
    for c in row:
        if isinstance(c.value, (int, float)):
            c.number_format = "0"                              # nº de linha, sem separador
ws.column_dimensions["E"].width = 70
for nome in PERFIS + ["Total"]:
    ws = wb[f"Top30_{ABA[nome]}"]; ws.column_dimensions["B"].width = 34
ws = wb["Por_eixo"]; ws.column_dimensions["A"].width = 32; ws.column_dimensions["C"].width = 32
ws = wb["Variantes_para_aprovacao"]
for col, wd in zip("ABCDEFGH", [30, 8, 30, 8, 34, 50, 22, 50]):
    ws.column_dimensions[col].width = wd
r1 = len(t_var) + 3
ws.cell(r1, 1, "Rótulos isolados com grafia irregular (sem par; só informados, nada alterado)").font = Font(bold=True)
for c in ws[r1 + 1]:
    c.font = Font(bold=True)
wb.save(SAIDA)

# ---------------------------------------------------------------- gráfico: 15 mais frequentes por perfil
cor = "#0072B2"
fig, axs = plt.subplots(1, 3, figsize=(16, 7.5))
for ax, p in zip(axs, PERFIS):
    d = df[df["Perfil"] == p]
    c = contar(d).head(15)
    pct = c["N de posts"] / len(d) * 100
    y = np.arange(len(c))[::-1]
    ax.barh(y, pct, color=cor, height=0.7)
    for yy, v, n in zip(y, pct, c["N de posts"]):
        ax.text(v + 0.4, yy, br_pct(v), va="center", fontsize=8.5)
    ax.set_yticks(y)
    ax.set_yticklabels([f"{t} *" if n < LIMIAR_TOP else t for t, n in zip(c["Temática"], c["N de posts"])], fontsize=9)
    ax.set_xlim(0, pct.max() * 1.22)
    ax.xaxis.set_major_formatter(lambda v_, _: f"{v_:.0f}%")
    ax.set_xlabel("% dos posts do perfil")
    ax.set_title(f"{p} (N = {br_int(len(d))})", fontsize=11, loc="left")   # nome do painel (não é título da figura)
    ax.spines[["top", "right"]].set_visible(False)
fig.text(0.99, 0.005, "Posts com o rótulo ÷ posts do perfil; um post pode ter vários rótulos. Variantes de grafia aprovadas (07/11/2026) aplicadas. Empates em ordem alfabética.",
         ha="right", fontsize=7.5, color="dimgray")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(PNG, dpi=300, bbox_inches="tight")

# ---------------------------------------------------------------- saída de conferência
pd.set_option("display.width", 220)
print(t_rot.round(3).to_string())
print(t_mais3.to_string())
for nome in PERFIS + ["Total"]:
    print(nome, "empates fora:", empates[nome]); print(tops[nome].head(10).round(3).to_string())
print(t_freq.round(3).to_string())
print(t_var["Tipo de semelhança"].value_counts().to_string())
print(t_var.iloc[:, [0,2,7]].to_string()); print(t_irreg.to_string())
