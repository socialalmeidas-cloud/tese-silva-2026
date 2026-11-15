"""Solicitação 13 — Hashtags e menções.
Fontes: hashtags/0..hashtags/27 e mentions/0..mentions/13 (uma por célula).
Hashtags: só a caixa é normalizada (minúsculas); grafias não são corrigidas.
Menções: mantidas como estão (nenhuma normalização).
Uso: python3 solicitacao13.py <planilha.xlsx>
Gera Solicitacao13_Hashtags_mencoes.xlsx.

Versão 2 (05/11/2026), aprovada pela autora: para o Imperatriz Online,
- menções vêm do arquivo Mencoes_Imperatriz_Online.xlsx (aba "Para colar"), enviado pela autora;
- hashtags vêm da aba "2024" da coleta COLETA_IMPERATRIZ_ONLINE_1.xlsx (cópia em io_coleta/coleta_io.xlsx).
Motivo: na planilha congelada, as colunas hashtags/ do IO juntam hashtags e menções, e as colunas mentions/ do IO estão vazias.
Ego e Roma seguem com as colunas da planilha (o arquivo de menções é idêntico para eles; conferido abaixo).
"""
import sys
import re
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao13_Hashtags_mencoes.xlsx"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
H = [f"hashtags/{i}" for i in range(28)]
M = [f"mentions/{i}" for i in range(14)]
TOP = 30

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
N_tot = {**N.to_dict(), "Total": len(df)}
for c in H + M + ["mentions"]:
    assert c in df.columns, c

# ---- Fontes corrigidas do Imperatriz Online (v2)
FONTE_MENCOES = "io_mencoes/mencoes_io.xlsx"
FONTE_HASHTAGS_IO = "io_coleta/coleta_io.xlsx"
IO = df["Perfil"] == "Imperatriz Online"
nv = pd.read_excel(FONTE_MENCOES, sheet_name="Para colar")
assert len(nv) == len(df) and (nv["Linha"].values == df.index.values + 2).all()
assert (nv["url"].astype(str).str.strip().values == df["url "].astype(str).str.strip().values).all()
assert (nv["Perfil"].values == df["Perfil"].values).all()
_S = lambda r: frozenset(str(x).strip() for x in r if pd.notna(x) and str(x).strip() != "")
assert (nv.loc[~IO, M].apply(_S, axis=1) == df.loc[~IO, M].apply(_S, axis=1)).all(), "Ego/Roma: menções diferentes da planilha"
assert df.loc[IO, M].isna().all().all() and df.loc[IO, "mentions"].isna().all()   # IO sem menções na planilha
df.loc[IO, M] = nv.loc[IO, M].values
df.loc[IO, "mentions"] = nv.loc[IO, "mentions"].values

col24 = pd.read_excel(FONTE_HASHTAGS_IO, sheet_name="2024")
HC = [c for c in col24.columns if str(c).startswith("hashtags/")]
col24["u"] = col24["url"].astype(str).str.strip()
assert (col24.groupby("u")[HC].nunique(dropna=False) <= 1).all().all()     # URLs repetidas têm o mesmo conteúdo
col24 = col24.drop_duplicates("u").set_index("u")
u_io = df.loc[IO, "url "].astype(str).str.strip()
assert u_io.isin(col24.index).all()                                          # os 2.721 posts do IO estão na coleta
h_io = col24.loc[u_io.values, HC].values
df.loc[IO, H] = np.nan
df.loc[IO, H[:len(HC)]] = h_io
N_POSTS_H_IO = int((pd.DataFrame(h_io).notna().any(axis=1)).sum())


def valores(row):
    return [str(x) for x in row if pd.notna(x) and str(x) != ""]


df["h_orig"] = df[H].apply(valores, axis=1)
df["h"] = df["h_orig"].apply(lambda l: sorted({x.lower() for x in l}))       # só caixa; uma vez por post
df["m"] = df[M].apply(lambda r: sorted(set(valores(r))), axis=1)              # sem normalização; uma vez por post

# ---- Conferência: coluna "mentions" × mentions/0..13
def tokens(x):
    if pd.isna(x):
        return []
    return sorted({t for t in re.split(r"[\s,]+", str(x)) if t})


tok = df["mentions"].apply(tokens)
igual_bruto = (tok == df["m"])
m_sem_virg = df["m"].apply(lambda l: sorted({x.rstrip(",") for x in l}))
igual_sem_virg = (tok == m_sem_virg)
assert igual_sem_virg.all(), "há diferenças além da vírgula final"
conf = pd.DataFrame({
    "url": df.loc[~igual_bruto, "url "],
    "Perfil": df.loc[~igual_bruto, "Perfil"],
    "Coluna \"mentions\" (texto)": df.loc[~igual_bruto, "mentions"],
    "mentions/0..13 (valores)": df.loc[~igual_bruto, "m"].apply(lambda l: " | ".join(l)),
})
n_vazio_ambos = int(((df["mentions"].isna()) & (df["m"].str.len() == 0)).sum())

# ---- Registro da normalização de caixa (hashtags)
todas = pd.Series([x for l in df["h_orig"] for x in l])
grupos = todas.groupby(todas.str.lower()).apply(lambda s: sorted(set(s)))
caixa = grupos[grupos.str.len() > 1]
t_caixa = pd.DataFrame({"Forma usada na contagem (minúsculas)": caixa.index,
                        "Grafias originais unidas (só diferença de caixa)": caixa.apply(lambda l: " | ".join(l)).values})


_hio = pd.Series([x for l in df.loc[IO, "h_orig"] for x in l])
_p = _hio[~_hio.str.fullmatch(r"[\wÀ-ÿ]+")]
pont = _p.value_counts().rename_axis("Hashtag com pontuação (como está na coleta)").reset_index(name="N de ocorrências")


def top(col, grupo):
    b = df if grupo == "Total" else df[df["Perfil"] == grupo]
    c = b[col].explode().dropna().value_counts().rename_axis("x").reset_index(name="n").sort_values(["n", "x"], ascending=[False, True])
    c["pos"] = c["n"].rank(method="min", ascending=False).astype(int)
    c = c[c["pos"] <= TOP].copy()
    rep = c["pos"].duplicated(keep=False)
    c["Posição"] = c["pos"].astype(str) + rep.map({True: "=", False: ""})
    c["%"] = c["n"] / N_tot[grupo]
    c["x"] = [x + (" *" if n < 10 else "") for x, n in zip(c["x"], c["n"])]
    nome = "Hashtag" if col == "h" else "Perfil mencionado"
    return c.rename(columns={"x": nome, "n": "N de posts que a usam" if col == "h" else "N de posts que o mencionam",
                             "%": "% dos posts do perfil"})[["Posição", nome, c.columns[1] if False else ("N de posts que a usam" if col == "h" else "N de posts que o mencionam"), "% dos posts do perfil"]].reset_index(drop=True)


tops = {(col, g): top(col, g) for col in ("h", "m") for g in PERFIS + ["Total"]}
for k, t in tops.items():  # perfil sem nenhum valor: registra a ausência na aba, em vez de deixá-la vazia
    if t.empty:
        fonte = "hashtags/0..27" if k[0] == "h" else "mentions/0..13 (e na coluna \"mentions\")"
        tops[k] = pd.DataFrame([["—", f"Nenhum valor registrado em {fonte} para este perfil ({N_tot[k[1]]:,} posts)".replace(",", "."), 0, 0.0]],
                               columns=t.columns)

# ---- Médias
def medidas(col, rotulo):
    rows = []
    for g in PERFIS + ["Total"]:
        b = df if g == "Total" else df[df["Perfil"] == g]
        n = b[col].str.len()
        rows.append([g, N_tot[g], int((n == 0).sum()), (n == 0).mean(), n.mean(), n.median(), int(n.max()),
                     n[n > 0].mean() if (n > 0).any() else np.nan])
    return pd.DataFrame(rows, columns=["Perfil", "N de posts", f"Posts sem {rotulo} (N)", f"Posts sem {rotulo} (%)",
                                       f"Média de {rotulo}s por post", f"Mediana de {rotulo}s por post",
                                       f"Máximo de {rotulo}s em um post", f"Média entre posts com {rotulo}"])


med_h = medidas("h", "hashtag")
med_m = medidas("m", "menção")
med_m.columns = [c.replace("menções", "menções").replace("mençãos", "menções") for c in med_m.columns]

n_arroba_io = int(df.loc[df["Perfil"] == "Imperatriz Online", "Legenda"].astype(str).str.contains("@").sum())
notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma."],
    ["Fontes (versão 2, 05/11/2026)", "Ego e Roma: colunas hashtags/0..27 e mentions/0..13 da planilha. Imperatriz Online: menções do arquivo Mencoes_Imperatriz_Online.xlsx, aba \"Para colar\" (enviado pela autora; para Ego e Roma é idêntico à planilha, conferido pelo script); hashtags da aba \"2024\" da coleta COLETA_IMPERATRIZ_ONLINE_1.xlsx (aprovado pela autora). Motivo: na planilha congelada, as colunas hashtags/ do IO juntavam hashtags e menções, e as colunas mentions/ do IO estavam vazias. A versão anterior foi guardada como Solicitacao13_Hashtags_mencoes_v1_antes_fonte_IO.xlsx."],
    ["Coluna \"mentions\"", f"Repete o conteúdo de mentions/0..13 em texto corrido (separado por espaço ou vírgula). Em 4.258 posts o conteúdo é idêntico; nos {len(conf)} restantes a única diferença é uma vírgula no fim de um valor dentro de mentions/0..13 (ex.: \"executivahousemusic,\"), que na coluna \"mentions\" é separador. Posts sem nenhuma menção nas duas fontes: {n_vazio_ambos}. A coluna \"mentions\" não foi usada na contagem. Ver aba Conferencia_mentions."],
    ["Normalização das hashtags", f"Só a caixa: tudo convertido para minúsculas antes de contar. Hashtags distintas: {todas.nunique()} antes e {todas.str.lower().nunique()} depois ({len(caixa)} grupos uniram grafias que só diferiam na caixa; ver aba Normalizacao_caixa). Acentos, abreviações e erros de digitação NÃO foram corrigidos."],
    ["Imperatriz Online", f"Menções: {int((df.loc[IO, 'm'].str.len() > 0).sum())} posts com pelo menos uma. No arquivo da autora, a pontuação colada ao nome na coleta (ex.: \"studiosalasar!\", \"jupitertelecom!✅\") já vem limpa. Hashtags: {N_POSTS_H_IO} posts com pelo menos uma (aba \"2024\" da coleta; URLs repetidas na coleta, com o mesmo conteúdo, contadas uma vez)."],
    ["Pontuação nas hashtags do IO", "Na coleta, algumas hashtags do IO têm pontuação colada ao fim (ex.: \"ServiçoDeQualidade”\", \"Hipopressivo,\", \"fisioterapia!\"). Seguindo a regra de não agrupar variantes sem aprovação, elas foram mantidas como estão (só a caixa foi normalizada); por isso \"fisioterapia!\" e \"fisioterapia\" contam separado. Ver aba Pontuacao_hashtags_IO. A limpeza depende de aprovação da autora."],
    ["Menções", "Nenhuma normalização (já estavam em minúsculas). As 8 vírgulas finais foram mantidas como estão (não corrigidas): por isso \"x,\" e \"x\" contam como perfis diferentes. Os perfis podem mencionar a si mesmos (ex.: egoimperatriz no Ego); essas menções foram mantidas."],
    ["Contagem", "Cada hashtag (ou perfil mencionado) conta no máximo uma vez por post. % = nº de posts do perfil que usam a hashtag (ou mencionam o perfil) ÷ total de posts do perfil × 100. A soma passa de 100% (um post pode ter várias)."],
    ["Top 30 e empates", "Ranking de competição; \"=\" indica empate; empates na 30ª posição foram todos incluídos (algumas listas têm mais de 30 linhas)."],
    ["*", "Hashtag ou perfil com menos de 10 posts no grupo."],
    ["Medias", "Contagens por post já sem repetições. Média e mediana incluem posts sem nenhuma hashtag/menção (valor 0); a última coluna considera só posts com pelo menos uma."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a hashtags ou menções."],
    ["Script", "solicitacao13.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    for col, pref in (("h", "Hashtags_"), ("m", "Mencoes_")):
        for g in PERFIS + ["Total"]:
            tops[(col, g)].to_excel(w, sheet_name=pref + g.replace(" ", "_"), index=False)
    med_h.to_excel(w, sheet_name="Medias", index=False, startrow=1)
    med_m.to_excel(w, sheet_name="Medias", index=False, startrow=len(med_h) + 4)
    t_caixa.to_excel(w, sheet_name="Normalizacao_caixa", index=False)
    pont.to_excel(w, sheet_name="Pontuacao_hashtags_IO", index=False)
    conf.to_excel(w, sheet_name="Conferencia_mentions", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
for ws in wb.worksheets:
    if ws.title.startswith(("Hashtags_", "Mencoes_")):
        for c in ws[1]:
            c.font = Font(bold=True)
        for row in ws.iter_rows(min_row=2):
            row[2].number_format = "#,##0"
            row[3].number_format = "0.0%"
        for L, wd in zip("ABCD", (10, 34, 26, 22)):
            ws.column_dimensions[L].width = wd
ws = wb["Medias"]
ws["A1"] = "Hashtags"
ws.cell(row=len(med_h) + 4, column=1, value="Menções")
for r in (1, 2, len(med_h) + 4, len(med_h) + 5):
    for c in ws[r]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
for row in ws.iter_rows():
    if isinstance(row[1].value, (int, float)):
        row[1].number_format = row[2].number_format = row[6].number_format = "#,##0"
        row[3].number_format = "0.0%"
        row[4].number_format = row[5].number_format = row[7].number_format = "0.00"
for L, wd in zip("ABCDEFGH", (20, 12, 16, 16, 16, 16, 16, 18)):
    ws.column_dimensions[L].width = wd
for nome, wds in (("Normalizacao_caixa", (40, 70)), ("Conferencia_mentions", (40, 18, 60, 60))):
    ws = wb[nome]
    for c in ws[1]:
        c.font = Font(bold=True)
    for L, wd in zip("ABCD", wds):
        ws.column_dimensions[L].width = wd
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 26
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

pd.set_option("display.width", 220)
for col in ("h", "m"):
    for g in PERFIS + ["Total"]:
        t = tops[(col, g)]
        print(col, g, len(t)); print(t.head(10).to_string(index=False))
print(med_h.to_string()); print(med_m.to_string()); print(t_caixa.head(20).to_string()); print(len(conf))
