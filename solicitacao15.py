"""Solicitação 15 — Validação do recorte da Descrição da Imagem.
Compara, por perfil, o recorte (Descrição ≠ "Não analisado") com o corpus completo do perfil.
Uso: python3 solicitacao15.py <planilha.xlsx>
Gera Solicitacao15_Validacao_recorte.xlsx.
"""
import sys
import re
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao15_Validacao_recorte.xlsx"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
DESC, EIXO, MED, REG = "Descrição da Imagem", "Eixo Temático", "Mediação da Informação", "Regime de Visibilidade"
NA = "Não analisado"
EMERG = {"Criminoso", "Animal", "Bens materiais", "Candidatos", "Outro perfil"}
MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
INV = {v: k for k, v in MESES.items()}

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
from correcoes_planilha import aplicar, CORRECOES
df = aplicar(df)  # correções aprovadas pela autora (a planilha original não é alterada)
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}
df["linha_excel"] = df.index + 2  # linha na planilha original (linha 1 = cabeçalho)
df["rec"] = df[DESC].astype(str).str.strip() != NA
n_rec = df[df["rec"]]["Perfil"].value_counts()
assert n_rec.to_dict() == {"Imperatriz Online": 756, "Roma": 551, "Ego": 298}, n_rec.to_dict()
p = df["Período"].astype(str).str.strip().str.strip("()").str.split("-", expand=True)
df["data"] = pd.to_datetime(dict(year=p[1].astype(int), month=p[0].str.lower().map(MESES), day=1))
df["mes"] = df["data"].apply(lambda d: f"{INV[d.month]}/{d.year}")
df["reg"] = df[REG].astype(str).str.split(";").apply(lambda l: sorted({x.strip() for x in l if x.strip()}))
ORDEM_MES = [f"{INV[d.month]}/{d.year}" for d in sorted(df["data"].unique())]


def comparar(perfil, col, rotulos, multi=False, fechar=True):
    b = df[df["Perfil"] == perfil]
    r = b[b["rec"]]
    Nc, nr = len(b), len(r)
    if multi:
        cc = b[col].explode().value_counts()
        cr = r[col].explode().value_counts()
    else:
        cc = b[col].value_counts()
        cr = r[col].value_counts()
    rows = []
    for l in rotulos:
        a, c = int(cc.get(l, 0)), int(cr.get(l, 0))
        rows.append([l + (" †" if l in EMERG else ""), a, a / Nc, c, c / nr])
    if fechar:
        resto_c = Nc - sum(x[1] for x in rows)
        resto_r = nr - sum(x[3] for x in rows)
        if resto_c or resto_r:
            rows.append(["Outros", resto_c, resto_c / Nc, resto_r, resto_r / nr])
        rows.append(["Total", Nc, 1.0, nr, 1.0])
    else:
        rows.append(["N de posts", Nc, np.nan, nr, np.nan])
    t = pd.DataFrame(rows, columns=["Rótulo", "Corpus (N)", "Corpus (%)", "Recorte (N)", "Recorte (%)"])
    t["Diferença (p.p.)"] = (t["Recorte (%)"] - t["Corpus (%)"]) * 100
    t["Rótulo"] = [x + (" *" if (n < 10 and x not in ("Total", "N de posts", "Outros")) else "") for x, n in zip(t["Rótulo"], t["Recorte (N)"])]
    if fechar:
        assert t.iloc[:-1]["Corpus (N)"].sum() == Nc and t.iloc[:-1]["Recorte (N)"].sum() == nr
    return t


def top10(col, multi=False):
    s = df[col].explode() if multi else df[col]
    c = s.value_counts().rename_axis("r").reset_index(name="n").sort_values(["n", "r"], ascending=[False, True])
    return c["r"].iloc[:10].tolist()


EIXOS, MEDS, REGS = top10(EIXO), top10(MED), top10("reg", multi=True)
abas = {
    "Periodo": {pf: comparar(pf, "mes", ORDEM_MES).rename(columns={"Rótulo": "Mês/Ano"}) for pf in PERFIS},
    "Eixo": {pf: comparar(pf, EIXO, EIXOS).rename(columns={"Rótulo": "Eixo"}) for pf in PERFIS},
    "Mediacao": {pf: comparar(pf, MED, MEDS).rename(columns={"Rótulo": "Mediação"}) for pf in PERFIS},
    "Regime": {pf: comparar(pf, "reg", REGS, multi=True, fechar=False).rename(columns={"Rótulo": "Regime"}) for pf in PERFIS},
}

# ---- Pesos
pesos = pd.DataFrame([[pf, int(N[pf]), int(n_rec[pf]), n_rec[pf] / N[pf], N[pf] / n_rec[pf]] for pf in PERFIS] +
                     [["Total", len(df), int(df["rec"].sum()), df["rec"].mean(), len(df) / df["rec"].sum()]],
                     columns=["Perfil", "N do corpus", "n do recorte", "Fração amostral (%)", "Peso de desenho (N/n)"])

# ---- Descrições curtas
conta = lambda s: len(re.findall(r"\b[\wÀ-ÿ]+\b", str(s)))
rec = df[df["rec"]].copy()
rec["palavras"] = rec[DESC].apply(conta)
curtas = rec[rec["palavras"] < 20][["linha_excel", "Perfil", DESC, "palavras"]].sort_values(["palavras", "linha_excel"])
curtas.columns = ["Linha (na planilha original)", "Perfil", "Descrição", "Nº de palavras"]
dist_pal = rec.groupby("Perfil")["palavras"].describe()[["min", "25%", "50%", "mean", "max"]].reindex(PERFIS)

# ---- Codificação das demais variáveis: quem codificou os posts do recorte × fora do recorte
cod = pd.crosstab([df["Perfil"], df["Responsável"]], df["rec"].map({True: "Recorte (n)", False: "Fora do recorte (n)"})).reindex(
    columns=["Recorte (n)", "Fora do recorte (n)"], fill_value=0).reset_index()
cod = cod.sort_values(["Perfil", "Recorte (n)"], ascending=[True, False])
suspeitas = curtas  # (mantido para conferência)[curtas["Descrição"].astype(str).str.contains(r"Ação/visita|Redirecionamento|Contato", case=False)]

maxdif = {aba: {pf: t.iloc[:-1]["Diferença (p.p.)"].abs().max() for pf, t in d.items()} for aba, d in abas.items()}

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N usado", "Corpus: 4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721). Recorte: 1.605 posts com Descrição da Imagem diferente de \"Não analisado\" (Ego 298; Roma 551; Imperatriz Online 756). \"Não analisado\" (2.661 posts) não foi contado como dado."],
    ["Comparação", "Cada perfil é comparado com o próprio corpus. % do corpus = N do rótulo no corpus do perfil ÷ N do corpus do perfil × 100; % do recorte = n do rótulo no recorte do perfil ÷ n do recorte do perfil × 100. Diferença (p.p.) = % do recorte − % do corpus (positivo = rótulo mais frequente no recorte)."],
    ["Rótulos", "Eixo e Regime: os 10 rótulos mais frequentes do corpus geral (mesmos nas três tabelas). Mediação: as 5 categorias. Eixo e Mediação têm linha \"Outros\" (quando há) e \"Total\", e cada coluna soma 100%."],
    ["Regime (multirrótulo)", "Contam-se posts com o rótulo (célula separada por \";\"); a soma passa de 100% e não há linha de total, só \"N de posts\"."],
    ["Período", "Coluna Período no formato \"(mmm-aaaa)\", convertida para ordem cronológica."],
    ["Pesos", "Fração amostral = n do recorte ÷ N do corpus do perfil. Peso de desenho = N ÷ n (cada post do recorte representa N/n posts do perfil). Para estimar proporções de um perfil a partir do recorte, o recorte proporcional por mês dispensa pesos dentro do perfil; os pesos são necessários para juntar os três perfis num total (pois as frações amostrais diferem)."],
    ["Descrições curtas", "Posts do recorte com menos de 20 palavras na Descrição da Imagem. Palavras contadas como sequências de letras/números (pontuação ignorada). \"Linha\" = número da linha na planilha original (linha 1 = cabeçalho)."],
    ["Formato", "Números no padrão brasileiro (exibição); valores gravados como números. * = rótulo com menos de 10 posts no recorte do perfil. † = categoria emergente (Criminoso, Animal, Bens materiais, Candidatos, Outro perfil em Regime): aplicada de forma desigual entre lotes; em Criminoso, 0% a 71% dos posts de Criminalidade e Segurança Pública conforme o lote; resultados indicativos."],
    ["Lista completa de emergentes", "Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix."],
    ["ATENÇÃO — codificadora", "Os posts do recorte tiveram TODAS as variáveis (não só a Descrição) codificadas pela pesquisadora (sozinha ou em dupla com outro codificador), e os posts fora do recorte por outras codificadoras (ver aba Codificadora_recorte). Assim, as diferenças entre recorte e corpus nas abas Eixo, Mediacao e Regime misturam dois efeitos: o da amostragem e o da codificadora. A planilha não permite separá-los. O Período não é afetado (foi a variável de estratificação). Confirmado pela autora em 05/11/2026: no recorte, todas as variáveis foram codificadas de novo por ela (os códigos anteriores dessas linhas foram substituídos)."],
    ["Correção aplicada", "Linha 1502 (Roma, id_texte_1089): a Descrição da Imagem era \"Ação/visita; Redirecionamento; contato\" (cópia do CTA). Substituída pela descrição enviada pela autora em 05/11/2026 (arquivo correcoes_planilha.py). A planilha congelada não foi alterada."],
    ["Script", "solicitacao15.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    for aba, d in abas.items():
        r = 0
        for pf in PERFIS:
            pd.DataFrame([[f"{pf} — corpus N = {N[pf]:,}; recorte n = {n_rec[pf]:,}".replace(",", ".")]]).to_excel(
                w, sheet_name=aba, index=False, header=False, startrow=r)
            d[pf].to_excel(w, sheet_name=aba, index=False, startrow=r + 1)
            r += len(d[pf]) + 4
    pesos.to_excel(w, sheet_name="Pesos", index=False)
    curtas.to_excel(w, sheet_name="Descricoes_curtas", index=False)
    cod.to_excel(w, sheet_name="Codificadora_recorte", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

amarelo = PatternFill("solid", fgColor="FFF2CC")
wb = load_workbook(SAIDA)
for aba in abas:
    ws = wb[aba]
    for row in ws.iter_rows():
        v0 = str(row[0].value)
        if row[1].value is None and row[0].value is not None or v0 in ("Mês/Ano", "Eixo", "Mediação", "Regime", "Total", "N de posts"):
            for c in row:
                c.font = Font(bold=True)
        if isinstance(row[1].value, (int, float)):
            row[1].number_format = row[3].number_format = "#,##0"
            row[2].number_format = row[4].number_format = "0.0%"
            row[5].number_format = "+0.0;-0.0;0.0"
        if "†" in v0:
            for c in row[:6]:
                c.fill = amarelo
    for L, wd in zip("ABCDEF", (30, 12, 12, 12, 12, 16)):
        ws.column_dimensions[L].width = wd
ws = wb["Pesos"]
for c in ws[1] + ws[ws.max_row]:
    c.font = Font(bold=True)
for row in ws.iter_rows(min_row=2):
    row[1].number_format = row[2].number_format = "#,##0"
    row[3].number_format = "0.0%"
    row[4].number_format = "0.000"
for L, wd in zip("ABCDE", (20, 14, 14, 20, 22)):
    ws.column_dimensions[L].width = wd
ws = wb["Descricoes_curtas"]
for c in ws[1]:
    c.font = Font(bold=True)
for L, wd in zip("ABCD", (14, 18, 100, 14)):
    ws.column_dimensions[L].width = wd
for row in ws.iter_rows(min_row=2):
    row[2].alignment = Alignment(wrap_text=True, vertical="top")
ws = wb["Codificadora_recorte"]
for c in ws[1]:
    c.font = Font(bold=True)
for L, wd in zip("ABCD", (20, 32, 14, 20)):
    ws.column_dimensions[L].width = wd
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 26
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
    if str(row[0].value).startswith(("ATENÇÃO", "Correção aplicada")):
        for c in row:
            c.fill = amarelo
            c.font = Font(bold=True)
wb.save(SAIDA)

pd.set_option("display.width", 220)
print(pesos.to_string())
print({a: {k: round(v, 1) for k, v in d.items()} for a, d in maxdif.items()})
for aba in abas:
    for pf in PERFIS:
        t = abas[aba][pf]
        print(aba, pf); print(t.assign(**{"Corpus (%)": (t["Corpus (%)"] * 100).round(1), "Recorte (%)": (t["Recorte (%)"] * 100).round(1),
                                         "Diferença (p.p.)": t["Diferença (p.p.)"].round(1)}).to_string(index=False))
print(len(curtas), curtas["Perfil"].value_counts().to_dict()); print(dist_pal.round(1).to_string())
print(curtas.head(8).to_string()); print(cod.to_string())
