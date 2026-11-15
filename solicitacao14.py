"""Solicitação 14 — Corpus de legendas para o IRaMuTeQ.

Prepara o arquivo de entrada do IRaMuTeQ a partir da planilha de codificação.
A análise lexical em si é feita no IRaMuTeQ. Este script não foi executado sobre
os dados da tese: ele acompanha o repositório para quem quiser replicar o
procedimento com os próprios dados.

Duas etapas:

  1) Relatório prévio. Não gera o arquivo final. Mostra quantas legendas entram,
     quantas estão vazias e quantas são duplicadas (texto idêntico) e exibe
     5 exemplos antes e depois da limpeza, para aprovação do critério.

        python3 -I solicitacao14.py planilha.xlsx --etapa relatorio

  2) Arquivo final, depois de aprovado o critério:

        python3 -I solicitacao14.py planilha.xlsx --etapa arquivo \
            [--duplicadas manter|uma] [--hashtags remover|palavra]

     Gera:
       legendas_iramuteq.txt  (UTF-8; cada legenda = linha de variáveis + texto em uma linha)
       Mapa_variaveis.xlsx    (código usado para cada perfil, eixo, mediação e mês, com o N
                               de posts e uma aba "Notas" com o critério de limpeza e o N final)

Linha de variáveis:
  **** *perfil_<perfil> *eixo_<eixo> *mediacao_<mediação> *mes_<AAAA-MM>
(códigos sem espaços, sem acentos e sem caracteres especiais; o mês usa hífen,
como definido na solicitação; confira se a sua versão do IRaMuTeQ o aceita.)

Critério de limpeza padrão (ajuste as funções abaixo se o seu critério for outro):
  links, menções (@perfil), hashtags (#palavra), emojis e símbolos, aspas,
  apóstrofos e asteriscos são removidos; quebras de linha e espaços repetidos viram
  um espaço. Acentos, pontuação e maiúsculas são mantidos. O asterisco precisa sair
  do texto porque o IRaMuTeQ o usa para marcar variáveis.

Os nomes das colunas podem ser trocados pelas opções --col-*. O período pode estar
no formato "(dez-2024)" (como na planilha da tese) ou ser uma data.
"""
import argparse
import re
import sys
import unicodedata

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
         "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}

ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument("planilha")
ap.add_argument("--etapa", choices=["relatorio", "arquivo"], default="relatorio")
ap.add_argument("--duplicadas", choices=["manter", "uma"], default="manter",
                help="manter todas as legendas idênticas (padrão) ou só uma por grupo")
ap.add_argument("--hashtags", choices=["remover", "palavra"], default="remover",
                help="remover a hashtag inteira (padrão) ou manter a palavra sem o #")
ap.add_argument("--col-legenda", default="Legenda")
ap.add_argument("--col-perfil", default="Perfil")
ap.add_argument("--col-eixo", default="Eixo Temático")
ap.add_argument("--col-mediacao", default="Mediação da Informação")
ap.add_argument("--col-periodo", default="Período")
ap.add_argument("--saida-txt", default="legendas_iramuteq.txt")
ap.add_argument("--saida-mapa", default="Mapa_variaveis.xlsx")
ap.add_argument("--semente", type=int, default=14, help="semente do sorteio dos exemplos")
A = ap.parse_args()


def br(n, casas=0):
    """Número no padrão brasileiro (só para exibição no terminal)."""
    s = f"{n:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def eh_emoji(c):
    return (unicodedata.category(c) in ("So", "Sk", "Cs", "Co", "Cn")
            or ord(c) >= 0x1F000 or c in "‍️⃣")


def limpar(s):
    s = re.sub(r"https?://\S+|www\.\S+", " ", s)                  # links
    s = re.sub(r"(?<!\w)@[\w.]+", " ", s)                        # menções
    if A.hashtags == "remover":
        s = re.sub(r"#\w+", " ", s)                               # hashtag inteira
    else:
        s = re.sub(r"#(\w+)", r"\1", s)                           # só o símbolo
    s = "".join(" " if eh_emoji(c) else c for c in s)             # emojis e símbolos
    s = re.sub(r"[\"“”'‘’`´*]", " ", s)                           # aspas, apóstrofos, asterisco
    s = re.sub(r"\s+", " ", s).strip()                            # quebras de linha e espaços
    return s


def codigo(rotulo):
    """'Período Eleitoral' -> 'PeriodoEleitoral' (sem acentos, espaços ou símbolos)."""
    s = unicodedata.normalize("NFKD", str(rotulo)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]", "", s.title() if " " in s else s) or "NA"


def mes(valor):
    v = str(valor).strip().lower()
    m = re.search(r"([a-zç]{3})\w*[-/ ](\d{4})", v)
    if m and m.group(1)[:3] in MESES:
        return f"{m.group(2)}-{MESES[m.group(1)[:3]]:02d}"
    d = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.isna(d):
        raise ValueError(f"período não reconhecido: {valor!r}")
    return f"{d.year}-{d.month:02d}"


# --- leitura -------------------------------------------------------------------
df = pd.read_excel(A.planilha)
for c in (A.col_legenda, A.col_perfil, A.col_eixo, A.col_mediacao, A.col_periodo):
    if c not in df.columns:
        sys.exit(f"coluna ausente na planilha: {c!r} (use as opções --col-*)")
df["Linha"] = df.index + 2
leg = df[A.col_legenda]
df["vazia"] = leg.isna() | leg.astype(str).str.strip().isin(["", "nan"])
df["texto"] = leg.where(~df["vazia"], "").astype(str)
df["limpo"] = df["texto"].map(limpar)
df["dup"] = ~df["vazia"] & df["texto"].duplicated(keep=False)
PERFIS = list(dict.fromkeys(df[A.col_perfil]))

# --- etapa 1: relatório prévio --------------------------------------------------
if A.etapa == "relatorio":
    print(f"Planilha: {A.planilha} | posts: {br(len(df))}\n")
    linhas = []
    for p in PERFIS + ["Total"]:
        s = df if p == "Total" else df[df[A.col_perfil] == p]
        d = s[s["dup"]]
        linhas.append({"Perfil": p, "Posts": len(s), "Com legenda": int((~s["vazia"]).sum()),
                       "Vazias": int(s["vazia"].sum()), "Duplicadas (posts)": len(d),
                       "Grupos de duplicadas": d["texto"].nunique()})
    print(pd.DataFrame(linhas).to_string(index=False))
    g = df[df["dup"]]
    print(f"\nSe mantida uma legenda por grupo, saem {br(len(g) - g['texto'].nunique())} posts."
          f" Grupos que atravessam perfis: {int((g.groupby('texto')[A.col_perfil].nunique() > 1).sum())}.")
    print("Maiores grupos de legendas idênticas:")
    for t, n in g["texto"].value_counts().head(5).items():
        print(f"  {n}x | {t[:90]!r}")
    com = df[~df["vazia"]]["texto"]
    print("\nElementos encontrados (legendas):")
    for nome, rx in [("links", r"https?://\S+|www\.\S+"), ("menções", r"(?<!\w)@[\w.]+"),
                     ("hashtags", r"#\w+"), ("quebras de linha", r"\n"), ("aspas", r"[\"“”'‘’]"),
                     ("asteriscos", r"\*")]:
        print(f"  {nome}: {br(int(com.str.contains(rx).sum()))}")
    print(f"  emojis/símbolos: {br(int(com.map(lambda s: any(eh_emoji(c) for c in s)).sum()))}")
    print(f"  hashtag no meio da frase: {br(int(com.str.contains(r'#[^\s#]+\s+[^\s#@]').sum()))}")
    vaz = int((~df["vazia"] & df["limpo"].eq("")).sum())
    curtas = int((~df["vazia"] & df["limpo"].str.split().str.len().lt(3)).sum())
    print(f"\nApós a limpeza: {br(vaz)} legendas ficam vazias; {br(curtas)} ficam com menos de 3 palavras.")
    ex = df[~df["vazia"] & df["texto"].str.len().between(80, 400)].drop_duplicates("texto")
    ex = ex.sample(min(5, len(ex)), random_state=A.semente)
    print("\n5 exemplos (antes / depois):")
    for _, r in ex.iterrows():
        print(f"\n[linha {r['Linha']} | {r[A.col_perfil]}]\nANTES:  {r['texto']!r}\nDEPOIS: {r['limpo']}")
    print("\nNada foi gravado. Aprovado o critério, rode com --etapa arquivo.")
    sys.exit(0)

# --- etapa 2: arquivo final -----------------------------------------------------
base = df[~df["vazia"] & df["limpo"].ne("")].copy()
excl_vazias = int(df["vazia"].sum())
excl_limpeza = int((~df["vazia"] & df["limpo"].eq("")).sum())
excl_dup = 0
if A.duplicadas == "uma":
    antes = len(base)
    base = base.drop_duplicates("texto", keep="first")
    excl_dup = antes - len(base)

base["c_perfil"] = base[A.col_perfil].map(codigo)
base["c_eixo"] = base[A.col_eixo].map(codigo)
base["c_med"] = base[A.col_mediacao].map(codigo)
base["c_mes"] = base[A.col_periodo].map(mes)
for orig, cod in ((A.col_perfil, "c_perfil"), (A.col_eixo, "c_eixo"), (A.col_mediacao, "c_med")):
    choque = base.groupby(cod)[orig].nunique()
    assert (choque == 1).all(), f"rótulos diferentes geraram o mesmo código: {list(choque[choque > 1].index)}"

with open(A.saida_txt, "w", encoding="utf-8", newline="\n") as f:
    for _, r in base.iterrows():
        f.write(f"**** *perfil_{r['c_perfil']} *eixo_{r['c_eixo']} *mediacao_{r['c_med']} *mes_{r['c_mes']}\n")
        f.write(r["limpo"] + "\n")

# conferência do arquivo gravado
txt = open(A.saida_txt, encoding="utf-8").read().splitlines()
assert len(txt) == 2 * len(base)
assert all(l.startswith("**** ") for l in txt[0::2]) and not any("*" in l for l in txt[1::2])


def tabela(nome, orig, cod, por_codigo=False):
    t = (base.groupby([orig, cod]).size().reset_index(name="N de posts")
         .rename(columns={orig: "Nome original", cod: "Código no arquivo"})
         .sort_values("Código no arquivo" if por_codigo else "N de posts", ascending=por_codigo))
    t.insert(0, "Variável", nome)
    t["Código no arquivo"] = f"*{nome}_" + t["Código no arquivo"]
    t["Nota"] = t["N de posts"].map(lambda n: "*" if n < 10 else "")
    total = pd.DataFrame([{"Variável": nome, "Nome original": "Total", "Código no arquivo": "",
                           "N de posts": int(t["N de posts"].sum()), "Nota": ""}])
    out = pd.concat([t, total], ignore_index=True)
    assert out.iloc[-1]["N de posts"] == len(base)
    return out


tabs = {"Perfil": tabela("perfil", A.col_perfil, "c_perfil"),
        "Eixo": tabela("eixo", A.col_eixo, "c_eixo"),
        "Mediacao": tabela("mediacao", A.col_mediacao, "c_med"),
        "Mes": tabela("mes", A.col_periodo, "c_mes", por_codigo=True)}
crit = ("Removidos: links, menções (@perfil), "
        + ("hashtags inteiras" if A.hashtags == "remover" else "símbolo # das hashtags (palavra mantida)")
        + ", emojis e símbolos, aspas, apóstrofos e asteriscos; quebras de linha e espaços repetidos"
          " substituídos por um espaço. Mantidos: acentos, pontuação e maiúsculas.")
notas = pd.DataFrame([
    ["Arquivo de entrada", A.planilha],
    ["Posts na planilha", len(df)],
    ["Legendas vazias (excluídas)", excl_vazias],
    ["Legendas vazias após a limpeza (excluídas)", excl_limpeza],
    ["Duplicadas", "mantidas todas" if A.duplicadas == "manter"
     else f"mantida uma por grupo de texto idêntico ({excl_dup} excluídas)"],
    ["Critério de limpeza", crit],
    ["N final (legendas no arquivo)", len(base)],
    ["Formato", "UTF-8; linha de variáveis '**** *perfil_ *eixo_ *mediacao_ *mes_AAAA-MM' e, na linha seguinte, o texto limpo em uma só linha."],
    ["Asterisco (*)", "códigos com menos de 10 posts."],
], columns=["Item", "Valor"])

with pd.ExcelWriter(A.saida_mapa, engine="openpyxl") as w:
    for nome, t in tabs.items():
        t.to_excel(w, sheet_name=nome, index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)
wb = load_workbook(A.saida_mapa)
for ws in wb.worksheets:
    for c in ws[1]:
        c.font = Font(bold=True)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(60, max(12, max(len(str(c.value or "")) for c in col) + 2))
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, int):
                c.number_format = "#,##0"
wb.save(A.saida_mapa)

print(f"{A.saida_txt}: {br(len(base))} legendas | {A.saida_mapa}: códigos de perfil, eixo, mediação e mês")
for nome, t in tabs.items():
    print(f"\n{nome} (10 primeiras linhas):")
    print(t.head(10).to_string(index=False))
    print(f"  soma = {br(int(t.iloc[-1]['N de posts']))} (N final = {br(len(base))})")
