"""
Solicitação 23 — Como o campo social é retratado nos perfis (versão 2).

Uso:  python3 -I solicitacao23.py copia.xlsx

Delimitação para LOCALIZAR posts (confirmada pela pesquisadora em 08/11/2026):
- Eixos: Solidariedade, Comunidade, Saúde, Educação, Denúncia, Protesto, Obituário, Justiça;
- Regime de Visibilidade "Pessoas comuns" (multirrótulo);
- Conteúdo Sensível: qualquer valor diferente de "Não há" (Acidente, Corpo ou ferimento,
  Pessoa detida, Violência, Procedimento †).
A delimitação só localiza posts e confere a leitura (Etapa 3). A base da análise (80%) é a
leitura de TODAS as descrições do recorte de cada perfil (Ego 298; Roma 551; Imperatriz Online 756),
em lotes de até 60 (arquivos s23b_<perfil>.txt, gerados abaixo).
Denominadores das variáveis: posts do perfil (Ego 414; Roma 1.131; Imperatriz Online 2.721).
A planilha original não é alterada (lê-se a cópia, com as correções aprovadas).
A versão 1 (leitura só dos eixos) foi guardada como v1_*.
"""
import os
import re
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
import correcoes_planilha

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
EIXOS = ["Solidariedade", "Comunidade", "Saúde", "Educação", "Denúncia", "Protesto", "Obituário", "Justiça"]
EMERGENTES = {"Procedimento"}            # categorias emergentes (†), indicativas
PESOS = {"Ego": 1.39, "Roma": 2.05, "Imperatriz Online": 3.60}


def rotulos(c):
    """Multirrótulo em ";" (o ";" dentro de parênteses não separa), sem 'Não há'/'Não analisado'."""
    if not isinstance(c, str):
        return []
    c = re.sub(r"\(([^)]*)\)", lambda m: "(" + m.group(1).replace(";", "§") + ")", c)
    ts = [t.strip().replace("§", ";") for t in c.split(";") if t.strip() and t.strip() not in ("Não há", "Não analisado")]
    return list(dict.fromkeys(ts))


df = correcoes_planilha.aplicar(pd.read_excel(ARQ))
assert len(df) == 4266
df["Linha"] = df.index + 2
NP = df["Perfil"].value_counts().to_dict()
DESC = json.load(open("desc_recorte.json", encoding="utf8"))
df["recorte"] = df["Linha"].astype(str).isin(DESC)
df["reg"] = df["Regime de Visibilidade"].apply(rotulos)
df["sens_l"] = df["Conteúdo Sensível"].apply(rotulos)
df["no_eixo"] = df["Eixo Temático"].isin(EIXOS)
df["pc"] = df["reg"].apply(lambda x: "Pessoas comuns" in x)
df["sens"] = df["sens_l"].str.len() > 0
df["delim"] = df["no_eixo"] | df["pc"] | df["sens"]

# ------------------------------------------------ arquivos de leitura (lotes de até 60, todo o recorte)
for p in PERFIS:
    s = df[df.recorte & (df.Perfil == p)]
    with open(f"s23b_{p.split()[0]}.txt", "w", encoding="utf8") as f:
        for i, (_, r) in enumerate(s.iterrows()):
            if i % 60 == 0:
                f.write(f"===== LOTE {i // 60 + 1}\n")
            f.write(f"[{r.Linha}|{r['Eixo Temático']}|{r['Regime de Visibilidade']}|{r['Conteúdo Sensível']}] {DESC[str(r.Linha)]}\n")

# ================================================ ETAPA 3: apoio das variáveis
def linha(nome, mask):
    return {"Categoria": nome, **{c: v for p in PERFIS for c, v in
            ((f"{p} (N)", int((mask & (df.Perfil == p)).sum())), (f"{p} (%)", (mask & (df.Perfil == p)).sum() / NP[p]))}}

lin = [linha(e, df["Eixo Temático"] == e) for e in EIXOS]
lin.append(linha("Subtotal: 8 eixos", df.no_eixo))
lin.append(linha('Regime "Pessoas comuns"', df.pc))
lin.append(linha("Conteúdo Sensível (qualquer)", df.sens))
lin.append(linha("Conteúdo Sensível nos 8 eixos", df.sens & df.no_eixo))
lin.append(linha('"Pessoas comuns" nos 8 eixos', df.pc & df.no_eixo))
lin.append(linha("Delimitação (eixo OU Pessoas comuns OU Sensível)", df.delim))
t_eixo = pd.DataFrame(lin)
t_eixo["< 10 posts"] = t_eixo.apply(lambda r: ", ".join(p for p in PERFIS if 0 < r[f"{p} (N)"] < 10), axis=1)

# Conteúdo Sensível por categoria, dentro dos 8 eixos e no total do perfil
cats = sorted({t for ls in df.sens_l for t in ls})
t_sens = pd.DataFrame([{"Conteúdo Sensível": c + (" †" if c in EMERGENTES else ""),
                        **{f"{p} — total (N)": int(df[(df.Perfil == p)].sens_l.apply(lambda x: c in x).sum()) for p in PERFIS},
                        **{f"{p} — nos 8 eixos (N)": int(df[(df.Perfil == p) & df.no_eixo].sens_l.apply(lambda x: c in x).sum()) for p in PERFIS}}
                       for c in cats])

# Mediação predominante por eixo (com empates)
def predominante(x):
    vc = x.value_counts()
    top = vc[vc == vc.iloc[0]]
    return " = ".join(top.index) + (" (empate, " if len(top) > 1 else " (") + f"{vc.iloc[0]} de {len(x)})"

med_eixo = (df[df.no_eixo].groupby(["Perfil", "Eixo Temático"])["Mediação da Informação"].agg(predominante)
            .rename("Mediação predominante (N de N)").reset_index())

# ================================================ ETAPAS 1 e 2: leitura (base da análise)
# s23b_anot_<Perfil>.tsv -> um post por linha: nível (S = situação social é o assunto; P = secundária;
#   N = nenhuma), situação (a), sujeitos (b), forma: nomeado/anônimo, rosto, ação, grupo (c),
#   papel: quem sofre, pede, ajuda, responde (d), cena sensível (e) e código do padrão.
# s23b_padroes_<Perfil>.json -> padrões nomeados na leitura, com evidências literais e contraexemplo.
PERG = {"a": "a) Situação social e como é apresentada", "b": "b) Quem aparece como sujeito e em que posição",
        "c": "c) Nomeado ou anônimo; rosto ou ocultado; ação; sozinho ou em grupo",
        "d": "d) Quem pede, ajuda, sofre e responde", "e": "e) Cenas sensíveis e como são tratadas"}
NIVEL = {"S": "Situação social é o assunto", "P": "Presente de forma secundária", "N": "Nenhuma situação social"}
lidos = [p for p in PERFIS if os.path.exists(f"s23b_anot_{p.split()[0]}.tsv")]
anot, pad_lin, ev, niv, div, cruz = {}, [], [], [], [], []
for p in lidos:
    a_ = pd.read_csv(f"s23b_anot_{p.split()[0]}.tsv", sep="\t", dtype=str).fillna("-")
    a_["linha"] = a_["linha"].astype(int)
    rec_p = set(df.loc[df.recorte & (df.Perfil == p), "Linha"])
    assert set(a_["linha"]) == rec_p and len(a_) == len(rec_p), f"anotação incompleta: {p}"
    anot[p] = a_
    for k, nome in NIVEL.items():
        niv.append({"Perfil": p, "Presença do campo social": nome, "N de posts lidos": int((a_.nivel == k).sum())})
    niv.append({"Perfil": p, "Presença do campo social": "Com cena sensível (pergunta e)", "N de posts lidos": int((a_.sensivel != "-").sum())})
    P = json.load(open(f"s23b_padroes_{p.split()[0]}.json", encoding="utf8"))
    for q in P["padroes"]:
        n = int((a_.padrao.isin(q["tags"]) & a_.nivel.isin(q["niveis"])).sum())
        pad_lin.append({"Perfil": p, "Código": q["cod"], "Padrão": q["nome"], "Definição": q["def"], "N de posts lidos": n,
                        "Contraexemplo": f"linha {q['contra'][0]}: {q['contra'][2]}" if q["contra"] else ""})
        for l, pg, t in q["ev"] + ([q["contra"]] if q["contra"] else []):
            assert t in DESC[str(l)], (l, t)                          # trecho literal da descrição
            assert int(l) in rec_p, (l, "fora do recorte do perfil")
            ev.append({"Linha": l, "Perfil": p, "Padrão": f"{q['cod']} {q['nome']}" + (" (contraexemplo)" if q["contra"] and [l, pg, t] == q["contra"] else ""),
                       "Pergunta de leitura": PERG[pg], "Trecho": t})
    m = df[df.recorte & (df.Perfil == p)].merge(a_, left_on="Linha", right_on="linha")
    for k, nome in NIVEL.items():
        div.append({"Perfil": p, "Presença lida": nome,
                    "Nos 8 eixos": int(((m.nivel == k) & m.no_eixo).sum()),
                    "Na delimitação, fora dos eixos": int(((m.nivel == k) & m.delim & ~m.no_eixo).sum()),
                    "Fora da delimitação": int(((m.nivel == k) & ~m.delim).sum()),
                    'Com Regime "Pessoas comuns"': int(((m.nivel == k) & m.pc).sum())})
    ct = pd.crosstab(m.padrao, m["Mediação da Informação"]).reset_index().rename(columns={"padrao": "Código de anotação"})
    ct.insert(0, "Perfil", p); cruz.append(ct)
t_niv, t_pad, t_div = pd.DataFrame(niv), pd.DataFrame(pad_lin), pd.DataFrame(div)
t_ev = pd.DataFrame(ev, columns=["Linha", "Perfil", "Padrão", "Pergunta de leitura", "Trecho"])
if cruz:
    t_cruz = pd.concat(cruz).fillna(0)
    t_cruz = t_cruz.astype({c: int for c in t_cruz.columns if c not in ("Perfil", "Código de anotação")})
else:
    t_cruz = pd.DataFrame()
t_anot = pd.concat([anot[p].assign(Perfil=p) for p in lidos]) if lidos else pd.DataFrame()
t_leit = pd.DataFrame([{"Perfil": p, "Posts do recorte (todos lidos)": int((df.recorte & (df.Perfil == p)).sum()),
                        "Posts do recorte na delimitação (apoio)": int((df.recorte & df.delim & (df.Perfil == p)).sum()),
                        "Lido": "sim" if p in lidos else "pendente", "Peso de desenho": PESOS[p]} for p in PERFIS])

notas = pd.DataFrame([
    ["Versão", "2 (08/11/2026): a leitura de todas as descrições do recorte é a base (80%); as variáveis são apoio. A versão 1 está salva como v1_*."],
    ["Base de leitura", "Todas as descrições do recorte (Ego 298; Roma 551; Imperatriz Online 756), em lotes de até 60. Perfis lidos: " + (", ".join(lidos) or "nenhum") + "."],
    ["Perguntas de leitura", "a) situações sociais; b) sujeitos e posição; c) nomeado/anônimo, rosto/ocultado, ação, grupo; d) quem pede, ajuda, sofre e responde; e) cenas sensíveis."],
    ["Anotação", "Nível de presença do campo social (S assunto; P secundário; N nenhum), situação, sujeitos, forma, papel, cena sensível e código de padrão. Contagens = posts lidos (sem percentuais)."],
    ["Delimitação (apoio)", "Eixos " + ", ".join(EIXOS) + "; Regime \"Pessoas comuns\"; Conteúdo Sensível ≠ \"Não há\" (confirmada em 08/11/2026)."],
    ["Base quantitativa (Etapa 3)", "4.266 posts (cópia da planilha congelada de 04/11/2026, com as correções aprovadas). Denominador: posts do perfil."],
    ["Emergentes", "† Procedimento (categoria emergente, indicativa)."],
    ["Pesos de desenho", "Ego 1,39; Roma 2,05; Imperatriz Online 3,60 (só ao somar perfis a partir do recorte; esta solicitação não soma)."],
    ["Cuidado ético", "Cenas com vítimas, crianças, corpos e violência são citadas por trechos curtos, sem acrescentar detalhes."],
    ["Limites", "As imagens não foram vistas; a leitura usa as descrições da pesquisadora, que às vezes já trazem interpretação. Níveis e códigos atribuídos por IA (Claude, Anthropic), a conferir."],
    ["Evidências", f"{len(t_ev)} trechos, todos cópia literal da descrição e do recorte do perfil."],
], columns=["Item", "Descrição"])

with pd.ExcelWriter("Solicitacao23_Tabelas.xlsx", engine="openpyxl") as w:
    t_eixo.to_excel(w, sheet_name="Eixos_Regime_Sensivel", index=False)
    t_sens.to_excel(w, sheet_name="Conteudo_sensivel", index=False)
    med_eixo.to_excel(w, sheet_name="Mediacao_por_eixo", index=False)
    t_leit.to_excel(w, sheet_name="Base_de_leitura", index=False)
    t_niv.to_excel(w, sheet_name="Leitura_presenca", index=False)
    t_pad.to_excel(w, sheet_name="Leitura_padroes", index=False)
    t_div.to_excel(w, sheet_name="Leitura_x_codificacao", index=False)
    t_cruz.to_excel(w, sheet_name="Leitura_x_mediacao", index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)
with pd.ExcelWriter("Solicitacao23_Evidencias.xlsx", engine="openpyxl") as w:
    t_ev.to_excel(w, sheet_name="Evidencias", index=False)
    t_pad.to_excel(w, sheet_name="Padroes", index=False)
    t_anot.to_excel(w, sheet_name="Anotacao_por_post", index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
for arq in ["Solicitacao23_Tabelas.xlsx", "Solicitacao23_Evidencias.xlsx"]:
    wb = load_workbook(arq)
    for ws in wb:
        for c in ws[1]:
            c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9E2F3")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.iter_cols(min_row=1):
            h = str(col[0].value or "")
            ws.column_dimensions[col[0].column_letter].width = 60 if h in ("Descrição", "Trecho", "Padrão", "Definição", "Contraexemplo", "situacao") else max(12, min(32, len(h) + 2))
            if "%" in h:
                for c in col[1:]:
                    c.number_format = "0.0%"
        ws.freeze_panes = "A2"
    wb.save(arq)

json.dump({k: v.to_dict("records") for k, v in {"eixo": t_eixo, "sens": t_sens, "med_eixo": med_eixo, "leit": t_leit,
           "niv": t_niv, "pad": t_pad, "div": t_div, "cruz": t_cruz}.items()},
          open("s23_tabelas.json", "w", encoding="utf8"), ensure_ascii=False, default=str)

pd.set_option("display.width", 250)
for t in (t_eixo, t_leit, t_niv, t_pad, t_div, t_cruz):
    print(t.to_string(index=False))
