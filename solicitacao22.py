"""
Solicitação 22 — Como o campo cultural é retratado nos perfis.

Uso:  python3 -I solicitacao22.py copia.xlsx

Delimitação (confirmada pela pesquisadora em 08/11/2026):
- Eixos: Cultura, Entretenimento, Humor, Religiosidade, Premiação, Reconhecimento.
- Regime de Visibilidade: "Personalidades Públicas" e "Digital Influencer" (multirrótulo).
- Período Especial: só festas e datas (lista FESTAS abaixo); demais períodos ficam de fora.
- Linguagem Regional na delimitação (uniões aprovadas na S20).
- Mediação predominante dentro dos seis eixos.
- Versão 2: a delimitação só localiza posts; a leitura de TODAS as descrições do recorte é a base.
Denominadores: posts do perfil (Ego 414; Roma 1.131; Imperatriz Online 2.721).
A planilha original não é alterada (lê-se a cópia, com as correções aprovadas da I18).
"""
import os
import re
import sys
import json
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
import correcoes_planilha

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
EIXOS = ["Cultura", "Entretenimento", "Humor", "Religiosidade", "Premiação", "Reconhecimento"]
REGIMES = ["Personalidades Públicas", "Digital Influencer"]
FESTAS = ["Carnaval", "Natal", "Festejos juninos", "São João", "Expoimp", "FECOIMP", "Lava Pratos",
          "Aniversário de Imperatriz", "Semana Santa/Páscoa", "Semana Santa", "Corpus Christi",
          "Romaria de Nossa Senhora Aparecida", "Festejo", "Festejos", "Bloco do Imprensa", "Fim de ano",
          "Festas de fim de ano", "Dia das Mães", "Dia dos Pais", "Dia das Crianças", "Dia dos Namorados",
          "Dia da Mulher", "Mês do Orgulho LGBTQIAPN+", "Tiradentes", "Veraneio", "Época de praia"]
UNIOES_LR = {"Óia só": "Oia", "E ai galera": "E aí"}
PESOS = {"Ego": 1.39, "Roma": 2.05, "Imperatriz Online": 3.60}


def rotulos(c, mapa=None):
    """Multirrótulo em ";" (o ";" dentro de parênteses não separa), sem 'Não há' e sem repetição."""
    if not isinstance(c, str):
        return []
    c = re.sub(r"\(([^)]*)\)", lambda m: "(" + m.group(1).replace(";", "§") + ")", c)
    ts = [t.strip().replace("§", ";") for t in c.split(";") if t.strip() and t.strip() != "Não há"]
    if mapa:
        ts = [mapa.get(t, t) for t in ts]
    return list(dict.fromkeys(ts))


df = correcoes_planilha.aplicar(pd.read_excel(ARQ))
assert len(df) == 4266
df["Linha"] = df.index + 2
NP = df["Perfil"].value_counts().to_dict()
DESC = json.load(open("desc_recorte.json", encoding="utf8"))
df["recorte"] = df["Linha"].astype(str).isin(DESC)
df["reg"] = df["Regime de Visibilidade"].apply(rotulos)
df["per"] = df["Período Especial"].apply(rotulos)
df["festas"] = df["per"].apply(lambda x: [t for t in x if t in FESTAS])
df["lr"] = df["Linguagem Regional"].apply(lambda c: rotulos(c, UNIOES_LR))
df["no_eixo"] = df["Eixo Temático"].isin(EIXOS)
for r in REGIMES:
    df[r] = df["reg"].apply(lambda x, r=r: r in x)
df["regime"] = df[REGIMES].any(axis=1)
df["delim"] = df["no_eixo"] | df["regime"]

# ---------------------------------------------------------------- 1.1 eixos e Regimes
def linha(nome, mask):
    return {"Categoria": nome, **{c: v for p in PERFIS for c, v in
            ((f"{p} (N)", int((mask & (df.Perfil == p)).sum())), (f"{p} (%)", (mask & (df.Perfil == p)).sum() / NP[p]))}}

lin = [linha(e, df["Eixo Temático"] == e) for e in EIXOS]
lin.append(linha("Subtotal: 6 eixos", df.no_eixo))
lin += [linha(f"Regime \"{r}\"", df[r]) for r in REGIMES]
lin.append(linha("Algum dos dois Regimes", df.regime))
lin.append(linha("Eixo da lista OU Regime (delimitação)", df.delim))
lin.append(linha("Posts com festa ou data (Período Especial)", df["festas"].str.len() > 0))
t_eixo = pd.DataFrame(lin)
t_eixo["< 10 posts"] = t_eixo.apply(lambda r: ", ".join(p for p in PERFIS if 0 < r[f"{p} (N)"] < 10), axis=1)

# Regimes por eixo: em que eixos aparecem Personalidades e Influenciadores
reg_eixo = []
for p in PERFIS:
    for r in REGIMES:
        s = df[(df.Perfil == p) & df[r]]
        for e, n in s["Eixo Temático"].value_counts().head(5).items():
            reg_eixo.append({"Perfil": p, "Regime": r, "Eixo Temático": e, "N": int(n), "% dos posts com o Regime": n / len(s)})
t_reg_eixo = pd.DataFrame(reg_eixo)

# ---------------------------------------------------------------- 1.2 festas e datas
c_f = {p: Counter(t for ls in df.loc[df.Perfil == p, "festas"] for t in ls) for p in PERFIS}
ordem = sorted(set().union(*c_f.values()), key=lambda t: (-sum(c_f[p][t] for p in PERFIS), t.casefold()))
t_fest = pd.DataFrame([{"Festa ou data": t, **{c: v for p in PERFIS for c, v in
                        ((f"{p} (N)", c_f[p][t]), (f"{p} (%)", c_f[p][t] / NP[p]))}} for t in ordem])
fora = sorted({t for ls in df["per"] for t in ls if t not in FESTAS})

# ---------------------------------------------------------------- 1.3 Mediação dentro dos 6 eixos
med = []
for p in PERFIS:
    s = df[(df.Perfil == p) & df.no_eixo]
    for m, n in s["Mediação da Informação"].value_counts().items():
        med.append({"Perfil": p, "Mediação da Informação": m, "N": int(n), "% dos posts do perfil nos 6 eixos": n / len(s),
                    "< 10 posts": "*" if n < 10 else ""})
t_med = pd.DataFrame(med)
med_eixo = (df[df.no_eixo].groupby(["Perfil", "Eixo Temático"])["Mediação da Informação"]
            .agg(lambda x: " = ".join(x.value_counts()[x.value_counts() == x.value_counts().iloc[0]].index)
                 + (" (empate, " if (x.value_counts() == x.value_counts().iloc[0]).sum() > 1 else " (")
                 + f"{x.value_counts().iloc[0]} de {len(x)})").rename("Mediação predominante (N de N)").reset_index())

# ---------------------------------------------------------------- 1.4 Linguagem Regional na delimitação
t_lr = []
for p in PERFIS:
    s = df[(df.Perfil == p) & df.delim]
    c = Counter(t for ls in s["lr"] for t in ls)
    t_lr.append({"Perfil": p, "Posts na delimitação": len(s), "Com linguagem regional (N)": int((s["lr"].str.len() > 0).sum()),
                 "Com linguagem regional (% da delimitação)": (s["lr"].str.len() > 0).mean(),
                 "Termos (N de posts)": "; ".join(f"{t} ({n})" for t, n in sorted(c.items(), key=lambda x: (-x[1], x[0].casefold())))})
t_lr = pd.DataFrame(t_lr)

# ================================================================ ETAPAS 1 e 2: LEITURA (base da análise)
# Versão 2 (pedido de 08/11/2026): lê-se TODAS as descrições do recorte de cada perfil;
# as variáveis acima (Etapa 3) servem só de apoio. Arquivos da leitura, por perfil lido:
#   s22b_anot_<Perfil>.tsv  -> um post por linha: nível de presença do campo cultural
#        (S = fato cultural é o assunto; P = presente de forma secundária, p. ex. trilha de anúncio;
#         N = nenhum), código do padrão e marca local de linguagem, gesto, roupa ou som (pergunta e);
#   s22b_padroes_<Perfil>.json -> padrões nomeados na leitura, com evidências e contraexemplo.
# As descrições lidas são as mesmas da S21 (s21b_<perfil>.txt, lotes de até 60), relidas com as perguntas da S22.
PERG = {"a": "a) Manifestações, eventos, personagens e como são descritos", "b": "b) Quem é mostrado e em que posição",
        "c": "c) Cenário e enquadramento", "d": "d) Notícia, divulgação ou entretenimento", "e": "e) Marcas locais de linguagem, gesto, roupa ou som"}
NIVEL = {"S": "Fato cultural é o assunto", "P": "Presente de forma secundária", "N": "Nenhum fato cultural"}
lidos = [p for p in PERFIS if os.path.exists(f"s22b_anot_{p.split()[0]}.tsv")]
anot, pad_lin, ev, niv, div, cruz = {}, [], [], [], [], []
for p in lidos:
    a_ = pd.read_csv(f"s22b_anot_{p.split()[0]}.tsv", sep="\t", dtype=str).fillna("-")
    a_["linha"] = a_["linha"].astype(int)
    rec_p = set(df.loc[df.recorte & (df.Perfil == p), "Linha"])
    assert set(a_["linha"]) == rec_p and len(a_) == len(rec_p), f"anotação incompleta: {p}"
    anot[p] = a_
    for k, nome in NIVEL.items():
        niv.append({"Perfil": p, "Presença do campo cultural": nome, "N de posts lidos": int((a_.nivel == k).sum())})
    niv.append({"Perfil": p, "Presença do campo cultural": "Com marca local (pergunta e)", "N de posts lidos": int((a_.marca_local == "sim").sum())})
    P = json.load(open(f"s22b_padroes_{p.split()[0]}.json", encoding="utf8"))
    for q in P["padroes"]:
        n = int((a_.padrao.isin(q["tags"]) & a_.nivel.isin(q["niveis"])).sum())
        pad_lin.append({"Perfil": p, "Código": q["cod"], "Padrão": q["nome"], "Definição": q["def"], "N de posts lidos": n,
                        "Contraexemplo": f"linha {q['contra'][0]}: {q['contra'][2]}" if q["contra"] else ""})
        for l, pg, t in q["ev"] + ([q["contra"]] if q["contra"] else []):
            assert t in DESC[str(l)], (l, t)
            assert int(l) in rec_p, (l, "fora do recorte do perfil")
            ev.append({"Linha": l, "Perfil": p, "Padrão": f"{q['cod']} {q['nome']}" + (" (contraexemplo)" if q["contra"] and [l, pg, t] == q["contra"] else ""),
                       "Pergunta de leitura": PERG[pg], "Trecho": t})
    m = df[df.recorte & (df.Perfil == p)].merge(a_, left_on="Linha", right_on="linha")
    for k, nome in NIVEL.items():
        div.append({"Perfil": p, "Presença lida": nome, "Nos 6 eixos": int(((m.nivel == k) & m.no_eixo).sum()),
                    "Fora dos 6 eixos": int(((m.nivel == k) & ~m.no_eixo).sum())})
    ct = pd.crosstab(m.padrao, m["Mediação da Informação"]).reset_index().rename(columns={"padrao": "Código de anotação"})
    ct.insert(0, "Perfil", p); cruz.append(ct)
t_niv, t_pad, t_div = pd.DataFrame(niv), pd.DataFrame(pad_lin), pd.DataFrame(div)
t_ev = pd.DataFrame(ev, columns=["Linha", "Perfil", "Padrão", "Pergunta de leitura", "Trecho"])
t_cruz = pd.concat(cruz).fillna(0).astype({c: int for c in pd.concat(cruz).columns if c not in ("Perfil", "Código de anotação")}) if cruz else pd.DataFrame()
t_anot = pd.concat([anot[p].assign(Perfil=p) for p in lidos]) if lidos else pd.DataFrame()
t_leit = pd.DataFrame([{"Perfil": p, "Posts do recorte (todos lidos)": int((df.recorte & (df.Perfil == p)).sum()),
                        "Posts do recorte na delimitação (apoio)": int((df.recorte & df.delim & (df.Perfil == p)).sum()),
                        "Lido": "sim" if p in lidos else "pendente", "Peso de desenho": PESOS[p]} for p in PERFIS])

notas = pd.DataFrame([
    ["Versão", "2 (08/11/2026): a leitura de todas as descrições do recorte é a base (80%); as variáveis são apoio. A versão 1, que lia só a delimitação, está salva como v1_*."],
    ["Base de leitura", "Todas as descrições do recorte do perfil (Ego 298; Roma 551; Imperatriz Online 756), lidas por inteiro, em lotes de até 60. Perfis lidos: " + (", ".join(lidos) or "nenhum") + "."],
    ["Perguntas de leitura", "a) manifestações, eventos e personagens; b) quem é mostrado e em que posição; c) cenário e enquadramento; d) notícia, divulgação ou entretenimento; e) marcas locais de linguagem, gesto, roupa ou som."],
    ["Anotação", "Cada post lido recebeu um nível de presença do campo cultural (S assunto; P secundário; N nenhum), um código de padrão e a marca local (pergunta e). Contagens dos padrões = posts lidos (sem percentuais). Código OUTRO = fato cultural que não forma padrão."],
    ["Base quantitativa (Etapa 3, apoio)", "4.266 posts (cópia da planilha congelada de 04/11/2026, com as correções aprovadas). Denominador: posts do perfil (Ego 414; Roma 1.131; Imperatriz Online 2.721)."],
    ["Delimitação para localizar posts", "Eixos Cultura, Entretenimento, Humor, Religiosidade, Premiação, Reconhecimento; Regimes \"Personalidades Públicas\" e \"Digital Influencer\"; Período Especial (festas e datas); Linguagem Regional (confirmada em 08/11/2026)."],
    ["Festas e datas", "; ".join(FESTAS) + "."],
    ["Multirrótulo", "Regime, Período Especial e Linguagem Regional divididos em \";\" (exceto dentro de parênteses); conta-se posts."],
    ["Pesos de desenho", "Ego 1,39; Roma 2,05; Imperatriz Online 3,60 (só ao somar perfis a partir do recorte; esta solicitação não soma)."],
    ["Limites", "As imagens não foram vistas; a leitura usa as descrições escritas pela pesquisadora, que às vezes já trazem interpretação. Os níveis e códigos foram atribuídos por IA (Claude, Anthropic) e devem ser conferidos."],
    ["Evidências", f"{len(t_ev)} trechos, todos cópia literal da descrição e do recorte do perfil."],
], columns=["Item", "Descrição"])

with pd.ExcelWriter("Solicitacao22_Tabelas.xlsx", engine="openpyxl") as w:
    t_eixo.to_excel(w, "Eixos_e_Regimes", index=False)
    t_reg_eixo.to_excel(w, "Regimes_por_eixo", index=False)
    t_fest.to_excel(w, "Festas_e_datas", index=False)
    t_med.to_excel(w, "Mediacao_nos_eixos", index=False)
    med_eixo.to_excel(w, "Mediacao_por_eixo", index=False)
    t_lr.to_excel(w, "Linguagem_regional", index=False)
    t_leit.to_excel(w, "Base_de_leitura", index=False)
    t_niv.to_excel(w, "Leitura_presenca", index=False)
    t_pad.to_excel(w, "Leitura_padroes", index=False)
    t_div.to_excel(w, "Leitura_x_eixos", index=False)
    t_cruz.to_excel(w, "Leitura_x_mediacao", index=False)
    notas.to_excel(w, "Notas", index=False)
with pd.ExcelWriter("Solicitacao22_Evidencias.xlsx", engine="openpyxl") as w:
    t_ev.to_excel(w, "Evidencias", index=False)
    t_pad.to_excel(w, "Padroes", index=False)
    t_anot.to_excel(w, "Anotacao_por_post", index=False)
    notas.to_excel(w, "Notas", index=False)

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
for arq in ["Solicitacao22_Tabelas.xlsx", "Solicitacao22_Evidencias.xlsx"]:
    wb = load_workbook(arq)
    for ws in wb:
        for c in ws[1]:
            c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9E2F3")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.iter_cols(min_row=1):
            h = str(col[0].value or "")
            ws.column_dimensions[col[0].column_letter].width = 60 if h in ("Descrição", "Trecho", "Termos (N de posts)", "Padrão", "Definição", "Contraexemplo") else max(12, min(32, len(h) + 2))
            if "%" in h:
                for c in col[1:]:
                    c.number_format = "0.0%"
        ws.freeze_panes = "A2"
    wb.save(arq)

json.dump({"eixo": t_eixo.to_dict("records"), "reg_eixo": t_reg_eixo.to_dict("records"), "fest": t_fest.to_dict("records"),
           "med": t_med.to_dict("records"), "med_eixo": med_eixo.to_dict("records"), "lr": t_lr.to_dict("records"),
           "leit": t_leit.to_dict("records"), "niv": t_niv.to_dict("records"), "pad": t_pad.to_dict("records"),
           "div": t_div.to_dict("records"), "cruz": t_cruz.to_dict("records")}, open("s22_tabelas.json", "w", encoding="utf8"), ensure_ascii=False, default=str)

pd.set_option("display.width", 250)
for t in (t_eixo, t_med, t_leit, t_niv, t_pad, t_div, t_cruz):
    print(t.to_string(index=False))
