"""
Solicitação 24 — Que usuário os posts pressupõem.

Uso:  python3 -I solicitacao24.py copia.xlsx

Base da análise (80%): leitura de TODAS as descrições do recorte de cada perfil
(Ego 298; Roma 551; Imperatriz Online 756), em lotes de até 60 (arquivos s24b_<perfil>.txt).
Cada post foi anotado em s24b_anot_<Perfil>.tsv com as cinco perguntas de leitura:
  a) posicao  — onde o usuário é colocado (espectador, testemunha, consumidor, morador, fã, participante...);
  b) dirige   — como a imagem/legenda se dirige a ele ("-" = não se dirige);
  c) supoe    — o que precisa saber para entender ("-" = o post identifica tudo);
  d) reacao   — reação que o post prepara (INFERÊNCIA, só quando a descrição sustenta);
  e) exige    — o que o post exige de quem olha;
  padrao      — código do padrão nomeado na leitura.
Os padrões (nome, definição, evidências literais, contraexemplo) estão em s24b_padroes_<Perfil>.json.

Etapa 3 (apoio): CTA (Algum/Não há e cinco formas mais frequentes, multirrótulo), Mediação,
Linguagem Regional (posts com termo regional) e Collab, por perfil, e o cruzamento padrão × variáveis.
Denominadores: posts do perfil (Ego 414; Roma 1.131; Imperatriz Online 2.721) e, no cruzamento, posts lidos.
A planilha original não é alterada (lê-se a cópia, com as correções aprovadas).
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
df["cta_l"] = df["CTA"].apply(rotulos)
df["cta"] = df["cta_l"].str.len() > 0
df["lr"] = df["Linguagem Regional"].apply(rotulos).str.len() > 0
df["collab"] = df["Collab"].astype(str).str.strip().eq("Sim")

# ------------------------------------------------ arquivos de leitura (lotes de até 60, todo o recorte)
for p in PERFIS:
    s = df[df.recorte & (df.Perfil == p)]
    with open(f"s24b_{p.split()[0]}.txt", "w", encoding="utf8") as f:
        for i, (_, r) in enumerate(s.iterrows()):
            if i % 60 == 0:
                f.write(f"===== LOTE {i // 60 + 1}\n")
            f.write(f"[{r.Linha}|CTA:{r.CTA}|Collab:{r.Collab}|LR:{r['Linguagem Regional']}] {DESC[str(r.Linha)]}\n")


# ================================================ ETAPA 3: apoio das variáveis (uma tabela por perfil)
def bloco(p):
    s = df[df.Perfil == p]
    n = NP[p]
    out = [("CTA", "Algum CTA", int(s.cta.sum())), ("CTA", "Não há", int((~s.cta).sum()))]
    formas = pd.Series([t for ls in s.cta_l for t in ls]).value_counts()
    out += [("CTA — cinco formas mais frequentes (multirrótulo)", f, int(v)) for f, v in formas.head(5).items()]
    out += [("Mediação da Informação", m, int(v)) for m, v in s["Mediação da Informação"].value_counts().items()]
    out += [("Linguagem Regional", "Posts com termo regional", int(s.lr.sum())),
            ("Linguagem Regional", "Não há", int((~s.lr).sum())),
            ("Collab", "Sim", int(s.collab.sum())), ("Collab", "Não", int((~s.collab).sum()))]
    t = pd.DataFrame(out, columns=["Variável", "Categoria", "N"])
    t["%"] = t.N / n
    t["Denominador"] = n
    t["Marca"] = t.N.apply(lambda v: "*" if 0 < v < 10 else "")
    return t


t_var = {p: bloco(p) for p in PERFIS}

# ================================================ ETAPAS 1 e 2: leitura (base da análise)
PERG = {"a": "a) Onde o usuário é colocado", "b": "b) A imagem se dirige ao usuário",
        "c": "c) Que conhecimento o post supõe", "d": "d) Que reação o post prepara (inferência)",
        "e": "e) O que o post exige de quem olha"}
lidos = [p for p in PERFIS if os.path.exists(f"s24b_anot_{p.split()[0]}.tsv")]
anot, pad_lin, ev, resp, cruz = {}, [], [], [], []
for p in lidos:
    a_ = pd.read_csv(f"s24b_anot_{p.split()[0]}.tsv", sep="\t", dtype=str).fillna("-")
    a_["linha"] = a_["linha"].astype(int)
    rec_p = set(df.loc[df.recorte & (df.Perfil == p), "Linha"])
    assert set(a_["linha"]) == rec_p and len(a_) == len(rec_p), f"anotação incompleta: {p}"
    anot[p] = a_
    # quantos posts responderam a cada pergunta (b–e); a) é sempre preenchida
    for col, pg in (("dirige", "b"), ("supoe", "c"), ("reacao", "d"), ("exige", "e")):
        resp.append({"Perfil": p, "Pergunta": PERG[pg], "Posts com registro": int((a_[col] != "-").sum()), "Posts lidos": len(a_)})
    for r_, v in a_.reacao[a_.reacao != "-"].value_counts().items():
        resp.append({"Perfil": p, "Pergunta": f"   d) {r_}", "Posts com registro": int(v), "Posts lidos": len(a_)})
    for r_, v in a_.exige[a_.exige != "-"].value_counts().items():
        resp.append({"Perfil": p, "Pergunta": f"   e) {r_}", "Posts com registro": int(v), "Posts lidos": len(a_)})
    P = json.load(open(f"s24b_padroes_{p.split()[0]}.json", encoding="utf8"))
    cods = {}
    for q in P["padroes"]:
        n = int(a_.padrao.isin(q["tags"]).sum())
        for t_ in q["tags"]:
            cods[t_] = f"{q['cod']} {q['nome']}"
        pad_lin.append({"Perfil": p, "Código": q["cod"], "Padrão": q["nome"], "Definição": q["def"],
                        "N de posts lidos": n, "Marca": "*" if n < 10 else "",
                        "Contraexemplo": f"linha {q['contra'][0]}: {q['contra'][2]}" if q["contra"] else ""})
        for l, pg, t in q["ev"] + ([q["contra"]] if q["contra"] else []):
            assert t in DESC[str(l)], (l, t)                          # trecho literal da descrição
            assert int(l) in rec_p, (l, "fora do recorte do perfil")
            assert a_.loc[a_.linha == int(l), "padrao"].iloc[0] in q["tags"], (l, "evidência fora do padrão")
            ev.append({"Linha": l, "Perfil": p,
                       "Padrão": f"{q['cod']} {q['nome']}" + (" (contraexemplo)" if q["contra"] and [l, pg, t] == q["contra"] else ""),
                       "Pergunta de leitura": PERG[pg], "Trecho": t})
    assert set(a_.padrao) <= set(cods), set(a_.padrao) - set(cods)
    # cruzamento padrão × variáveis (posts lidos)
    m = df[df.recorte & (df.Perfil == p)].merge(a_, left_on="Linha", right_on="linha")
    m["Padrão"] = m.padrao.map(cods)
    g = m.groupby("Padrão")
    ct = pd.DataFrame({"N lidos": g.size(), "Algum CTA": g.cta.sum(),
                       **{f"CTA {f}": g.cta_l.apply(lambda s, f=f: s.apply(lambda x: f in x).sum())
                          for f in ["Interação", "Contato", "Ação/Visita", "Redirecionamento"]},
                       "Collab Sim": g.collab.sum(), "Com termo regional": g.lr.sum(),
                       "Mediação predominante": g["Mediação da Informação"].agg(lambda x: f"{x.value_counts().index[0]} ({x.value_counts().iloc[0]})")}).reset_index()
    ordem = {f"{q['cod']} {q['nome']}": i for i, q in enumerate(P["padroes"])}
    ct = ct.sort_values("Padrão", key=lambda s: s.map(ordem))       # ordem dos padrões (E1 ... E10X)
    ct.insert(0, "Perfil", p)
    cruz.append(ct)

t_pad = pd.DataFrame(pad_lin)
t_ev = pd.DataFrame(ev, columns=["Linha", "Perfil", "Padrão", "Pergunta de leitura", "Trecho"])
t_resp = pd.DataFrame(resp)
t_cruz = pd.concat(cruz) if cruz else pd.DataFrame()
t_anot = pd.concat([anot[p].assign(Perfil=p) for p in lidos]) if lidos else pd.DataFrame()
t_leit = pd.DataFrame([{"Perfil": p, "Posts do recorte (todos lidos)": int((df.recorte & (df.Perfil == p)).sum()),
                        "Lido": "sim" if p in lidos else "pendente", "Peso de desenho": PESOS[p]} for p in PERFIS])

notas = pd.DataFrame([
    ["Pergunta", "Que usuário os posts pressupõem: posição oferecida a quem olha, a partir do que a imagem mostra, como mostra e o que pressupõe."],
    ["Base de leitura", "Todas as descrições do recorte (Ego 298; Roma 551; Imperatriz Online 756), em lotes de até 60. Perfis lidos: " + (", ".join(lidos) or "nenhum") + "."],
    ["Perguntas de leitura", "a) onde o usuário é colocado; b) se a imagem se dirige a ele; c) que conhecimento supõe; d) que reação prepara (inferência); e) o que exige."],
    ["Anotação", "Um post por linha, com as cinco respostas e um código de padrão. Contagens = posts lidos; sem percentuais na leitura. * = menos de 10 posts."],
    ["Etapa 3 (apoio)", "CTA (multirrótulo; \"Algum CTA\" = ao menos uma forma), Mediação, Linguagem Regional (posts com termo diferente de \"Não há\") e Collab. Denominador: posts do perfil."],
    ["Base quantitativa", "4.266 posts (cópia da planilha congelada de 04/11/2026, com as correções aprovadas)."],
    ["Pesos de desenho", "Ego 1,39; Roma 2,05; Imperatriz Online 3,60 (só ao somar perfis; esta solicitação não soma)."],
    ["Limites", "As imagens não foram vistas; a leitura usa as descrições da pesquisadora, que às vezes já trazem interpretação. Reações (d) são inferências. No Roma, a leitura inicial (420 a 650 caracteres, de média 752) foi refeita na íntegra em 09/11/2026, com 50 anotações corrigidas (5 mudanças de padrão; s24c_rev_Roma.tsv); no Imperatriz Online, a leitura inicial (260 caracteres nos anúncios, 420 nos demais, de média 1.264) foi refeita na íntegra em 09/11/2026: 76 anotações corrigidas, 14 com mudança de padrão (s24c_rev_Imperatriz.tsv; versão anterior em s24b_anot_Imperatriz_v1_leitura_truncada.tsv); as evidências foram conferidas no texto integral. Códigos atribuídos por IA (Claude, Anthropic), a conferir."],
    ["Evidências", f"{len(t_ev)} trechos, todos cópia literal da descrição e do recorte do perfil."],
], columns=["Item", "Descrição"])

with pd.ExcelWriter("Solicitacao24_Tabelas.xlsx", engine="openpyxl") as w:
    for p in PERFIS:
        t_var[p].to_excel(w, sheet_name=f"Variaveis_{p.split()[0]}", index=False)
    t_leit.to_excel(w, sheet_name="Base_de_leitura", index=False)
    t_resp.to_excel(w, sheet_name="Leitura_perguntas", index=False)
    t_pad.to_excel(w, sheet_name="Leitura_padroes", index=False)
    t_cruz.to_excel(w, sheet_name="Padrao_x_variaveis", index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)
with pd.ExcelWriter("Solicitacao24_Evidencias.xlsx", engine="openpyxl") as w:
    t_ev.to_excel(w, sheet_name="Evidencias", index=False)
    t_pad.to_excel(w, sheet_name="Padroes", index=False)
    t_anot.to_excel(w, sheet_name="Anotacao_por_post", index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
for arq in ["Solicitacao24_Tabelas.xlsx", "Solicitacao24_Evidencias.xlsx"]:
    wb = load_workbook(arq)
    for ws in wb:
        for c in ws[1]:
            c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9E2F3")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.iter_cols(min_row=1):
            h = str(col[0].value or "")
            ws.column_dimensions[col[0].column_letter].width = 60 if h in ("Descrição", "Trecho", "Padrão", "Definição", "Contraexemplo") else max(12, min(32, len(h) + 2))
            if h == "%":
                for c in col[1:]:
                    c.number_format = "0.0%"
        ws.freeze_panes = "A2"
    wb.save(arq)

json.dump({"var": {p: t.to_dict("records") for p, t in t_var.items()}, "resp": t_resp.to_dict("records"),
           "pad": t_pad.to_dict("records"), "cruz": t_cruz.to_dict("records"), "leit": t_leit.to_dict("records")},
          open("s24_tabelas.json", "w", encoding="utf8"), ensure_ascii=False, default=str)

pd.set_option("display.width", 250)
pd.set_option("display.max_colwidth", 40)
for p in PERFIS:
    print(t_var[p].to_string(index=False))
for t in (t_leit, t_resp, t_pad.drop(columns=["Definição"]), t_cruz):
    print(t.to_string(index=False))
