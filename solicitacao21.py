"""
Solicitação 21 — Como a cidade de Imperatriz é vista através dos perfis.

Uso:  python3 -I solicitacao21.py copia.xlsx

Delimitação (confirmada pela pesquisadora em 08/11/2026):
- Eixos: Cidade, Infraestrutura, Trânsito, Meio Ambiente, Alerta, Serviço.
- Regime de Visibilidade: "Espaço da cidade" (multirrótulo, separado por ";").
- Referência Local (multirrótulo): top 15 sem "Imperatriz" e "Maranhão" (e as variantes
  "Imperatriz-MA" e "Maranhã"); lugares fora de Imperatriz permanecem.
  Proporção de "Não identificado" + medida complementar: posts que citam só Imperatriz/Maranhão.
- Período Especial (multirrótulo) por perfil; Linguagem Regional dentro da delimitação
  (uniões aprovadas na S20: "Óia só" → "Oia"; "E ai galera" → "E aí").
- Delimitação usada só para localizar posts e conferir a leitura.
Denominadores: posts do perfil (Ego 414; Roma 1.131; Imperatriz Online 2.721).
A planilha original não é alterada (lê-se a cópia, com as correções aprovadas da I18).

Versão 2: a leitura de TODAS as descrições do recorte é a base (Etapas 1 e 2); as tabelas
abaixo são o apoio (Etapa 3). Saídas: Solicitacao21_Tabelas.xlsx, Solicitacao21_Evidencias.xlsx,
s21_tabelas.json e s21b_<perfil>.txt (descrições em lotes de até 60).
"""
import os
import sys
import json
import re
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
import correcoes_planilha

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
EIXOS = ["Cidade", "Infraestrutura", "Trânsito", "Meio Ambiente", "Alerta", "Serviço"]
GERAIS = {"Imperatriz", "Maranhão", "Imperatriz-MA", "Maranhã"}
UNIOES_LR = {"Óia só": "Oia", "E ai galera": "E aí"}
PESOS = {"Ego": 1.39, "Roma": 2.05, "Imperatriz Online": 3.60}   # pesos de desenho do recorte


def rotulos(c, mapa=None):
    """Divide célula multirrótulo em ";", remove vazios e 'Não há', sem repetição."""
    if not isinstance(c, str):
        return []
    # ";" dentro de parênteses não separa rótulos (ex.: "Campanhas do calendário de saúde
    # (Setembro Amarelo; Outubro Rosa etc.)" é um único rótulo de Período Especial)
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
df["local"] = df["Referência Local"].apply(rotulos)
df["periodo"] = df["Período Especial"].apply(rotulos)
df["lr"] = df["Linguagem Regional"].apply(lambda c: rotulos(c, UNIOES_LR))
df["no_eixo"] = df["Eixo Temático"].isin(EIXOS)
df["espaco"] = df["reg"].apply(lambda r: "Espaço da cidade" in r)
df["delim"] = df["no_eixo"] | df["espaco"]


def pct(n, p):
    return n / NP[p]


# ---------------------------------------------------------------- 1.1 eixos e Regime
lin = []
for e in EIXOS:
    lin.append({"Categoria": e, **{f"{p} (N)": int(((df.Perfil == p) & (df["Eixo Temático"] == e)).sum()) for p in PERFIS}})
lin.append({"Categoria": "Subtotal: 6 eixos", **{f"{p} (N)": int(((df.Perfil == p) & df.no_eixo).sum()) for p in PERFIS}})
lin.append({"Categoria": "Regime \"Espaço da cidade\"", **{f"{p} (N)": int(((df.Perfil == p) & df.espaco).sum()) for p in PERFIS}})
lin.append({"Categoria": "Eixo da lista OU Espaço da cidade", **{f"{p} (N)": int(((df.Perfil == p) & df.delim).sum()) for p in PERFIS}})
lin.append({"Categoria": "Eixo da lista E Espaço da cidade", **{f"{p} (N)": int(((df.Perfil == p) & df.no_eixo & df.espaco).sum()) for p in PERFIS}})
t_eixo = pd.DataFrame(lin)
for p in PERFIS:
    t_eixo[f"{p} (%)"] = t_eixo[f"{p} (N)"] / NP[p]
t_eixo = t_eixo[["Categoria"] + [c for p in PERFIS for c in (f"{p} (N)", f"{p} (%)")]]
t_eixo["< 10 posts"] = ""   # sinalização preenchida abaixo
t_eixo["< 10 posts"] = t_eixo.apply(lambda r: ", ".join(p for p in PERFIS if 0 < r[f"{p} (N)"] < 10), axis=1)

# Espaço da cidade por eixo (onde o espaço aparece)
esp_eixo = (df[df.espaco].groupby(["Perfil", "Eixo Temático"]).size().rename("N").reset_index())
esp_eixo["% dos posts com Espaço da cidade no perfil"] = esp_eixo.apply(
    lambda r: r["N"] / df[(df.Perfil == r["Perfil"]) & df.espaco].shape[0], axis=1)
esp_eixo = esp_eixo.sort_values(["Perfil", "N"], ascending=[True, False])
esp_eixo = pd.concat([esp_eixo[esp_eixo.Perfil == p].head(8) for p in PERFIS])

# ---------------------------------------------------------------- 1.2 locais
top = []
for p in PERFIS:
    c = Counter(t for ls in df.loc[df.Perfil == p, "local"] for t in ls if t not in GERAIS and t != "Não identificado")
    for i, (t, n) in enumerate(sorted(c.items(), key=lambda x: (-x[1], x[0].casefold()))[:15], 1):
        top.append({"Perfil": p, "Posição": i, "Local": t, "N de posts": n, "% dos posts do perfil": n / NP[p],
                    "< 10 posts": "*" if n < 10 else ""})
t_top = pd.DataFrame(top)
semloc = []
for p in PERFIS:
    s = df[df.Perfil == p]
    ni = s["local"].apply(lambda x: "Não identificado" in x).sum()
    ger = s["local"].apply(lambda x: len(x) > 0 and set(x) <= GERAIS).sum()
    esp = s["local"].apply(lambda x: any(t not in GERAIS and t != "Não identificado" for t in x)).sum()
    semloc.append({"Perfil": p, "Posts do perfil": NP[p],
                   "\"Não identificado\" (N)": int(ni), "\"Não identificado\" (%)": ni / NP[p],
                   "Só Imperatriz/Maranhão (N)": int(ger), "Só Imperatriz/Maranhão (%)": ger / NP[p],
                   "Com local específico (N)": int(esp), "Com local específico (%)": esp / NP[p]})
t_semloc = pd.DataFrame(semloc)

# ---------------------------------------------------------------- 1.3 Período Especial
c_per = {p: Counter(t for ls in df.loc[df.Perfil == p, "periodo"] for t in ls) for p in PERFIS}
todos = sorted(set().union(*c_per.values()), key=lambda t: (-sum(c_per[p][t] for p in PERFIS), t.casefold()))
t_per = pd.DataFrame([{"Período Especial": t, **{c: v for p in PERFIS for c, v in
                       ((f"{p} (N)", c_per[p][t]), (f"{p} (%)", c_per[p][t] / NP[p]))}} for t in todos])
t_per = pd.concat([t_per, pd.DataFrame([{"Período Especial": "Posts com algum período especial",
        **{c: v for p in PERFIS for c, v in ((f"{p} (N)", int((df.loc[df.Perfil == p, "periodo"].str.len() > 0).sum())),
                                             (f"{p} (%)", (df.loc[df.Perfil == p, "periodo"].str.len() > 0).mean()))}}])])

# ---------------------------------------------------------------- 1.4 Linguagem Regional na delimitação
lr = []
for p in PERFIS:
    s = df[(df.Perfil == p) & df.delim]
    c = Counter(t for ls in s["lr"] for t in ls)
    lr.append({"Perfil": p, "Posts na delimitação": len(s), "Com linguagem regional (N)": int((s["lr"].str.len() > 0).sum()),
               "Com linguagem regional (% da delimitação)": (s["lr"].str.len() > 0).mean() if len(s) else None,
               "Termos (N de posts)": "; ".join(f"{t} ({n})" for t, n in sorted(c.items(), key=lambda x: (-x[1], x[0].casefold())))})
t_lr = pd.DataFrame(lr)

# ================================================================ ETAPAS 1 e 2: LEITURA (base da análise)
# Versão 2 (pedido de 08/11/2026): lê-se TODAS as descrições do recorte de cada perfil,
# não só as da delimitação. As variáveis acima (Etapa 3) servem apenas de apoio.
# Arquivos da leitura, um por perfil já lido:
#   s21b_anot_<Perfil>.tsv  -> uma linha por post: nível de presença da cidade
#        (A = lugar é o assunto da imagem; F = lugar aparece ao fundo; C = lugar só citado;
#         N = nenhum lugar da cidade), lugar, situação, ocupantes, pressuposição, código do padrão;
#   s21b_padroes_<Perfil>.json -> padrões nomeados na leitura: definição, códigos de anotação,
#        níveis contados, evidências (linha, pergunta de leitura a–e, trecho) e contraexemplo.
for p in PERFIS:
    s = df[df.recorte & (df.Perfil == p)]
    with open(f"s21b_{p.split()[0]}.txt", "w", encoding="utf8") as f:   # lotes de até 60 posts
        for i, (_, r) in enumerate(s.iterrows()):
            if i % 60 == 0:
                f.write(f"===== LOTE {i // 60 + 1}\n")
            f.write(f"[{r.Linha}|{r['Eixo Temático']}|{r['Referência Local']}] {DESC[str(r.Linha)]}\n")

PERG = {"a": "a) Que lugares aparecem e como são nomeados", "b": "b) Lugar como assunto ou fundo; em cena ou citado",
        "c": "c) Estado ou situação da cidade", "d": "d) Quem ocupa o espaço", "e": "e) Pressupõe que o usuário conhece o lugar"}
NIVEL = {"A": "Lugar é o assunto da imagem", "F": "Lugar aparece como fundo", "C": "Lugar só citado", "N": "Nenhum lugar da cidade"}
lidos = [p for p in PERFIS if os.path.exists(f"s21b_anot_{p.split()[0]}.tsv")]
anot, pad_lin, ev, niv, div = {}, [], [], [], []
for p in lidos:
    a_ = pd.read_csv(f"s21b_anot_{p.split()[0]}.tsv", sep="\t", dtype=str).fillna("-")
    a_["linha"] = a_["linha"].astype(int)
    rec_p = set(df.loc[df.recorte & (df.Perfil == p), "Linha"])
    assert set(a_["linha"]) == rec_p and len(a_) == len(rec_p), f"anotação incompleta: {p}"   # todas lidas
    anot[p] = a_
    for k, nome in NIVEL.items():
        niv.append({"Perfil": p, "Presença da cidade na imagem": nome, "N de posts lidos": int((a_.nivel == k).sum())})
    niv.append({"Perfil": p, "Presença da cidade na imagem": "Pressupõe conhecimento local (pergunta e)",
                "N de posts lidos": int(a_.pressupoe.str.startswith("sim").sum())})
    P = json.load(open(f"s21b_padroes_{p.split()[0]}.json", encoding="utf8"))
    for q in P["padroes"]:
        n = int((a_.padrao.isin(q["tags"]) & a_.nivel.isin(q["niveis"])).sum())   # contagem da leitura
        pad_lin.append({"Perfil": p, "Código": q["cod"], "Padrão": q["nome"], "Definição": q["def"], "N de posts lidos": n,
                        "Contraexemplo": f"linha {q['contra'][0]}: {q['contra'][2]}" if q["contra"] else ""})
        for l, pg, t in q["ev"] + ([q["contra"]] if q["contra"] else []):
            assert t in DESC[str(l)], (l, t)                       # trecho literal
            assert int(l) in rec_p, (l, "fora do recorte do perfil")
            ev.append({"Linha": l, "Perfil": p, "Padrão": f"{q['cod']} {q['nome']}" + (" (contraexemplo)" if q["contra"] and l == q["contra"][0] and t == q["contra"][2] else ""),
                       "Pergunta de leitura": PERG[pg], "Trecho": t})
    # divergências leitura × codificação: Regime "Espaço da cidade" × nível de presença lido
    m = df[df.recorte & (df.Perfil == p)].merge(a_, left_on="Linha", right_on="linha")
    for k, nome in NIVEL.items():
        div.append({"Perfil": p, "Presença lida": nome, "Com Regime Espaço da cidade": int(((m.nivel == k) & m.espaco).sum()),
                    "Sem Regime Espaço da cidade": int(((m.nivel == k) & ~m.espaco).sum())})
t_niv, t_pad, t_ev, t_div = pd.DataFrame(niv), pd.DataFrame(pad_lin), pd.DataFrame(ev, columns=["Linha", "Perfil", "Padrão", "Pergunta de leitura", "Trecho"]), pd.DataFrame(div)
t_anot = pd.concat([anot[p].assign(Perfil=p) for p in lidos]) if lidos else pd.DataFrame()
t_leit = pd.DataFrame([{"Perfil": p, "Posts do recorte (todos lidos)": int((df.recorte & (df.Perfil == p)).sum()),
                        "Posts do recorte na delimitação (apoio)": int((df.recorte & df.delim & (df.Perfil == p)).sum()),
                        "Lido": "sim" if p in lidos else "pendente", "Peso de desenho": PESOS[p]} for p in PERFIS])

notas = pd.DataFrame([
    ["Versão", "2 (08/11/2026): a leitura de todas as descrições do recorte é a base (80%); as variáveis são apoio. A versão 1, que lia só a delimitação, está salva como v1_*."],
    ["Base de leitura", "Todas as descrições do recorte do perfil (Ego 298; Roma 551; Imperatriz Online 756), lidas por inteiro, em lotes de até 60. Perfis lidos: " + (", ".join(lidos) or "nenhum") + "."],
    ["Perguntas de leitura", "a) lugares e como são nomeados; b) assunto ou fundo, em cena ou citado; c) estado ou situação; d) quem ocupa o espaço; e) pressupõe conhecimento local."],
    ["Anotação", "Cada post lido recebeu um nível de presença da cidade (A assunto; F fundo; C só citado; N nenhum) e, quando cabia, um código de padrão. Contagens dos padrões = posts lidos com o código e o nível indicados (sem percentuais)."],
    ["Base quantitativa (Etapa 3, apoio)", "4.266 posts (cópia da planilha congelada de 04/11/2026, com as correções aprovadas). Denominador: posts do perfil (Ego 414; Roma 1.131; Imperatriz Online 2.721)."],
    ["Delimitação para localizar posts", "Eixos Cidade, Infraestrutura, Trânsito, Meio Ambiente, Alerta, Serviço; Regime \"Espaço da cidade\"; Referência Local; Período Especial; Linguagem Regional (confirmada em 08/11/2026)."],
    ["Multirrótulo", "Regime, Referência Local, Período Especial e Linguagem Regional divididos em \";\" (exceto \";\" dentro de parênteses); conta-se posts."],
    ["Locais excluídos do top 15", "\"Imperatriz\", \"Maranhão\", \"Imperatriz-MA\" e \"Maranhã\". \"Não identificado\" medido à parte. Lugares fora de Imperatriz permanecem."],
    ["Pesos de desenho", "Ego 1,39; Roma 2,05; Imperatriz Online 3,60 (só ao somar perfis a partir do recorte; esta solicitação não soma)."],
    ["Limites", "As imagens não foram vistas; a leitura usa as descrições escritas pela pesquisadora, que às vezes já trazem interpretação. A anotação dos níveis e padrões foi feita por IA (Claude, Anthropic) e deve ser conferida."],
    ["Evidências", f"{len(t_ev)} trechos, todos cópia literal da descrição e do recorte do perfil."],
], columns=["Item", "Descrição"])

with pd.ExcelWriter("Solicitacao21_Tabelas.xlsx", engine="openpyxl") as w:
    t_eixo.to_excel(w, "Eixos_e_Espaco", index=False)
    esp_eixo.to_excel(w, "Espaco_por_eixo", index=False)
    t_top.to_excel(w, "Locais_top15", index=False)
    t_semloc.to_excel(w, "Posts_sem_local", index=False)
    t_per.to_excel(w, "Periodo_especial", index=False)
    t_lr.to_excel(w, "Linguagem_regional", index=False)
    t_leit.to_excel(w, "Base_de_leitura", index=False)
    t_niv.to_excel(w, "Leitura_presenca", index=False)
    t_pad.to_excel(w, "Leitura_padroes", index=False)
    t_div.to_excel(w, "Leitura_x_codificacao", index=False)
    notas.to_excel(w, "Notas", index=False)
with pd.ExcelWriter("Solicitacao21_Evidencias.xlsx", engine="openpyxl") as w:
    t_ev.to_excel(w, "Evidencias", index=False)
    t_pad.to_excel(w, "Padroes", index=False)
    t_anot.to_excel(w, "Anotacao_por_post", index=False)
    notas.to_excel(w, "Notas", index=False)

# formatação: % com 1 casa
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
for arq in ["Solicitacao21_Tabelas.xlsx", "Solicitacao21_Evidencias.xlsx"]:
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

json.dump({"eixo": t_eixo.to_dict("records"), "esp_eixo": esp_eixo.to_dict("records"), "top": t_top.to_dict("records"),
           "semloc": t_semloc.to_dict("records"), "per": t_per.to_dict("records"), "lr": t_lr.to_dict("records"),
           "leit": t_leit.to_dict("records"), "niv": t_niv.to_dict("records"),
           "pad": t_pad.to_dict("records"), "div": t_div.to_dict("records")}, open("s21_tabelas.json", "w", encoding="utf8"), ensure_ascii=False, default=str)

pd.set_option("display.width", 250)
print(t_eixo.to_string(index=False)); print(esp_eixo.to_string(index=False)); print(t_top.to_string(index=False))
print(t_semloc.to_string(index=False)); print(t_per.head(12).to_string(index=False)); print(t_per.tail(1).to_string(index=False))
print(t_lr.to_string(index=False)); print(t_leit.to_string(index=False))
for t in (t_niv, t_pad, t_div):
    print(t.to_string(index=False))
