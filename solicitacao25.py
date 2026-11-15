"""
Solicitação 25 — O que não aparece nos perfis.

Uso:  python3 -I solicitacao25.py copia.xlsx

Base (80%): leitura de TODAS as descrições do recorte de cada perfil (Ego 298; Roma 551;
Imperatriz Online 756), em lotes de até 60 (arquivos s25b_<perfil>.txt, gerados abaixo).
Cada post foi anotado em s25b_anot_<Perfil>.tsv:
  quem   — quem aparece (texto curto);
  pos    — posição: C centro, F fundo, N só nomeado, S sem rosto/ocultado ("-" ninguém);
  cena   — espaço em cena (EVENTO, COMERCIO, CONSULTORIO, RUA, CASA, ARTE = só arte gráfica, "-" nenhum...);
  citado — espaços só citados no texto/legenda;
  escala — I Imperatriz, R região/Maranhão, B Brasil/exterior, "-" não situado;
  sit    — situação (PUBLI, SHOW, HUMOR, CONQUISTA, CELEB, POLITICA, CRIME...).
A linha de base (o que é visível) e os padrões de ausência estão em s25b_padroes_<Perfil>.json.

Ausência não deixa evidência: nada aqui afirma exclusão ou intenção dos perfis.
"Não observado" = não observado neste recorte.
Nível 3 (referência externa, INFERÊNCIA) em nivel3_s25.py: lista de bairros (cepbrasil.org, NÃO oficial,
ref_bairros_cepbrasil.txt), religião do Censo 2022 (IBGE via Imperatriz Notícias/UFMA) e cor/raça (só estadual).
Indígenas/quilombolas, LGBT e gestor: sem referência externa obtida — não rodado no nível 3.
A planilha original não é alterada (lê-se a cópia, com as correções aprovadas).
"""
import os
import re
import sys
import json
import collections

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
import correcoes_planilha

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
FAIXA = {"Ego": (2, 415), "Roma": (416, 1546), "Imperatriz Online": (1547, 4267)}
REFERENCIA_EXTERNA = "bairros: cepbrasil.org/maranhao/imperatriz; religião: IBGE Censo 2022 via imperatriznoticias.ufma.br; cor/raça: IBGE via O Imparcial (estadual)"
import nivel3_s25 as n3


def rotulos(c):
    """Multirrótulo em ";" (o ";" dentro de parênteses não separa), sem 'Não há'/'Não analisado'."""
    if not isinstance(c, str):
        return []
    c = re.sub(r"\(([^)]*)\)", lambda m: "(" + m.group(1).replace(";", "§") + ")", c)
    ts = [t.strip().replace("§", ";") for t in c.split(";") if t.strip() and t.strip() not in ("Não há", "Não analisado")]
    return list(dict.fromkeys(ts))


def perfil_da_linha(l):
    return next(p for p, (a, b) in FAIXA.items() if a <= int(l) <= b)


df = correcoes_planilha.aplicar(pd.read_excel(ARQ))
assert len(df) == 4266
df["Linha"] = df.index + 2
NP = df["Perfil"].value_counts().to_dict()
DESC = json.load(open("desc_recorte.json", encoding="utf8"))
df["recorte"] = df["Linha"].astype(str).isin(DESC)

# ------------------------------------------------ arquivos de leitura (lotes de até 60, todo o recorte)
for p in PERFIS:
    s = df[df.recorte & (df.Perfil == p)]
    with open(f"s25b_{p.split()[0]}.txt", "w", encoding="utf8") as f:
        for i, (_, r) in enumerate(s.iterrows()):
            if i % 60 == 0:
                f.write(f"===== LOTE {i // 60 + 1}\n")
            f.write(f"[{r.Linha}|R:{r['Regime de Visibilidade']}|E:{r['Eixo Temático']}|L:{r['Referência Local']}] {DESC[str(r.Linha)]}\n")

# ================================================ NÍVEL 1 (apoio): busca de termos nas descrições dos três perfis
# Serve para CONFIRMAR presença em um perfil (trecho literal) e sinalizar onde um termo não ocorre.
# A ausência no perfil analisado é confirmada pela leitura integral; a busca não substitui a leitura.
TERMOS = {
    "Pessoa detida": r"algemad|\bdetid[oa]s?\b|\bpres[oa]s?\b",
    "Corpo de vítima, IML, caixão": r"cad[aá]ver|corpo (da vítima|estirado|estendido|no chão|coberto|caído|ensanguentado)|\bIML\b|caixão",
    "Policial": r"policia(l|is)\b|\bPM\b|PRF|polícia",
    "Protesto ou interdição": r"protest|manifesta[çc][ãa]o|interdit",
    "Buraco, lama, esgoto, lixo, alagamento": r"buraco|\blama\b|esgoto|\blixo|alagad|enchente",
    "UPA, hospital, posto de saúde": r"\bUPA\b|hospital|posto de saúde|Socorrão|maternidade",
    "Situação de rua": r"situação de rua|morador(es)? de rua",
    "Igreja, missa, culto": r"igreja|\bmissa\b|\bculto\b",
    "Terreiro, matriz africana": r"terreiro|umbanda|candombl|macumba",
    "Povos indígenas, quilombolas": r"ind[ií]gen|aldeia|quilomb",
    "LGBT": r"LGBT",
    "Vereador, câmara, prefeitura": r"vereador|câmara municipal|prefeitura",
    "Acidente de trânsito": r"acidente|colis[ãa]o|capot|atropel",
    "Trabalhador, greve, emprego": r"trabalhador|greve|sindicat|vagas de emprego",
    "Procedimento estético": r"antes e depois|harmoniza|transplante capilar|botox|lifting|peeling",
    "Carlinhos Maia": r"Carlinhos",
    "Camarote, open bar": r"camarote|open bar",
    "Exterior (Paris, EUA, Tokyo...)": r"Paris|EUA|Estados Unidos|Nova York|Tokyo|México|Espanha|Itália|Madrid|San Diego",
}
lin1 = []
for nome, rx in TERMOS.items():
    c = collections.Counter(perfil_da_linha(l) for l, d in DESC.items() if re.search(rx, d))
    ex = {p: next((l for l, d in DESC.items() if perfil_da_linha(l) == p and re.search(rx, d)), "") for p in PERFIS}
    lin1.append({"Termo": nome, **{f"{p} (posts do recorte)": c.get(p, 0) for p in PERFIS},
                 **{f"{p} — 1ª linha": ex[p] for p in PERFIS},
                 "Não observado em": ", ".join(p for p in PERFIS if c.get(p, 0) == 0)})
t_n1 = pd.DataFrame(lin1)

# ================================================ NÍVEL 2: dentro do banco (por perfil, posts do perfil)
n2_cat, n2_uni, n2_ni = [], [], []
for p in PERFIS:
    s = df[df.Perfil == p]
    reg = collections.Counter(t for x in s["Regime de Visibilidade"] for t in rotulos(x))
    eix = s["Eixo Temático"].value_counts()
    n2_cat += [{"Perfil": p, "Variável": "Regime de Visibilidade", "Categoria": k, "N (posts do perfil)": v,
                "Marca": "*" if v < 10 else ""} for k, v in sorted(reg.items(), key=lambda x: -x[1]) if v < 30]
    n2_cat += [{"Perfil": p, "Variável": "Eixo Temático", "Categoria": k, "N (posts do perfil)": int(v),
                "Marca": "*" if v < 10 else ""} for k, v in eix.items() if v < 30]
    loc = collections.Counter(t for x in s["Referência Local"] for t in rotulos(x) if t != "Não identificado")
    for k, v in loc.items():
        if v == 1:
            l = int(s[s["Referência Local"].apply(lambda x: k in rotulos(x))].Linha.iloc[0])
            n2_uni.append({"Perfil": p, "Local (uma só vez)": k, "Linha": l, "No recorte": "sim" if str(l) in DESC else "não"})
    ni = s[s["Referência Local"].astype(str).str.contains("Não identificado")]
    for _, r in ni.iterrows():
        n2_ni.append({"Perfil": p, "Linha": int(r.Linha), "No recorte": "sim" if r.recorte else "não"})
t_cat, t_uni, t_ni = pd.DataFrame(n2_cat), pd.DataFrame(n2_uni), pd.DataFrame(n2_ni)

# ================================================ ETAPA 1 (linha de base) e ETAPA 3 (padrões): leitura
lidos = [p for p in PERFIS if os.path.exists(f"s25b_anot_{p.split()[0]}.tsv")]
base, base_txt, pad, ev, post_pad, ni_leit = [], [], [], [], [], []
for p in lidos:
    a_ = pd.read_csv(f"s25b_anot_{p.split()[0]}.tsv", sep="\t", dtype=str).fillna("-")
    a_["linha"] = a_["linha"].astype(int)
    rec_p = set(df.loc[df.recorte & (df.Perfil == p), "Linha"])
    assert set(a_["linha"]) == rec_p and len(a_) == len(rec_p), f"anotação incompleta: {p}"
    for col, rot in (("cena", "Espaço em cena"), ("escala", "Escala"), ("sit", "Situação")):
        for k, v in a_[col].value_counts().items():
            base.append({"Perfil": p, "Dimensão": rot, "Categoria": k, "Posts lidos": int(v)})
    for k, v in collections.Counter(x for s_ in a_.pos for x in s_.split(",")).items():
        base.append({"Perfil": p, "Dimensão": "Posição de quem aparece (multirrótulo)", "Categoria": k, "Posts lidos": v})
    P = json.load(open(f"s25b_padroes_{p.split()[0]}.json", encoding="utf8"))
    base_txt += [{"Perfil": p, "Linha de base": t} for t in P["linha_de_base"]]
    if P.get("nota_leitura"):
        base_txt.append({"Perfil": p, "Linha de base": "Nota de método: " + P["nota_leitura"]})
    for q in P["padroes"]:
        linhas = list(a_.loc[a_.escala == "B", "linha"]) if q["linhas"] == "escala B" else q["linhas"]
        assert set(linhas) <= rec_p, (q["cod"], set(linhas) - rec_p)
        pad.append({"Perfil": p, "Código": q["cod"], "Padrão": q["nome"], "Nível": q["nivel"], "FATO/INFERÊNCIA": q["tipo"],
                    "Definição": q["def"], "N no recorte do perfil": len(linhas), "Marca": "*" if len(linhas) < 10 else "",
                    "Contraexemplo": f"linha {q['contra'][0]}: {q['contra'][1]}", "Explicação alternativa": q["alternativa"]})
        post_pad += [{"Perfil": p, "Código": q["cod"], "Linha": int(l)} for l in linhas]
        for l, t in q["ev"] + [q["contra"]]:
            assert t in DESC[str(l)], (l, t)                                # trecho literal
            ev.append({"Linha": l, "Perfil": perfil_da_linha(l), "Nível": q["nivel"],
                       "Padrão": f"{q['cod']} {q['nome']}" + (" (contraexemplo)" if [l, t] == q["contra"] else ""),
                       "Trecho": t, "FATO/INFERÊNCIA": "FATO"})
    m = df[df.Perfil == p].merge(a_, left_on="Linha", right_on="linha")
    for _, r in m[m["Referência Local"].astype(str).str.contains("Não identificado")].iterrows():
        ni_leit.append({"Perfil": p, "Linha": r.Linha, "Espaço em cena (leitura)": r.cena, "Lugar citado (leitura)": r.citado, "Escala (leitura)": r.escala})

# ================================================ NÍVEL 3: contra referência externa (INFERÊNCIA)
t_bai, t_fora, t_rel, t_relres, t_ibge, t_fontes = n3.rodar(df, DESC, PERFIS, rotulos, perfil_da_linha, FAIXA)
P3, OBS3, NENHUM3 = n3.padroes_n3(t_bai)
for q in P3:
    pad.append({"Perfil": q["perfil"], "Código": q["cod"], "Padrão": q["nome"], "Nível": 3, "FATO/INFERÊNCIA": "INFERÊNCIA",
                "Definição": q["def"], "N no recorte do perfil": q["linhas_n"], "Marca": "*" if q["linhas_n"] < 10 else "",
                "Contraexemplo": f"linha {q['contra'][0]}: {q['contra'][1]}", "Explicação alternativa": q["alt"]})
    for l, t in q["ev"] + [q["contra"]]:
        if l != "—":
            assert t in DESC[str(l)], (q["cod"], l, t)
        ev.append({"Linha": l, "Perfil": "Referência externa" if l == "—" else perfil_da_linha(l), "Nível": 3,
                   "Padrão": f"{q['cod']} {q['nome']}" + (" (contraexemplo)" if [l, t] == q["contra"] else ""),
                   "Trecho": t, "FATO/INFERÊNCIA": "INFERÊNCIA"})
for l, c, t in n3.SO_FATO:     # sem referência externa: só o banco (FATO)
    ev.append({"Linha": l, "Perfil": perfil_da_linha(l), "Nível": 2, "Padrão": f"{c} — presença no banco (nível 3 não rodado: sem referência)",
               "Trecho": t, "FATO/INFERÊNCIA": "FATO"})

t_base, t_btxt, t_pad = pd.DataFrame(base), pd.DataFrame(base_txt), pd.DataFrame(pad)
t_ev = pd.DataFrame(ev, columns=["Linha", "Perfil", "Nível", "Padrão", "Trecho", "FATO/INFERÊNCIA"])
t_pp, t_nil = pd.DataFrame(post_pad), pd.DataFrame(ni_leit)

notas = pd.DataFrame([
    ["Pergunta", "O que fica fora de cena: sujeitos, espaços e situações pouco presentes, só citados, presentes em um perfil e não em outro, ou não observados."],
    ["Cautela", "Ausência não deixa evidência. \"Não observado\" = não observado neste recorte; nada afirma exclusão, intenção ou inexistência na cidade. Cada padrão traz uma explicação alternativa (escopo editorial, tipo de conteúdo, formato)."],
    ["Base de leitura", "Todas as descrições do recorte (Ego 298; Roma 551; Imperatriz Online 756), em lotes de até 60. Perfis lidos nesta solicitação: " + (", ".join(lidos) or "nenhum") + "."],
    ["Nível 1", "Comparativo entre perfis (FATO). Presença confirmada por trecho literal; ausência no perfil analisado confirmada pela leitura integral. A busca de termos (aba Nivel1_termos) é apoio e pode ter falsos positivos."],
    ["Nível 2", "Dentro do banco (FATO), por perfil: categorias de Regime (multirrótulo) e de Eixo com menos de 30 posts; locais da Referência Local (multirrótulo) que aparecem uma só vez; posts com \"Não identificado\"; sujeitos e espaços só citados."],
    ["Nível 3", "Contra referência externa (INFERÊNCIA; a comparação depende da fonte). Rodado: (a) lista de bairros de cepbrasil.org/maranhao/imperatriz (NÃO oficial, 114 nomes) x Referência Local de todos os posts e descrições do recorte; (b) religião em Imperatriz, Censo 2022 (IBGE), via Imperatriz Notícias/UFMA (31/10/2025) x menções religiosas lidas nas descrições; (c) cor/raça: só referência estadual (66% pardos no MA, O Imparcial, 22/12/2023) — comparação não aplicável. Não rodado por falta de referência: povos indígenas e quilombolas (municipal), pessoas LGBT, gestor em cena. Todos os dados externos devem ser conferidos pela autora na fonte primária (IBGE/SIDRA, Prefeitura)."],
    ["Nível 3 — variantes", "A Referência Local foi casada com a lista sem acento e com variantes de grafia listadas em nivel3_s25.VARIANTES (ex.: \"Parque Alvorada I\" = Parque Alvorada). Homônimos de município (Santa Inês, Santa Rita, Santa Luzia, Centro Novo etc.) estão marcados. Variantes duvidosas não foram aceitas (Planalto, Jardim Sumaré, Vila JK)."],
    ["Contagens", "Padrões: posts do recorte (leitura), sem percentual. Nível 2: posts do perfil (Ego 414; Roma 1.131; Imperatriz Online 2.721). * = menos de 10."],
    ["Base quantitativa", "4.266 posts (cópia da planilha congelada de 04/11/2026, com as correções aprovadas)."],
    ["Limites", "As imagens não foram vistas; a leitura usa as descrições da pesquisadora, que às vezes já trazem interpretação. Códigos atribuídos por IA (Claude, Anthropic), a conferir."],
    ["Evidências", f"{len(t_ev)} trechos, todos cópia literal das descrições."],
], columns=["Item", "Descrição"])

with pd.ExcelWriter("Solicitacao25_Linha_de_base.xlsx", engine="openpyxl") as w:
    t_btxt.to_excel(w, sheet_name="Linha_de_base_texto", index=False)
    t_base.to_excel(w, sheet_name="Linha_de_base_contagens", index=False)
    t_pad.to_excel(w, sheet_name="Padroes_de_ausencia", index=False)
    t_n1.to_excel(w, sheet_name="Nivel1_termos", index=False)
    t_cat.to_excel(w, sheet_name="Nivel2_categorias_menos30", index=False)
    t_uni.to_excel(w, sheet_name="Nivel2_locais_uma_vez", index=False)
    t_ni.to_excel(w, sheet_name="Nivel2_nao_identificado", index=False)
    t_nil.to_excel(w, sheet_name="Nivel2_NI_x_leitura", index=False)
    t_fontes.to_excel(w, sheet_name="Nivel3_fontes", index=False)
    t_bai.to_excel(w, sheet_name="Nivel3_bairros", index=False)
    t_fora.to_excel(w, sheet_name="Nivel3_locais_fora_lista", index=False)
    t_ibge.to_excel(w, sheet_name="Nivel3_religiao_IBGE", index=False)
    t_relres.to_excel(w, sheet_name="Nivel3_religiao_resumo", index=False)
    t_rel.to_excel(w, sheet_name="Nivel3_religiao_posts", index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)
with pd.ExcelWriter("Solicitacao25_Evidencias.xlsx", engine="openpyxl") as w:
    t_ev.to_excel(w, sheet_name="Evidencias", index=False)
    t_pad.to_excel(w, sheet_name="Padroes", index=False)
    t_pp.to_excel(w, sheet_name="Posts_por_padrao", index=False)
    pd.concat([pd.read_csv(f"s25b_anot_{p.split()[0]}.tsv", sep="\t", dtype=str).assign(Perfil=p) for p in lidos]).to_excel(w, sheet_name="Anotacao_por_post", index=False)
    notas.to_excel(w, sheet_name="Notas", index=False)

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
for arq in ["Solicitacao25_Linha_de_base.xlsx", "Solicitacao25_Evidencias.xlsx"]:
    wb = load_workbook(arq)
    for ws in wb:
        for c in ws[1]:
            c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9E2F3")
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in ws.iter_cols(min_row=1):
            h = str(col[0].value or "")
            ws.column_dimensions[col[0].column_letter].width = 60 if h in ("Descrição", "Trecho", "Definição", "Fonte", "Uso", "Conferir", "Contraexemplo", "Explicação alternativa", "Linha de base", "Padrão") else max(12, min(30, len(h) + 2))
        ws.freeze_panes = "A2"
    wb.save(arq)

json.dump({"n1": t_n1.to_dict("records"), "cat": t_cat.to_dict("records"), "uni": t_uni.to_dict("records"),
           "ni": t_ni.to_dict("records"), "nil": t_nil.to_dict("records"), "base": t_base.to_dict("records"),
           "pad": t_pad.to_dict("records"), "bai": t_bai.to_dict("records"), "fora": t_fora.to_dict("records"),
           "relres": t_relres.to_dict("records"), "ibge": t_ibge.to_dict("records"), "fontes": t_fontes.to_dict("records"),
           "obs3": OBS3, "nenhum3": NENHUM3},
          open("s25_tabelas.json", "w", encoding="utf8"), ensure_ascii=False, default=str)

json.dump({q["cod"]: {"ev": q["ev"], "contra": q["contra"]} for q in P3}, open("s25_n3.json", "w", encoding="utf8"), ensure_ascii=False)

pd.set_option("display.width", 250); pd.set_option("display.max_colwidth", 50)
print(t_n1.iloc[:, :5].to_string(index=False))
print(t_cat[t_cat.Perfil == "Ego"].to_string(index=False))
print(t_pad.drop(columns=["Definição", "Explicação alternativa"]).to_string(index=False))
print(len(t_ev), "evidências;", "NI Ego:", len(t_ni[t_ni.Perfil == "Ego"]), "; locais únicos Ego:", len(t_uni[t_uni.Perfil == "Ego"]))
