"""Números do memorando S24 (Roma) a partir das anotações: compara v1 (leitura truncada) e v2 (leitura integral)."""
import sys, os, re, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd, correcoes_planilha

def rotulos(c):
    if not isinstance(c, str):
        return []
    c = re.sub(r"\(([^)]*)\)", lambda m: "(" + m.group(1).replace(";", "§") + ")", c)
    ts = [t.strip().replace("§", ";") for t in c.split(";") if t.strip() and t.strip() not in ("Não há", "Não analisado")]
    return list(dict.fromkeys(ts))

df = correcoes_planilha.aplicar(pd.read_excel(sys.argv[1]))
df["Linha"] = (df.index + 2).astype(str)
df = df.set_index("Linha")
COD = {"CONSUMIDOR_ROMA": "R1", "TESTEMUNHA_CRIME": "R2", "COLABORADOR": "R3", "ALERTA": "R4",
       "RISO_CRIME": "R5", "FA_ROMA": "R6", "CIDADAO": "R7", "DENUNCIA": "R7", "FESTA": "R8X"}
REM = re.compile(r"post(o|agem|agens)? anterior|postagens anteriores|atrás|anteriormente|story anterior|postado|na época", re.I)

def numeros(arq):
    a = pd.read_csv(arq, sep="\t", dtype=str, keep_default_na=False).set_index("linha")
    a["R"] = a.padrao.map(COD)
    j = a.join(df[["CTA", "Linguagem Regional", "Collab", "Mediação da Informação"]])
    j["cta_l"] = j.CTA.apply(rotulos); j["cta"] = j.cta_l.str.len() > 0
    j["lr"] = j["Linguagem Regional"].apply(rotulos).str.len() > 0
    j["dir"] = j.dirige.ne("-"); j["sup"] = j.supoe.ne("-"); j["exi"] = j.exige.ne("-")
    j["rem"] = j.dirige.str.contains(REM)
    j["ci"] = j.cta_l.apply(lambda x: "Consumo interno" in x)
    j["cont"] = j.cta_l.apply(lambda x: "Contato" in x)
    j["acao"] = j.cta_l.apply(lambda x: any("Ação" in t or "Visita" in t for t in x))
    n = {}
    n["padroes"] = j.R.value_counts().sort_index().to_dict()
    n["dirige"] = int(j.dir.sum()); n["remissao"] = int(j.rem.sum())
    n["supoe"] = int(j.sup.sum()); n["supoe_caso"] = int(j.supoe.str.startswith("ter visto").sum())
    ml = j.reacao.str.split(";").explode().str.strip(); n["reacoes"] = ml[ml.ne("-")].value_counts().head(10).to_dict()  # multirrótulo, como no memorando
    n["exige"] = int(j.exi.sum())
    n["exige_img"] = int(j.exige.str.startswith("aceitar").sum())
    n["exige_corpo"] = int(j.exige.str.contains("corpo", case=False).sum())
    n["exige_humor"] = int(j.exige.str.contains("humor|trocadilho", case=False).sum())
    n["exige_ler"] = int(j.exige.str.contains(r"^(?:ler|ouvir)", case=False).sum())
    n["dir_sem_cta"] = int((j.dir & ~j.cta).sum()); n["dir_sem_cta_R"] = j[j.dir & ~j.cta].R.value_counts().to_dict()
    n["cta_sem_dir"] = int((~j.dir & j.cta).sum()); n["cta_sem_dir_R"] = j[~j.dir & j.cta].R.value_counts().to_dict()
    n["ci"] = int(j.ci.sum()); n["ci_e_rem"] = int((j.ci & j.rem).sum()); n["so_ci"] = int((j.ci & ~j.rem).sum()); n["so_rem"] = int((~j.ci & j.rem).sum())
    n["sup_sem_lr"] = int((j.sup & ~j.lr).sum())
    g = j.groupby("R")
    n["R_cta"] = g.cta.sum().to_dict(); n["R_contato"] = g.cont.sum().to_dict(); n["R_acao"] = g.acao.sum().to_dict()
    n["R_lr"] = g.lr.sum().to_dict(); n["R_collab"] = g.Collab.apply(lambda s: s.astype(str).str.strip().eq("Sim").sum()).to_dict()
    n["R_med"] = g["Mediação da Informação"].agg(lambda x: x.value_counts().head(2).to_dict()).to_dict()
    return n

v1 = numeros("s24b_anot_Roma_v1_leitura_truncada.tsv"); v2 = numeros("s24b_anot_Roma.tsv")
for k in v1:
    print(k, "\n  v1:", v1[k], "\n  v2:", v2[k]) if v1[k] != v2[k] else print(k, "=", v1[k])
json.dump({"v1": v1, "v2": v2}, open("s24c_numeros_roma.json", "w", encoding="utf8"), ensure_ascii=False, indent=1, default=int)
