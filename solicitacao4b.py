"""Solicitação 4B — Eixo × Mediação: teste qui-quadrado refeito em três versões.
(a) 10 eixos mais frequentes do total geral, sem Publicidade (= 9 eixos) × 5 mediações
(b) idem, com Institucional + Opinativa unidas
(c) idem (b), p-valor por simulação de Monte Carlo (margens fixas; 10.000 simulações)
Uso: python3 solicitacao4b.py <planilha.xlsx>
"""
import sys
import numpy as np
import pandas as pd
from scipy.stats import chi2_contingency
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao4B_Teste_refeito.xlsx"
EIXO, MED = "Eixo Temático", "Mediação da Informação"
MEDS = ["Factual", "Publicitária", "De Entretenimento", "Institucional", "Opinativa"]
SEMENTE = 20261104
B = 10_000

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
assert df["Perfil"].value_counts().to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}

freq = df[EIXO].value_counts().rename_axis("e").reset_index(name="n").sort_values(["n", "e"], ascending=[False, True])
TOP10 = freq["e"].iloc[:10].tolist()
EIXOS = [e for e in TOP10 if e != "Publicidade"]
assert len(EIXOS) == 9

sub = df[df[EIXO].isin(EIXOS)].copy()
sub["Med_b"] = sub[MED].replace({"Institucional": "Institucional/Opinativa", "Opinativa": "Institucional/Opinativa"})
MEDS_B = ["Factual", "Publicitária", "De Entretenimento", "Institucional/Opinativa"]

obs_a = pd.crosstab(sub[EIXO], sub[MED]).reindex(index=EIXOS, columns=MEDS, fill_value=0)
obs_b = pd.crosstab(sub[EIXO], sub["Med_b"]).reindex(index=EIXOS, columns=MEDS_B, fill_value=0)
assert obs_a.values.sum() == obs_b.values.sum() == len(sub)


def chi2_stat(o):
    e = o.sum(1, keepdims=True) * o.sum(0, keepdims=True) / o.sum()
    return ((o - e) ** 2 / e).sum()


def resumo(obs, rotulo, monte_carlo=False):
    o = obs.values
    chi2, p, gl, esp = chi2_contingency(o, correction=False)
    n = o.sum()
    v = np.sqrt(chi2 / (n * (min(o.shape) - 1)))
    r = {"Versão": rotulo, "N": int(n), "Tabela": f"{o.shape[0]} eixos × {o.shape[1]} mediações",
         "χ²": chi2, "Graus de liberdade": int(gl), "p-valor": p, "Método do p-valor": "Assintótico (distribuição χ²)",
         "V de Cramér": v,
         "Células esperada < 5 (N)": int((esp < 5).sum()), "Células esperada < 5 (%)": (esp < 5).mean(),
         "Células esperada < 1 (N)": int((esp < 1).sum()), "Células esperada < 1 (%)": (esp < 1).mean(),
         "Total de células": esp.size}
    if p == 0:
        r["p-valor"] = "p < 0,001 (menor que a precisão numérica; o programa retorna 0)"
    if monte_carlo:
        # Permutação dos rótulos de mediação entre os posts: preserva os totais de linha e coluna
        rng = np.random.default_rng(SEMENTE)
        linhas = np.repeat(np.arange(o.shape[0]), o.sum(1))
        cols = np.concatenate([np.repeat(np.arange(o.shape[1]), o[i]) for i in range(o.shape[0])])
        cont = 0
        for _ in range(B):
            perm = rng.permutation(cols)
            t = np.zeros_like(o)
            np.add.at(t, (linhas, perm), 1)
            if chi2_stat(t.astype(float)) >= chi2 - 1e-9:
                cont += 1
        r["p-valor"] = (cont + 1) / (B + 1)
        r["Método do p-valor"] = (f"Monte Carlo: {B:,} simulações com margens fixas (permutação), semente {SEMENTE}; "
                                  f"p = (nº de χ² simulados ≥ observado + 1) ÷ (simulações + 1); "
                                  f"χ² simulados ≥ observado: {cont}").replace(",", ".")
        r["Células esperada < 5 (critério)"] = "Não se aplica ao p-valor simulado (não depende da aproximação χ²)"
    crit = r["Células esperada < 5 (%)"] <= 0.20 and r["Células esperada < 1 (N)"] == 0
    if monte_carlo:
        r["Atende ao critério ≤ 20% esperada < 5?"] = ("Sim" if crit else "Não") + " — mas o p-valor simulado não depende desse critério"
    else:
        r["Atende ao critério ≤ 20% esperada < 5?"] = "Sim" if r["Células esperada < 5 (%)"] <= 0.20 else "Não"
    return r, pd.DataFrame(esp, index=obs.index, columns=obs.columns)


ra, esp_a = resumo(obs_a, "(a) 9 eixos, sem Publicidade, 5 mediações")
rb, esp_b = resumo(obs_b, "(b) = (a) com Institucional/Opinativa unidas")
rc, _ = resumo(obs_b, "(c) = (b) com p-valor por Monte Carlo", monte_carlo=True)

ordem = ["Versão", "Tabela", "N", "χ²", "Graus de liberdade", "p-valor", "Método do p-valor", "V de Cramér",
         "Células esperada < 5 (N)", "Células esperada < 5 (%)", "Células esperada < 1 (N)",
         "Células esperada < 1 (%)", "Total de células", "Atende ao critério ≤ 20% esperada < 5?"]
comp = pd.DataFrame([ra, rb, rc])[ordem].set_index("Versão").T.reset_index().rename(columns={"index": "Item"})


def pct_linha(obs, cats):
    out = pd.DataFrame(index=obs.index)
    out["N do eixo"] = obs.sum(axis=1)
    for m in cats:
        out[f"{m} (N)"] = obs[m]
        out[f"{m} (%)"] = obs[m] / out["N do eixo"]
    tot = obs.sum()
    out.loc["Total"] = [tot.sum()] + sum([[tot[m], tot[m] / tot.sum()] for m in cats], [])
    out = out.reset_index().rename(columns={"index": "Eixo", EIXO: "Eixo"})
    for c in [x for x in out.columns if "(N)" in x or x == "N do eixo"]:
        out[c] = out[c].astype(int)
    pc = out[[f"{m} (%)" for m in cats]].sum(axis=1)
    assert np.allclose(pc, 1)
    return out


tab_a = pct_linha(obs_a, MEDS)
tab_b = pct_linha(obs_b, MEDS_B)
assert tab_a.iloc[-1]["N do eixo"] == len(sub)


def com_total(esp):
    e = esp.copy()
    e.loc["Total"] = e.sum()
    return e.reset_index().rename(columns={"index": "Eixo", EIXO: "Eixo"})


notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", f"{len(sub):,} posts dos 9 eixos (de 4.266 no corpus)".replace(",", ".")],
    ["Exclusões", "Eixo Publicidade (1.742 posts) e os 24 eixos fora do top 10 (588 posts). Total excluído: 2.330."],
    ["Eixos incluídos (9)", "; ".join(EIXOS) + ". São os 10 mais frequentes do total geral (Solicitação 4), sem Publicidade; não houve substituição pelo 11º eixo."],
    ["Fórmula do %", "Percentual de linha: % = N da mediação no eixo ÷ N do eixo × 100. Cada linha soma 100%."],
    ["Versão (b)/(c)", "Institucional (N do recorte informado na Tabela_a) e Opinativa somadas em \"Institucional/Opinativa\"."],
    ["χ² e V de Cramér", "Pearson sem correção de continuidade (scipy.stats.chi2_contingency, correction=False). V = √(χ² / (N × (min(linhas, colunas) − 1)))."],
    ["Monte Carlo (c)", f"{B:,} tabelas simuladas por permutação dos rótulos de mediação entre os posts (margens de linha e coluna fixas); semente {SEMENTE} (numpy.random.default_rng). p = (b + 1) ÷ (B + 1); menor p possível = 1 ÷ 10.001 ≈ 0,0001.".replace(",", ".")],
    ["Critério", "Até 20% das células com esperada < 5 e nenhuma com esperada < 1 (critério informado na solicitação; a referência metodológica precisa ser verificada pela autora)."],
    ["Categorias emergentes", "Nenhuma categoria emergente (†) pertence a Eixo Temático ou Mediação."],
    ["Script", "solicitacao4b.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    comp.to_excel(w, sheet_name="Teste", index=False)
    tab_a.to_excel(w, sheet_name="Tabela_a", index=False)
    tab_b.to_excel(w, sheet_name="Tabela_b_c", index=False)
    com_total(esp_a).to_excel(w, sheet_name="Esperadas_Teste_a", index=False)
    com_total(esp_b).to_excel(w, sheet_name="Esperadas_Teste_b_c", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)

wb = load_workbook(SAIDA)
ws = wb["Teste"]
fmt = {"N": "#,##0", "χ²": "#,##0.00", "p-valor": "0.0000", "V de Cramér": "0.000",
       "Células esperada < 5 (%)": "0.0%", "Células esperada < 1 (%)": "0.0%"}
for r in ws.iter_rows(min_row=2):
    for c in r[1:]:
        if r[0].value in fmt and isinstance(c.value, (int, float)):
            c.number_format = fmt[r[0].value]
            if r[0].value == "p-valor" and c.value < 0.0001:
                c.number_format = "0.0E+00"
        c.alignment = Alignment(wrap_text=True, vertical="top")
for c in ws[1]:
    c.font = Font(bold=True)
ws.column_dimensions["A"].width = 38
for L in "BCD":
    ws.column_dimensions[L].width = 42
for nome in ["Tabela_a", "Tabela_b_c"]:
    ws = wb[nome]
    for c in ws[1] + ws[ws.max_row]:
        c.font = Font(bold=True)
    for col in ws.columns:
        h = str(col[0].value)
        for c in col[1:]:
            c.number_format = "0.0%" if "(%)" in h else "#,##0" if ("(N)" in h or h == "N do eixo") else c.number_format
        ws.column_dimensions[col[0].column_letter].width = max(12, len(h) + 2)
    ws.column_dimensions["A"].width = 22
for nome in ["Esperadas_Teste_a", "Esperadas_Teste_b_c"]:
    ws = wb[nome]
    for c in ws[1] + ws[ws.max_row]:
        c.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2, min_col=2):
        for c in row:
            c.number_format = "0.0"
    ws.column_dimensions["A"].width = 22
ws = wb["Notas"]
ws["A1"].font = ws["B1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 26
ws.column_dimensions["B"].width = 110
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

pd.set_option("display.width", 250)
print(comp.to_string())
print(tab_a.to_string()); print(tab_b.to_string())
print(esp_a.round(1).to_string()); print(esp_b.round(1).to_string())
