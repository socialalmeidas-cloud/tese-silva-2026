"""Solicitação 16 — Leitura interpretativa da Descrição da Imagem (procedimento de apoio).

O que este script faz (e o que NÃO faz):
  - Seleciona o recorte: posts com "Descrição da Imagem" diferente de "Não analisado" (1.605 posts).
  - Divide o recorte em lotes de até 60 posts, do mesmo perfil, em ordem de linha da planilha.
  - Exporta o texto de cada lote (linha + descrição) para leitura integral.
  - Monta a tabela de evidências a partir do registro de leitura (evidencias_s16.py), conferindo que
    cada trecho citado existe LITERALMENTE na descrição da linha indicada.
  - Confere a cobertura: todas as linhas do recorte lidas, cada uma uma única vez.
  - NÃO interpreta: a leitura e os padrões foram registrados à parte (evidencias_s16.py / memorando),
    a partir da leitura integral das descrições, feita por IA e sujeita a conferência por amostragem.

Uso: python3 solicitacao16.py <planilha.xlsx> [--exportar]
"""
import os
import sys
import json
import pandas as pd

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from correcoes_planilha import aplicar  # correções aprovadas pela autora (ex.: linha 1502)

ARQ = next((a for a in sys.argv[1:] if not a.startswith("--")), "copia.xlsx")
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
DESC = "Descrição da Imagem"
NA = "Não analisado"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
TAM_LOTE = 60
N_ESPERADO = {"Ego": 298, "Roma": 551, "Imperatriz Online": 756}
PASTA_LOTES = os.path.join(AQUI, "s16_lotes")

# 1) Leitura da planilha congelada (cópia) e aplicação das correções aprovadas
df = pd.read_excel(ARQ)
assert df.shape == (4266, 65)
df = aplicar(df)
df["Linha"] = df.index + 2  # número da linha na planilha original (linha 1 = cabeçalho)

# 2) Recorte: só posts com descrição; "Não analisado" nunca entra
rec = df[df[DESC].astype(str).str.strip() != NA].copy()
assert rec["Perfil"].value_counts().to_dict() == N_ESPERADO

# 3) Lotes: por perfil, em ordem de linha, até 60 posts por lote
lotes = []
for p in PERFIS:
    linhas = rec.loc[rec["Perfil"] == p, "Linha"].sort_values().tolist()
    for i in range(0, len(linhas), TAM_LOTE):
        bloco = linhas[i:i + TAM_LOTE]
        lotes.append({"lote": f"{p[:3].upper()}-{i // TAM_LOTE + 1:02d}", "perfil": p,
                      "primeira": bloco[0], "ultima": bloco[-1], "linhas": bloco, "N": len(bloco)})

# 4) Conferência de cobertura: todas as linhas do recorte em exatamente um lote
todas = [l for lt in lotes for l in lt["linhas"]]
assert len(todas) == len(set(todas)) == len(rec) == 1605
assert set(todas) == set(rec["Linha"])
for p in PERFIS:
    assert sum(lt["N"] for lt in lotes if lt["perfil"] == p) == N_ESPERADO[p]

# 5) Exportação do texto de cada lote para leitura integral (opcional)
if "--exportar" in sys.argv:
    os.makedirs(PASTA_LOTES, exist_ok=True)
    d = rec.set_index("Linha")[DESC].astype(str)
    for lt in lotes:
        with open(os.path.join(PASTA_LOTES, lt["lote"] + ".txt"), "w", encoding="utf-8") as f:
            for l in lt["linhas"]:
                f.write(f"[{l}] {d[l].strip()}\n")
    with open(os.path.join(PASTA_LOTES, "lotes.json"), "w", encoding="utf-8") as f:
        json.dump([{k: v for k, v in lt.items()} for lt in lotes], f, ensure_ascii=False)
    print("lotes exportados:", len(lotes))

# 6) Tabela de evidências: confere cada trecho contra a descrição da linha (correspondência literal)
from evidencias_s16 import EVIDENCIAS, PADROES

d = rec.set_index("Linha")
erros = []
linhas_ev = []
for linha, perfil, cod, trecho in EVIDENCIAS:
    if linha not in d.index:
        erros.append((linha, "linha fora do recorte"))
        continue
    if d.loc[linha, "Perfil"] != perfil:
        erros.append((linha, f"perfil diferente: {d.loc[linha, 'Perfil']}"))
    if trecho not in str(d.loc[linha, DESC]):
        erros.append((linha, f"trecho não encontrado: {trecho[:40]}..."))
    linhas_ev.append([linha, perfil, cod, PADROES[perfil][cod], trecho])
assert not erros, erros
ev = pd.DataFrame(linhas_ev, columns=["Linha", "Perfil", "Código", "Padrão", "Trecho da descrição"])

# 7) Perfis já lidos (a leitura é feita perfil a perfil, com autorização da autora entre eles)
PERFIS_LIDOS = sorted(set(ev["Perfil"]), key=PERFIS.index)
cob = pd.DataFrame([{"Lote": lt["lote"], "Perfil": lt["perfil"],
                     "Linhas lidas": f"{lt['primeira']} a {lt['ultima']} ({lt['N']} linhas do recorte)",
                     "N": lt["N"], "Status": "lido" if lt["perfil"] in PERFIS_LIDOS else "aguardando autorização"}
                    for lt in lotes])
soma = cob.groupby("Perfil", sort=False)["N"].sum().reindex(PERFIS).reset_index()
soma["Esperado"] = [N_ESPERADO[p] for p in PERFIS]
assert (soma["N"] == soma["Esperado"]).all()

if "--planilha" in sys.argv:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment
    SAIDA = os.path.join(AQUI, "Solicitacao16_Evidencias.xlsx")
    lista_linhas = pd.DataFrame([{"Lote": lt["lote"], "Perfil": lt["perfil"], "Linha": l}
                                 for lt in lotes for l in lt["linhas"]])
    notas = [
        ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia) + correcoes_planilha.py (linha 1502 e 19 descrições do Imperatriz Online revisadas pela autora em 05/11/2026)."],
        ["Data da leitura", "05/11/2026"],
        ["N analisado", "Recorte de 1.605 posts com Descrição da Imagem diferente de \"Não analisado\" (Ego 298; Roma 551; Imperatriz Online 756). \"Não analisado\" (2.661 posts) não entra. Perfis lidos até agora: " + ", ".join(PERFIS_LIDOS) + "."],
        ["Lotes", f"{len(lotes)} lotes de até {TAM_LOTE} posts, do mesmo perfil, em ordem de linha (Ego 5; Roma 10; Imperatriz Online 13). A aba Cobertura e a aba Linhas_por_lote registram as linhas de cada lote; o script confere que as 1.605 linhas aparecem uma única vez."],
        ["Pergunta da leitura", "Quem ou o que é posto como sujeito da imagem e de que posição o perfil se coloca ao publicá-la."],
        ["Procedimento", "Leitura integral, lote a lote, de todas as descrições do perfil; registro de padrões, evidências (linha + trecho literal) e casos que fogem ao padrão; síntese por perfil no memorando. Os trechos citados foram conferidos automaticamente contra a descrição da linha."],
        ["Leitura feita por IA", "A leitura interpretativa foi feita por IA (Claude, Anthropic) a partir das descrições escritas pela pesquisadora, sem acesso às imagens, e deve ser conferida pela pesquisadora por amostragem."],
        ["Limitações", "1) A leitura vê só o texto das descrições, não as imagens. 2) As descrições já trazem interpretações da pesquisadora (ex.: \"publicidade velada\"); quando citadas, são FATO do texto, não juízo da IA. 3) Descrições variam em extensão e estilo. 4) \"Recorrente\", \"frequente\" e \"raro\" expressam a impressão da leitura integral, sem contagem por categoria. 5) As evidências são exemplos, não a lista exaustiva dos casos."],
        ["Código R19 (Roma)", "Não é padrão do perfil: marca descrições cuja redação limita a leitura (frase-modelo de publicidade em post factual: linhas 705, 722, 784, 1119; lembrança pessoal: 1162; redação mais longa e formal a partir de cerca da linha 975). Esses posts foram lidos pelo conteúdo."],
        ["Código I18 (Imperatriz Online)", "Não é padrão do perfil: marca a redação das descrições. Em 05/11/2026 a autora revisou as 19 descrições com problemas de redação (linhas 1680, 2563, 2931, 3005, 3318, 3357, 3361, 3414, 3457, 3466, 3487, 3507, 3520, 3522, 3557, 3578, 3678, 3859, 3993); a leitura usa o texto revisado (correcoes_planilha.py + revisao_i18/). Na linha 1680, a revisão manteve a última frase repetida. A redação muda ao longo do perfil: a partir de cerca da linha 1820, descrições mais longas e formais; nova variação a partir de cerca da linha 3150; as 19 revisadas seguem um modelo próprio, com a frase \"o Imperatriz Online assume o papel de mediador e difusor\"."],
        ["Formato", "Linha = linha da planilha original (linha 1 = cabeçalho). Código = padrão do perfil (ver memorando). Trecho = cópia literal da descrição, com a grafia original."],
        ["Script", "solicitacao16.py (+ evidencias_s16.py, correcoes_planilha.py)"],
    ]
    with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
        ev.to_excel(w, sheet_name="Evidencias", index=False)
        cob.to_excel(w, sheet_name="Cobertura", index=False)
        soma.to_excel(w, sheet_name="Cobertura", index=False, startrow=len(cob) + 2)
        lista_linhas.to_excel(w, sheet_name="Linhas_por_lote", index=False)
        pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, sheet_name="Notas", index=False)
    wb = load_workbook(SAIDA)
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font = Font(bold=True)
    ws = wb["Evidencias"]
    for L, wd in zip("ABCDE", (8, 10, 8, 60, 90)):
        ws.column_dimensions[L].width = wd
    for row in ws.iter_rows(min_row=2):
        for c in row[3:5]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws = wb["Cobertura"]
    for c in ws[len(cob) + 3]:
        c.font = Font(bold=True)
    for L, wd in zip("ABCDE", (10, 20, 40, 8, 24)):
        ws.column_dimensions[L].width = wd
    ws = wb["Notas"]
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 110
    for row in ws.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(SAIDA)
    print("evidências:", len(ev), "| perfis lidos:", PERFIS_LIDOS)

# 8) Exporta as evidências conferidas para o gerador do memorando (memorando_s16.js)
if "--json" in sys.argv:
    ev.to_json(os.path.join(AQUI, "s16_evidencias.json"), orient="records", force_ascii=False)
    print("json ok")
