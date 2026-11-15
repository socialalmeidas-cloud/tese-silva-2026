"""Gera Revisao_I18_descricoes.xlsx: descrições do Imperatriz Online com problemas de redação
(código I18 da Solicitação 16) para a autora revisar. Não altera a planilha; as revisões
aprovadas entram depois em correcoes_planilha.py."""
import sys, pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
sys.path.insert(0, '.')
from privado import carregar
from correcoes_planilha import aplicar

df = aplicar(pd.read_excel('copia.xlsx')); df['Linha'] = df.index + 2
col = [c for c in df.columns if c.startswith('Descri')][0]
txt = df.set_index('Linha')[col].astype(str)
assert (df.set_index('Linha').loc[list(range(1548, 1549)), 'Perfil'] == 'Imperatriz Online').all()

L = carregar('revisao_i18', 'L', [(2, "1 Alta", "Problema fictício de exemplo.")])
u = df.set_index("Linha")["url "].astype(str).str.strip()
rows = [{"Linha": l, "url (id do post)": u[l], "Perfil": df.loc[l-2, 'Perfil'], "Prioridade": p, "Problema": pr,
         "Descrição atual": txt[l], "Descrição revisada (preencher)": "", "Observação da autora": ""} for l, p, pr in L]
for r in rows: assert r["Perfil"] == "Imperatriz Online"
out = pd.DataFrame(rows).sort_values(["Prioridade", "Linha"])
notas = pd.DataFrame([
 ["Arquivo", "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx (cópia), com a correção aprovada da linha 1502."],
 ["O que é", "19 descrições do Imperatriz Online marcadas com o código I18 na Solicitação 16 (problemas de redação que limitam a leitura). A planilha original NÃO foi alterada."],
 ["Como devolver", "Preencha a coluna \"Descrição revisada\" (texto completo, como deve ficar) ou escreva \"manter\". As revisões aprovadas entram em correcoes_planilha.py, aplicadas só na cópia."],
 ["Prioridade", "1 Alta = falta conteúdo (imagem, legenda, título ou empresa). 2 Média = frase vazia ou duplicada, sem perda de conteúdo. 3 Baixa = grafia. Decisão = avaliação que depende da autora."],
 ["Escopo", "Só o Imperatriz Online (I18). As descrições do Roma marcadas com R19 (705, 722, 784, 1119, 1162, 975) são outra lista e podem ser enviadas depois."],
 ["Script", "revisao_i18.py"],
], columns=["Item", "Descrição"])
S = "Revisao_I18_descricoes.xlsx"
with pd.ExcelWriter(S, engine="openpyxl") as w:
    out.to_excel(w, sheet_name="Para_revisar", index=False); notas.to_excel(w, sheet_name="Notas", index=False)
wb = load_workbook(S); ws = wb["Para_revisar"]
for c, wd in zip("ABCDEFGH", [8, 42, 18, 11, 45, 90, 60, 30]): ws.column_dimensions[c].width = wd
for row in ws.iter_rows():
    for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
for c in ws[1]: c.font = Font(bold=True)
for r in range(2, ws.max_row + 1): ws.cell(r, 7).fill = PatternFill("solid", fgColor="FFF2CC")
ws.freeze_panes = "A2"
wn = wb["Notas"]; wn.column_dimensions["A"].width = 18; wn.column_dimensions["B"].width = 110
for row in wn.iter_rows():
    for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
wb.save(S); print(out[["Linha", "Prioridade"]].to_string(index=False))
