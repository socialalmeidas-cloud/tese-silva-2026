"""Solicitação 16 — Etapa 2: piloto (60 posts) e mini-livro v0.1.
Sorteio: 20 posts por perfil entre os 1.605 descritos, semente 20261105 (pandas .sample, random_state).
Gera Piloto_para_classificar.xlsx (para a pesquisadora; sem a classificação da assistente).
A classificação da assistente fica em outro script, entregue junto com Piloto_classificacao_Claude.xlsx
depois que a pesquisadora devolver o piloto preenchido (classificação cega).
Uso: python3 solicitacao16_etapa2.py <planilha.xlsx>
"""
import sys
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from correcoes_planilha import aplicar

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
SEMENTE = 20261105
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
VERSAO = "mini-livro v0.1 (05/11/2026)"

SUJEITO = ["Pessoa comum", "Ocorrência e envolvidos", "Personalidade pública", "Profissional ou empresário promovido",
           "O próprio perfil ou seu apresentador", "Produto, marca ou estabelecimento", "Autoridade, instituição ou agente público",
           "Espaço da cidade ou fenômeno", "Outro (especificar)", "Não se aplica"]
POSICAO = ["Testemunhal/noticiosa", "Institucional/utilidade pública", "Denúncia/cobrança", "Comercial", "Afetiva",
           "Lúdica (humor/entretenimento)", "Outro (especificar)", "Não se aplica"]

df = aplicar(pd.read_excel(ARQ))
assert df.shape == (4266, 65)
df["Linha"] = df.index + 2
rec = df[df["Descrição da Imagem"].astype(str).str.strip() != "Não analisado"]
assert len(rec) == 1605
amostra = pd.concat([rec[rec["Perfil"] == p].sample(20, random_state=SEMENTE) for p in PERFIS])
assert len(amostra) == 60 and amostra["Linha"].is_unique

base = amostra[["Linha", "Perfil", "Eixo Temático", "Descrição da Imagem"]].rename(columns={"Descrição da Imagem": "Descrição"})

# (1) arquivo para a pesquisadora
p1 = base.copy()
p1["Sujeito da imagem — classificação da pesquisadora"] = ""
p1["Posição do perfil — classificação da pesquisadora"] = ""
p1["Se \"Outro (especificar)\": qual?"] = ""
p1["Observação"] = ""
instr = pd.DataFrame([
    ["Versão", VERSAO],
    ["Amostra", "60 posts (20 por perfil) sorteados entre os 1.605 do recorte; semente 20261105."],
    ["Como preencher", "Escolha UMA categoria por variável na lista suspensa (a predominante, segundo a descrição). Use \"Outro (especificar)\" e escreva qual; use \"Não se aplica\" quando a pergunta não couber."],
    ["Definições", "Ver o arquivo Solicitacao16_mini_livro_v0.1.md."],
    ["Linha", "Número da linha na planilha original (linha 1 = cabeçalho)."],
], columns=["Item", "Descrição"])

for arq, t, nome in (("Piloto_para_classificar.xlsx", p1, "Piloto"),):
    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        t.to_excel(w, sheet_name=nome, index=False)
        pd.DataFrame({"Sujeito da imagem": pd.Series(SUJEITO), "Posição do perfil": pd.Series(POSICAO)}).to_excel(w, sheet_name="Categorias", index=False)
        instr.to_excel(w, sheet_name="Instrucoes", index=False)
    wb = load_workbook(arq)
    ws = wb[nome]
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    larg = {"A": 8, "B": 16, "C": 18, "D": 90, "E": 30, "F": 30, "G": 45, "H": 40}
    for L, wd in larg.items():
        ws.column_dimensions[L].width = wd
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    if nome == "Piloto":
        amarelo = PatternFill("solid", fgColor="FFF2CC")
        dv1 = DataValidation(type="list", formula1=f"=Categorias!$A$2:$A${len(SUJEITO) + 1}", allow_blank=True)
        dv2 = DataValidation(type="list", formula1=f"=Categorias!$B$2:$B${len(POSICAO) + 1}", allow_blank=True)
        ws.add_data_validation(dv1)
        ws.add_data_validation(dv2)
        dv1.add(f"E2:E{ws.max_row}")
        dv2.add(f"F2:F{ws.max_row}")
        for row in ws.iter_rows(min_row=2, min_col=5, max_col=8):
            for c in row:
                c.fill = amarelo
    ws.freeze_panes = "E2"
    for s in ("Categorias", "Instrucoes"):
        w2 = wb[s]
        for c in w2[1]:
            c.font = Font(bold=True)
        w2.column_dimensions["A"].width = 42
        w2.column_dimensions["B"].width = 100 if s == "Instrucoes" else 42
    wb.save(arq)

print(amostra["Perfil"].value_counts().to_dict(), sorted(amostra["Linha"]))
