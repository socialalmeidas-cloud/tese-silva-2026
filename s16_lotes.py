# Etapa 1 da Solicitação 16 — divide o recorte em lotes e exporta os textos para leitura.
# Uso: python3 s16_lotes.py copia.xlsx <Perfil> <nº do lote | todos>
import sys, os, json
import pandas as pd
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from correcoes_planilha import aplicar

ARQ, PERFIL, QUAL = sys.argv[1], sys.argv[2], sys.argv[3]
df = aplicar(pd.read_excel(ARQ))
df["Linha"] = df.index + 2
rec = df[df["Descrição da Imagem"].astype(str).str.strip() != "Não analisado"]
b = rec[rec["Perfil"] == PERFIL].sort_values("Linha")
lotes = [b.iloc[i:i + 60] for i in range(0, len(b), 60)]
print(PERFIL, len(b), "posts em", len(lotes), "lotes:", [len(l) for l in lotes])
sel = range(len(lotes)) if QUAL == "todos" else [int(QUAL) - 1]
for k in sel:
    L = lotes[k]
    print(f"\n===== LOTE {k+1} — linhas {L['Linha'].min()} a {L['Linha'].max()} (N={len(L)}) =====")
    for _, r in L.iterrows():
        print(f"[{r['Linha']}] {r['Descrição da Imagem']}")
