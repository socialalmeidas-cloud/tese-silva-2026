"""Solicitação 1 — Panorama dos perfis.
Lê a cópia da planilha congelada e gera SolicitacaoN_Panorama.xlsx + .png.
Uso: python3 solicitacao1.py <planilha.xlsx>
"""
import sys
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

ARQ = sys.argv[1] if len(sys.argv) > 1 else "copia.xlsx"
NOME_ORIGINAL = "Planilha_Normalizada_FINAL-GABRIELAALMEIDA-04_11_2026.xlsx"
SAIDA = "Solicitacao1_Panorama.xlsx"
PNG = "Solicitacao1_Panorama.png"
PERFIS = ["Ego", "Roma", "Imperatriz Online"]
EMERGENTES = {"Vídeo sem som", "Desenho"}
MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
         "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}

df = pd.read_excel(ARQ)
assert df.shape == (4266, 65), df.shape
N = df["Perfil"].value_counts()
assert N.to_dict() == {"Imperatriz Online": 2721, "Roma": 1131, "Ego": 414}

# ---- Período: formato "(mmm-aaaa)" -> data
p = df["Período"].astype(str).str.strip().str.strip("()")
partes = p.str.split("-", expand=True)
df["mes_num"] = partes[0].str.lower().map(MESES)
df["ano"] = partes[1].astype(int)
assert df["mes_num"].notna().all()
df["data"] = pd.to_datetime(dict(year=df["ano"], month=df["mes_num"], day=1))

# ---- Tabela 1: posts por mês
t1 = (df.pivot_table(index="data", columns="Perfil", values="url ",
                     aggfunc="count", fill_value=0)[PERFIS].sort_index())
t1["Total"] = t1.sum(axis=1)
inv = {v: k for k, v in MESES.items()}
t1.index = [f"{inv[d.month]}/{d.year}" + (" *" if t1.loc[d, "Total"] < 20 else "")
            for d in t1.index]
t1.loc["Total"] = t1.sum()
t1 = t1.reset_index().rename(columns={"index": "Mês/Ano"})
t1.columns = ["Mês/Ano"] + [f"{c} (N)" for c in PERFIS] + ["Total (N)"]


def distribuicao(col, marcar=set()):
    ct = pd.crosstab(df[col], df["Perfil"])[PERFIS]
    ct["Total"] = ct.sum(axis=1)
    ct = ct.sort_values("Total", ascending=False)
    out = pd.DataFrame(index=ct.index)
    for c in PERFIS:
        out[f"{c} (N)"] = ct[c]
        out[f"{c} (%)"] = ct[c] / N[c]
    out["Total (N)"] = ct["Total"]
    out["Total (%)"] = ct["Total"] / len(df)
    rotulos = []
    for r in out.index:
        s = str(r) + (" †" if r in marcar else "")
        if out.loc[r, "Total (N)"] < 10:
            s += " *"
        rotulos.append(s)
    out.index = rotulos
    tot = {f"{c} (N)": N[c] for c in PERFIS} | {f"{c} (%)": 1.0 for c in PERFIS}
    tot |= {"Total (N)": len(df), "Total (%)": 1.0}
    out.loc["Total"] = pd.Series(tot)
    return out.reset_index().rename(columns={"index": "Tipo" if col == "Tipo de Conteúdo" else "Recurso"})


t2 = distribuicao("Tipo de Conteúdo")
t3 = distribuicao("Recursos Audiovisuais", EMERGENTES)

# ---- Conferências
for t in (t1,):
    assert t.iloc[-1, 1:].tolist() == [414, 1131, 2721, 4266]
for t in (t2, t3):
    corpo = t.iloc[:-1]
    for c in PERFIS + ["Total"]:
        assert corpo[f"{c} (N)"].sum() == t.iloc[-1][f"{c} (N)"]
        assert abs(corpo[f"{c} (%)"].sum() - 1) < 1e-9

notas = [
    ["Arquivo", NOME_ORIGINAL + " (versão congelada de 04/11/2026; trabalhado em cópia)"],
    ["N analisado", "4.266 posts (Ego 414; Roma 1.131; Imperatriz Online 2.721)"],
    ["Exclusões", "Nenhuma. A Descrição da Imagem não foi usada nesta solicitação."],
    ["Período", "Formato encontrado: texto \"(mmm-aaaa)\", ex.: \"(jan-2024)\". Convertido para data e ordenado cronologicamente (jan/2024 a dez/2024)."],
    ["Fórmula do %", "% = N do rótulo no perfil ÷ total de posts do perfil (Ego 414; Roma 1.131; IO 2.721; Total 4.266) × 100."],
    ["Rótulo único", "Tipo de Conteúdo e Recursos Audiovisuais têm um rótulo por post; os percentuais de cada coluna somam 100%."],
    ["Nomenclatura", "Tipo de Conteúdo mantém a nomenclatura do programa de coleta: Vídeo, Image (imagem única), Sidecar (carrossel)."],
    ["* em Posts_por_mes", "Mês com menos de 20 posts no total (nenhum caso nesta tabela)."],
    ["* nas demais abas", "Rótulo com menos de 10 posts no total."],
    ["† Categorias emergentes", "Vídeo sem som †, Desenho † (surgiram por \"Outro (especificar)\" durante a codificação; aplicadas apenas a partir do momento em que surgiram). Resultados indicativos; diferenças entre perfis ou codificadoras podem refletir o momento de aplicação da categoria, não diferenças reais entre os perfis."],
    ["Lista completa de emergentes", "Criminoso, Animal, Bens materiais, Candidatos, Outro perfil, Procedimento, Animal morto, Restrição de idade, Suicídio, Nudez, Vídeo sem som, Desenho, Pix (só as duas de Recursos aparecem aqui)."],
    ["Script", "solicitacao1.py"],
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    t1.to_excel(w, "Posts_por_mes", index=False)
    t2.to_excel(w, "Tipo_de_Conteudo", index=False)
    t3.to_excel(w, "Recursos_Audiovisuais", index=False)
    pd.DataFrame(notas, columns=["Item", "Descrição"]).to_excel(w, "Notas", index=False)

wb = load_workbook(SAIDA)
for ws in wb.worksheets:
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col in ws.columns:
        head = str(col[0].value)
        for c in col[1:]:
            if "(%)" in head:
                c.number_format = "0.0%"
            elif "(N)" in head:
                c.number_format = "#,##0"
        largura = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(12, largura + 2), 90)
    if ws.title != "Notas":
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
ws = wb["Notas"]
for row in ws.iter_rows(min_row=2):
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(SAIDA)

# ---- Gráfico (Okabe-Ito; linhas e marcadores diferentes)
g = t1.iloc[:-1]
x = range(len(g))
estilo = {"Ego": ("#E69F00", "o", "-"), "Roma": ("#0072B2", "s", "--"),
          "Imperatriz Online": ("#009E73", "^", "-.")}
fig, ax = plt.subplots(figsize=(10, 5.5))
for perf in PERFIS:
    cor, mk, ls = estilo[perf]
    y = g[f"{perf} (N)"].tolist()
    ax.plot(x, y, color=cor, marker=mk, linestyle=ls, linewidth=2, label=perf)
    for xi, yi in zip(x, y):
        # se outra linha estiver próxima (< 20 posts) e acima, rótulo vai abaixo do ponto
        outros = [g[f"{o} (N)"].iloc[xi] for o in PERFIS if o != perf]
        abaixo = any(0 < (v - yi) < 20 for v in outros)
        ax.annotate(f"{yi:,}".replace(",", "."), (xi, yi), textcoords="offset points",
                    xytext=(0, -14 if abaixo else 7), ha="center", fontsize=8, color=cor)
ax.set_xticks(list(x))
ax.set_xticklabels([s.replace(" *", "") for s in g["Mês/Ano"]], rotation=45, ha="right")
ax.set_xlabel("Mês/Ano")
ax.set_ylabel("Número de posts")
ax.set_ylim(0, g[[f"{p} (N)" for p in PERFIS]].values.max() * 1.15)
ax.legend(frameon=False, loc="upper right", ncol=3)
ax.spines[["top", "right"]].set_visible(False)
ax.grid(axis="y", alpha=0.3)
fig.tight_layout()
fig.savefig(PNG, dpi=300)
print("ok")
