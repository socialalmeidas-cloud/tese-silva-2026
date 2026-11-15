"""Nível 3 da Solicitação 25 (importado por solicitacao25.py). Tudo aqui é INFERÊNCIA dependente da fonte."""
from privado import carregar
import re, collections, unicodedata
import pandas as pd

def norm(s):
    s = ''.join(c for c in unicodedata.normalize('NFD', str(s).lower()) if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s).strip()

FONTES = [
    {"Referência": "Lista de bairros de Imperatriz", "Fonte": "CEP Brasil — cepbrasil.org/maranhao/imperatriz (acesso em 09/11/2026)",
     "Tipo": "NÃO OFICIAL (site de CEP); 114 nomes", "Uso": "Rodado: comparação com a Referência Local (todos os posts) e com as descrições do recorte",
     "Conferir": "Substituir por lista oficial (Prefeitura/Lei de bairros ou setores censitários do IBGE). A lista não traz bairros codificados no banco (ver aba Nivel3_locais_fora_lista)."},
    {"Referência": "Religião em Imperatriz — Censo 2022", "Fonte": "IBGE, Censo 2022 (Religiões, resultados preliminares da amostra), via Imperatriz Notícias/UFMA, 31/10/2025 (Elenilson Silva)",
     "Tipo": "Dado oficial por FONTE SECUNDÁRIA; tabela SIDRA não citada", "Uso": "Rodado: comparação com menções religiosas nas descrições do recorte",
     "Conferir": "Conferir os percentuais no SIDRA/IBGE para o código 2105302 antes de citar na tese."},
    {"Referência": "Cor ou raça — Censo 2022", "Fonte": "IBGE, Censo 2022, via O Imparcial, 22/12/2023 (Maranhão: 66% pardos)",
     "Tipo": "Dado ESTADUAL por fonte secundária; dado municipal não obtido", "Uso": "Rodado só como limite: as descrições não registram cor/raça de modo sistemático",
     "Conferir": "Dado municipal: SIDRA, tabela 9605 (Imperatriz, 2105302) — não acessado nesta sessão."},
    {"Referência": "Povos indígenas e quilombolas em Imperatriz", "Fonte": "—", "Tipo": "Não obtida",
     "Uso": "NÃO RODADO no nível 3 (só FATO do banco)", "Conferir": "IBGE Censo 2022, indígenas e quilombolas por município."},
    {"Referência": "População LGBT em Imperatriz", "Fonte": "—", "Tipo": "Não há referência municipal conhecida pela assistente",
     "Uso": "NÃO RODADO no nível 3 (só FATO do banco)", "Conferir": "Indicar fonte, se houver (ex.: dados de eventos ou de coletivos locais)."},
    {"Referência": "Gestores públicos em cena", "Fonte": "—", "Tipo": "Referência seria a agenda do Executivo no período do corpus",
     "Uso": "NÃO RODADO no nível 3 (só FATO do banco)", "Conferir": "Diário Oficial/agenda da Prefeitura no período da coleta, se a autora quiser rodar."},
]

RELIGIAO_IBGE = [("Católica", "50,31%"), ("Evangélica", "37,21%"), ("Espírita", "0,38%"), ("Umbanda e candomblé", "0,2%"),
                 ("Tradições indígenas", "0,01%"), ("Outras religiosidades", "3,37%"), ("Sem religião", "8,45%"), ("Não declarada", "0,08%")]

# Variantes da Referência Local aceitas como o mesmo bairro (a conferir pela autora)
VARIANTES = {
    "Parque Alvorada": ["Parque Alvorada I", "Parque Alvorada 1"], "Parque Buriti": ["Parque do Buriti"],
    "Parque Anhangüera": ["Parque Anhanguera II"], "Residencial Teotônio Vilela": ["Teotônio Vilela"],
    "Vila Airton Senna": ["Parque Ayrton Senna", "Ayrton Senna"], "Vila Redenção": ["Vila Redenção II", "Vila Redenção 2", "Redenção"],
    "Vila Parati": ["Vila Paraty"], "Conjunto Habitar Brasil": ["Habitar Brasil", "Habitar Brasil II"],
    "Residencial Dom Afonso Felippe Gregory": ["Dom Afonso Felipe"], "Jardim Cinco Estrelas": ["Cinco Estrelas"],
    "Vila Maderminas": ["Vila Madermina"], "Parque Santa Lúcia": ["Santa Lúcia"], "Vila Cafeteira": ["Cafeteira"],
    "Caema": ["Bairro da CAEMA", "Bairro CAEMA"], "Bom Sucesso": ["Bairro Bom Sucesso", "Entroncamento e Bom Sucesso"],
    "Entroncamento": ["Entroncamento e Bom Sucesso"], "Santa Rita": ["Bairro Santa Rita", "Grande Santa Rita"],
    "Residencial Colina Park": ["Colinas Park"], "Park Imperial": ["Parque Imperial"], "Jardim Morada do Sol": ["Morada do Sol"],
    "Alto Bonito do Triângulo": ["Alto Bonito"], "Residencial Kubitschek": ["Residencial Juscelino Kubitschek"],
    "Alto da Boa Vista": ["bairro Alto da Boa Vista"], "Leandra": ["Bairro Leandra", "Vila Leandra"], "Mutirão": ["Bairro Mutirão"],
    "Parque Sanharol": ["Sanharol"], "Parque das Palmeiras": ["Park das Palmeiras", "Rua da Paz. Parque das Palmeiras"],
    "Parque Independência": ["Independência"], "Juçara": ["Jussara"], "Centro": ["Centro Imperatriz"], "Beira Rio": ["Setor Beira-Rio"],
    "Nova Imperatriz": ["rua Leão XIIINova Imperatriz"], "Parque do Bosque": ["Parque dos Bosques"], "Conjunto Nova Vitória": ["Conjunto Vitória"],
    "Bom Jesus": ["Grande Bom Jesus"], "União": ["Vila União"],
}
HOMONIMO = {"Santa Inês", "Santa Rita", "Santa Luzia", "Centro Novo", "Camaçari", "Asa Norte", "São Salvador", "Bom Jesus"}
NAO_BAIRRO = {"Vila Nova dos Martírios", "Parque de Exposições"}   # município e equipamento, não bairro
PREFIXO_BAIRRO = r"^(Vila|Parque|Park|Residencial|Conjunto|Jardim|Bairro|Cidade Nova|Setor|Grande )"

# Menções religiosas: linhas lidas e classificadas (falsos positivos excluídos: ruas Frei Damião/Frei Caneca, "indigente", "App Pardal" etc.)
RELIG = carregar('nivel3_s25', 'RELIG', [(2, "Categoria fictícia", "Descrição fictícia")])
COR = carregar('nivel3_s25', 'COR', [])
SO_FATO = carregar('nivel3_s25', 'SO_FATO', [])


def rodar(df, DESC, PERFIS, rotulos, perfil_da_linha, faixa):
    ref = [l.strip() for l in open("ref_bairros_cepbrasil.txt", encoding="utf8") if l.strip()]
    assert len(ref) == 114
    alvo = {}                                   # valor normalizado da Referência Local -> bairro da lista
    for b in ref:
        alvo[norm(b)] = b
        for v in VARIANTES.get(b, []):
            alvo.setdefault(norm(v), b)
    cont = {p: collections.Counter() for p in PERFIS}
    nao_casados = collections.Counter()
    for _, r in df.iterrows():
        for t in rotulos(r["Referência Local"]):
            b = alvo.get(norm(t))
            if b:
                cont[r.Perfil][b] += 1
            elif re.search(PREFIXO_BAIRRO, t) and t not in NAO_BAIRRO:
                nao_casados[(r.Perfil, t)] += 1
    # menção nas descrições do recorte (nome completo, sem acento; homônimos marcados)
    descn = {l: norm(d) for l, d in DESC.items()}
    linhas = []
    for b in ref:
        rx = r"\b" + re.escape(norm(b)) + r"\b"
        md = collections.Counter(perfil_da_linha(l) for l, d in descn.items() if re.search(rx, d))
        ns = {p: cont[p][b] for p in PERFIS}
        tot = sum(ns.values())
        sit = ("Não observado em nenhum perfil" if tot == 0 and sum(md.values()) == 0 else
               "Só nas descrições" if tot == 0 else
               "Observado nos três" if all(ns.values()) else
               "Observado só em " + ", ".join(p for p in PERFIS if ns[p]))
        linhas.append({"Bairro (lista de referência)": b, "Variantes aceitas na Referência Local": "; ".join(VARIANTES.get(b, [])),
                       "Homônimo possível": "sim" if b in HOMONIMO else "",
                       **{f"{p} — posts (Referência Local)": ns[p] for p in PERFIS},
                       **{f"{p} — descrições do recorte que citam": md.get(p, 0) for p in PERFIS},
                       "Situação": sit})
    t_bai = pd.DataFrame(linhas)
    t_fora = pd.DataFrame([{"Perfil": p, "Valor da Referência Local (forma de bairro, fora da lista)": v, "N (posts do perfil)": n,
                            "Marca": "*" if n < 10 else ""} for (p, v), n in sorted(nao_casados.items(), key=lambda x: (x[0][0], -x[1]))])
    # religião
    for l, c, t in RELIG + [(l, "", t) for l, t in COR] + SO_FATO:
        assert t in DESC[str(l)], (l, t)
    t_rel = pd.DataFrame([{"Linha": l, "Perfil": perfil_da_linha(l), "Categoria (leitura)": c, "Trecho": t} for l, c, t in RELIG])
    resumo = t_rel.groupby(["Categoria (leitura)", "Perfil"]).size().unstack(fill_value=0).reindex(columns=PERFIS, fill_value=0).reset_index()
    t_ibge = pd.DataFrame(RELIGIAO_IBGE, columns=["Religião (IBGE, Censo 2022, Imperatriz — via fonte secundária)", "Percentual"])
    return t_bai, t_fora, t_rel, resumo, t_ibge, pd.DataFrame(FONTES)


def ref_ev(t_bai, bairro):
    """Evidência de referência externa: não há trecho de ausência; registra-se a linha da lista e as contagens."""
    r = t_bai[t_bai["Bairro (lista de referência)"] == bairro].iloc[0]
    ns = "; ".join(f"{p.split()[0] if p != 'Imperatriz Online' else 'IO'} {r[c]}" for p, c in
                   (("Ego", "Ego — posts (Referência Local)"), ("Roma", "Roma — posts (Referência Local)"),
                    ("Imperatriz Online", "Imperatriz Online — posts (Referência Local)")))
    return ["—", f"Lista de referência: \"{bairro}\" — posts com o bairro na Referência Local: {ns}"]


def padroes_n3(t_bai):
    """Padrões interpretativos do Nível 3. O texto analítico e as evidências (trechos literais)
    ficam em privado/nivel3_padroes.py; sem ele, devolve um padrão fictício de exemplo."""
    import sys
    from privado import PASTA
    if PASTA not in sys.path:
        sys.path.insert(0, PASTA)
    try:
        from nivel3_padroes import padroes_n3 as _p
    except ImportError:
        return [{"perfil": "Ego", "cod": "N3-E1", "nome": "Padrão fictício", "linhas_n": 0,
                 "def": "Definição fictícia.", "ev": [[2, "Descrição fictícia"]],
                 "contra": [2, "Descrição fictícia"], "alt": "Leitura alternativa fictícia."}]
    return _p(t_bai)
