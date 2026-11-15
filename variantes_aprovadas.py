"""Variantes de grafia da "Temática Secundária" aprovadas pela autora em 07/11/2026 (Solicitação 17, aba Variantes_para_aprovacao).
Aplicadas SÓ nos scripts; a planilha original não é alterada.

Decisão da autora (07/11/2026, corrigida no mesmo dia): aprova os blocos A e B e todos os pares de palavras
relacionadas do bloco C, EXCETO três, que ficam separados: Cantor/Cantora (gênero), Campeã/Campeão (gênero) e Feira/Feirão.
Raio/Raio-X: unificados por decisão da autora (registro: na l. 3922 "Raio" é chuva/tempo; na l. 1897 "Raio-X" é o exame).
Pares do bloco D e pares "fora da regra" (Desaparecido/Desaparecimento; Pavimentação/Pavimentação pública): não unificados.
A forma final de cada grupo do bloco C foi escolhida pela assistente e está a confirmar.
"""

MAPA = {
    # Bloco A — singular/plural e grafia (forma final proposta e aprovada)
    "Alimento": "Alimentos",
    "Cantor": "Cantores",              # bloco A (singular/plural). Cantora NÃO entra (decisão da autora)
    "Dica": "Dicas",
    "Equipamento": "Equipamentos",
    "Funcionário": "Funcionários",
    "Material": "Materiais",
    "Matrícula": "Matrículas",
    "Motorista": "Motoristas",
    "Perfume": "Perfumes",
    "Visita": "Visitas",
    "LGBT": "LGBTQIAPN+",
    "Câmara dos vereadores": "Câmara de Vereadores",
    # Bloco B — erros de digitação e maiúsculas
    "Rastramento Veicular": "Rastreamento Veicular",
    "Transplanete Capilar": "Transplante Capilar",
    "Gentiliza": "Gentileza",
    "Corpus Christ": "Corpus Christi",
    "Arraia": "Arraial",               # l. 217: Arraial de Imperatriz (conferido na legenda)
    "MóVéis": "Móveis",
    "área de vegetação": "Área de vegetação",
    "coleta": "Coleta",
    "semifinal": "Semifinal",
    # Bloco C — palavras relacionadas (forma final escolhida pela assistente; a confirmar)
    "Policial": "Polícia",
    "Pizza": "Pizzaria",
    "Humoristas": "Humor",
    "Raio": "Raio / Raio-X",           # unificados por decisão da autora (ver docstring)
    "Raio-X": "Raio / Raio-X",
}
DATA_DECISAO = "07/11/2026"


def normalizar(rotulo):
    """Devolve a forma unificada do rótulo (ou o próprio rótulo, se não houver variante aprovada)."""
    return MAPA.get(rotulo, rotulo)


def rotulos_do_post(celula, sem="Não há"):
    """Separa a célula por ';', tira espaços, aplica as variantes aprovadas e remove repetições no mesmo post
    (ex.: um post com 'Cantor; Cantora' passa a ter 'Cantores' uma vez só)."""
    out = []
    for x in str(celula).split(";"):
        x = x.strip()
        if x and x != sem:
            x = normalizar(x)
            if x not in out:
                out.append(x)
    return out


def tabela_correspondencia():
    import pandas as pd
    return pd.DataFrame([{"Grafia na planilha": k, "Forma unificada": v} for k, v in MAPA.items()])
