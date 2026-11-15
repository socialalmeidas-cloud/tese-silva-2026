"""Correções aprovadas pela autora, aplicadas SÓ nos scripts (a planilha congelada não é alterada).
Cada correção: chave (url), coluna, valor atual esperado (para conferência), valor novo, data e origem.
Uso nos scripts: from correcoes_planilha import aplicar; df = aplicar(df)
"""
from privado import carregar
import pandas as pd

CORRECOES = carregar('correcoes_planilha', 'CORRECOES', [])


# Revisões das descrições I18 (Solicitação 16), enviadas pela autora em 05/11/2026.
# Ficam em privado/correcoes_i18.json (fora do repositório), com valor_atual para conferência.
import json as _json, os as _os
from privado import PASTA as _PASTA
_I18 = _os.path.join(_PASTA, "correcoes_i18.json")
if _os.path.exists(_I18):
    with open(_I18, encoding="utf8") as _f:
        CORRECOES += _json.load(_f)


def aplicar(df):
    df = df.copy()
    for c in CORRECOES:
        m = df["url "] == c["url"]
        assert m.sum() == 1, c["url"]
        atual = str(df.loc[m, c["coluna"]].iloc[0])
        assert atual == c["valor_atual"], f"valor atual inesperado em {c['url']}: {atual!r}"
        df.loc[m, c["coluna"]] = c["valor_novo"]
    return df
