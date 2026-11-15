"""Carrega o conteúdo analítico que fica FORA do repositório.

Trechos literais das descrições, evidências, correções caso a caso e anotações
contêm dados do corpus (inclusive nomes de terceiros) e ficam sob guarda da autora,
na pasta privado/ (ou na pasta indicada pela variável de ambiente TESE_PRIVADO).
Sem essa pasta, os scripts usam os EXEMPLOS fictícios definidos em cada módulo.
"""
import json, os

PASTA = os.environ.get("TESE_PRIVADO", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "privado"))


def carregar(arquivo, nome, exemplo):
    p = os.path.join(PASTA, arquivo + ".json")
    if os.path.exists(p):
        with open(p, encoding="utf8") as f:
            return json.load(f)[nome]
    return exemplo
