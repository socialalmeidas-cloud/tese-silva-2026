"""Registro de leitura da Solicitação 16 (Descrição da Imagem).
Cada evidência: linha da planilha, perfil, código do padrão e trecho LITERAL da descrição.
O script solicitacao16.py confere que cada trecho existe exatamente na descrição da linha indicada.
Leitura feita por IA (Claude) a partir das descrições escritas pela pesquisadora; deve ser conferida por amostragem.
"""

# Padrões e evidências da leitura ficam em privado/evidencias_s16.json (fora do repositório).
# Formato: PADROES = {perfil: {código: nome}}; EVIDENCIAS = [(linha, perfil, código, trecho literal), ...]
from privado import carregar

PADROES = carregar("evidencias_s16", "PADROES", {"Ego": {"E01": "Padrão fictício de exemplo"}})
EVIDENCIAS = [tuple(e) for e in carregar("evidencias_s16", "EVIDENCIAS", [(2, "Ego", "E01", "Descrição fictícia")])]
