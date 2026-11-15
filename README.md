# tese-analise-perfis-ia

Scripts em Python usados na análise da tese **OS PERFIS INFORMATIVOS NO INSTAGRAM EM IMPERATRIZ: um estudo de caso da sua atuação como agentes de circulação de informação em um ambiente midiático local** (Programa de Pós-Graduação em Comunicação, Universidade Federal de Minas Gerais).

O repositório documenta o **procedimento** de análise: como os dados codificados pela Ficha Analítica foram contados, cruzados e testados. **Os dados da pesquisa não estão incluídos.**

## Dados

- A planilha de codificação (4.266 posts dos perfis Ego, Roma e Imperatriz Online) e as descrições das imagens contêm informações pessoais de terceiros (inclusive nomes de vítimas e suspeitos citados nos posts). Por isso, permanecem sob guarda da autora e podem ser disponibilizadas **mediante solicitação justificada**.
- O conteúdo analítico que estava escrito dentro de alguns scripts (trechos literais das descrições usados como evidência, correções caso a caso e textos da autora) foi retirado e fica numa pasta `privado/`, que não é versionada. Sem essa pasta, os scripts usam exemplos fictícios (ver `scripts/privado.py`).
- `dados_exemplo/planilha_ficticia.xlsx` tem as mesmas 65 colunas e o mesmo número de linhas por perfil da planilha real, mas todos os valores são **fictícios**: os rótulos de cada variável foram sorteados de modo uniforme entre as categorias do Livro de Códigos, e legendas (modelos inventados, com emojis, menções, hashtags e links fictícios para testar a limpeza), descrições, links, menções e codificadores foram substituídos por textos genéricos. Ela serve apenas para testar a execução e **não tem relação com o corpus**; os resultados obtidos com ela não têm significado.

## Estrutura

```
scripts/            scripts de análise (uma "solicitação" por etapa)
dados_exemplo/      planilha fictícia e desc_recorte.json fictício
privado/            (não versionada) conteúdo analítico sob guarda da autora
```

## Como executar

```bash
pip install -r requirements.txt
mkdir trabalho && cd trabalho
cp ../dados_exemplo/planilha_ficticia.xlsx copia.xlsx
cp ../dados_exemplo/desc_recorte.json .
python3 -I ../scripts/solicitacao1.py copia.xlsx
```

Cada script lê a planilha indicada e grava suas saídas (`.xlsx`, `.png`) na pasta atual. A planilha original nunca é alterada: as correções aprovadas pela autora são aplicadas em memória por `correcoes_planilha.py`.

## Etapas (solicitações)

Situação ao executar com a planilha fictícia (teste de 15/11/2026):

- **roda**: executa até o fim;
- **conferência**: o script contém verificações (`assert`) com valores do corpus real e para de propósito com dados fictícios;
- **depende de etapa anterior**: precisa da saída de outro script;
- **depende de arquivo privado**: precisa de arquivo mantido fora do repositório.

| Script | Etapa | Situação |
|---|---|---|
| solicitacao1 | Panorama dos perfis | roda |
| solicitacao2 | Agenda temática (Eixo Temático) | roda |
| solicitacao3 | Função da informação (Mediação da Informação) | roda |
| solicitacao4 | Eixo Temático × Mediação, qui-quadrado e V de Cramér | roda |
| solicitacao4b | Eixo × Mediação: qui-quadrado em três versões | conferência |
| solicitacao5 | Regime de Visibilidade | roda |
| solicitacao6 | Regime de Visibilidade × Eixo Temático | roda |
| solicitacao7 | Conteúdo Sensível | roda |
| solicitacao8 | Impacto Visual por Recurso e por Perfil, qui-quadrado | conferência |
| solicitacao9 | Referência Local (variantes de grafia; usa `decisoes_locais.py`) | conferência |
| solicitacao9b | Gráficos de Referência Local | depende da 9 |
| solicitacao9c | Bairro × Eixo Temático | depende da 9 |
| solicitacao10 | Calendário (Período Especial) | roda |
| solicitacao11 | Chamadas para ação (CTA) | roda |
| solicitacao12 | Collab | conferência |
| solicitacao13 | Hashtags e menções | depende de arquivo privado (menções) |
| solicitacao14 | Corpus de legendas para o IRaMuTeQ: relatório prévio (`--etapa relatorio`) e arquivo final (`--etapa arquivo`) | roda |
| solicitacao15 | Validação do recorte da Descrição da Imagem | roda |
| solicitacao16 | Leitura interpretativa das descrições (organização em lotes e conferência das evidências) | roda (com exemplo fictício) |
| solicitacao16_sorteio | Sorteio da amostra-piloto | roda |
| solicitacao17, 17b, 17c | Temática Secundária e cruzamento com Eixo (usa `variantes_aprovadas.py`) | roda |
| solicitacao17_rostos | Escala territorial × Referência Local | depende da 9 |
| solicitacao18 | Eixo × Temática Secundária: associação (lift) e agrupamento | conferência (dados fictícios não geram pares) |
| solicitacao18b | Dispersão (PCoA) a partir da 18 | depende da 18 |
| solicitacao18_ambiente | Menções a outros perfis (ambiente midiático) | depende de arquivo privado (menções) |
| solicitacao19 | Identidade visual dos perfis | roda |
| solicitacao20 | Nuvem de palavras da Linguagem Regional | conferência |
| solicitacao20_mediacao | Eixo × Mediação sem Publicidade | conferência |
| solicitacao21 | A cidade vista pelos perfis | roda (sem anotações, tabelas vazias) |
| solicitacao21_recursos_tipo | Recursos Audiovisuais × Tipo de Conteúdo | roda |
| solicitacao22 | O campo cultural | roda (sem anotações, tabelas vazias) |
| solicitacao23 | O campo social | roda (sem anotações, tabelas vazias) |
| solicitacao24 | Que usuário os posts pressupõem | depende de arquivo privado (anotações) |
| solicitacao25 | O que não aparece nos perfis (usa `nivel3_s25.py`) | depende de arquivo privado (lista de bairros e anotações) |

Módulos auxiliares: `correcoes_planilha.py`, `decisoes_locais.py`, `variantes_aprovadas.py`, `evidencias_s16.py`, `nivel3_s25.py`, `privado.py`. Ferramentas de apoio: `s16_lotes.py` (lotes de leitura), `revisao_i18.py` (lista de descrições a revisar) e `s24c_numeros_roma.py` (conferência dos números da Solicitação 24); as duas últimas dependem de arquivos privados.

A Solicitação 14 não foi aplicada aos dados da tese; o script acompanha o repositório para quem quiser preparar as próprias legendas para o IRaMuTeQ. As versões anteriores dos scripts (substituídas durante a análise) e os geradores dos memorandos interpretativos não estão incluídos, porque reproduzem o texto analítico da tese.

## Uso de inteligência artificial

Em conformidade com o Guia de Inteligência Artificial da UFMG (2026): os scripts foram elaborados com apoio da ferramenta Claude (modelo Claude Opus 5.5, Anthropic), entre abril e novembro de 2026, com conhecimento e anuência do orientador e da coorientadora. A coleta, a codificação do corpus (pela autora e pela equipe de codificadoras), a definição das categorias e a interpretação dos resultados são de responsabilidade da autora. A declaração completa está na seção 4.4 da tese.

## Como citar

> SILVA, Gabriela Almeida. **Os perfis informativos no Instagram em Imperatriz**: um estudo de caso da sua atuação como agentes de circulação de informação em um ambiente midiático local. 2026. Tese (Doutorado em Comunicação) – Universidade Federal de Minas Gerais, Belo Horizonte, 2026.

## Licença

Código sob a licença Apache 2.0 (ver `LICENSE`). A licença cobre só o código; não concede acesso nem direito de uso sobre os dados da pesquisa.
