# Inteligente para Verificação de Referências Bibliográficas

O objetivo deste projeto é automatizar o processo de conferência bibliográfica, atualmente realizado manualmente, o que consome um tempo significativo dos colaboradores. 

## O que o código faz? 

O código compara as referências bibliográficas extraídas da planilha exportada do SAG com o catálogo das bibliotecas digitais disponíveis, como Minha Biblioteca e ProQuest. Caso uma correspondência exata (match) seja identificada, a respectiva linha é automaticamente destacada em verde, indicando que a bibliografia está contemplada nas bases institucionais. Essa automação reduz o esforço manual, aumenta a eficiência e melhora a precisão na validação do acervo bibliográfico virtual.

## Scripts

### 1. `tratamento_bibliografia_SAG.py:`
- **Descrição**: Realiza a limpeza de dados na planilha de Bibliografia, removendo espaços em branco no início e no final das células, além de tratar textos (como remover quebras de linha). O objetivo é padronizar os dados para análise ou integração com outras planilhas.
- **Uso**: Executar o script para limpar dados espaços em branco.

### 2. `tratamento_catalogo.py`
- **Descrição**: Processa a planilha do Catálogo MB, removendo tags HTML, espaços extras e trechos indesejados (ex.: "Ebook"). O código também garante que os valores da coluna "ABNT" estejam padronizados para uso em comparações ou integrações com outras bases de dados.
- **Uso**: Executar o script para padronizar a referência bibliográfica.

### 3. `tratamento_catalogo.py`
- **Descrição**: 
Compara os valores da coluna "BIBLIOGRAFIA" da planilha do curso com os valores da coluna "ABNT" do Catálogo MB, por exemplo. Células com correspondências são destacadas em verde usando PatternFill do openpyxl, facilitando a identificação visual de dados relacionados.
- **Uso**: Executar o script para identificar e destacar visualmente correspondências entre duas planilhas.

## Requisitos

- Python 3.x
- Bibliotecas:
  - `openpyxl`
  - `pandas` 


