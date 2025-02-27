# OBJETIVO

O objetivo deste projeto é automatizar o processo de conferência bibliográfica, atualmente realizado manualmente, o que consome um tempo significativo dos colaboradores. O sistema compara as referências bibliográficas extraídas da planilha exportada do SAG com o catálogo das bibliotecas digitais disponíveis, como Minha Biblioteca e ProQuest. Caso uma correspondência exata (match) seja identificada, a respectiva linha é automaticamente destacada em verde, indicando que a bibliografia está contemplada nas bases institucionais. Essa automação reduz o esforço manual, aumenta a eficiência e melhora a precisão na validação do acervo bibliográfico virtual.

## Scripts

### 1. `remove_leading_trailing_spaces.py`
- **Descrição**: Remove espaços em branco no início e no final das células de uma coluna específica em uma planilha Excel, mantendo a formatação original.
- **Uso**: Executar o script para limpar dados em uma coluna específica.

### 2. `highlight_matching_cells.py`
- **Descrição**: Destaca células em uma planilha Excel que contêm valores correspondentes a uma coluna específica de outra planilha, aplicando um preenchimento verde.
- **Uso**: Executar o script para identificar e destacar visualmente correspondências entre duas planilhas.

## Requisitos

- Python 3.x
- Bibliotecas:
  - `openpyxl`
  - `pandas` (apenas para o primeiro código, se necessário)

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/nome-do-repositorio.git
