# Projeto de Tratamento e Destaque de Planilhas Excel

Este projeto contém scripts em Python para tratamento e destaque de dados em planilhas Excel, utilizando a biblioteca `openpyxl`. Os scripts foram desenvolvidos para manipular arquivos Excel mantendo a formatação original e realizar operações como remoção de espaços em branco e destaque de células correspondentes.

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
