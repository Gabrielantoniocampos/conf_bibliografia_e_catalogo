### PADRONIZAR A COLUNA DE REFERÊNCIA BIBLIOGRÁFICA CONFORME A ABNT, REMOVENDO AS STRINGS DESNECESSÁRIAS

import pandas as pd

# Carregar a planilha Excel para um DataFrame
df = pd.read_excel("/content/Catálogo MB (1).xlsx")

# Função para tratar a coluna "ABNT"
def tratar_abnt(texto):
    # Garantir que o texto seja uma string antes de realizar substituições
    texto = str(texto).replace("<p>", "")  # Remove a tag <p> do texto
    
    # Verificar se a palavra "Ebook" está presente no texto
    if "Ebook" in texto:
        # Remover tudo a partir da palavra "Ebook"
        texto = texto.split("Ebook")[0]
    
    # Retornar o texto tratado, removendo espaços em branco no início e no final
    return texto.strip()

# Aplicar a função `tratar_abnt` a todos os elementos da coluna "ABNT"
df["ABNT"] = df["ABNT"].apply(tratar_abnt)

# Salvar o DataFrame tratado em um novo arquivo Excel
df.to_excel("Catálogo MB Tratado.xlsx", index=False)
