import os

import os

# Nome do arquivo a ser verificado (incluindo extensão)
nome_arquivo = "f_Servicenow_combinado.xlsx"

# Verificar se o arquivo existe no diretório atual
if os.path.exists(nome_arquivo):
    print(f"O arquivo '{nome_arquivo}' EXISTE no diretório atual.")
else:
    print(f"O arquivo '{nome_arquivo}' não existe no diretório atual.")

