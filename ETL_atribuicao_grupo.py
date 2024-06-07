import pandas as pd
import os

# Caminho do arquivo de entrada
caminho_arquivo_entrada = r"Z:\#BASES DE APOIO - GERENCIAL\01 - BASES\PROJETOS\Controle Contabil\REQUISITANTE\LISTA REQUISITANTES.xlsx"

# Caminho do arquivo de saída
caminho_arquivo_saida = r"C:\Users\2160011883\Via Varejo S.A\DB_MANUTENCAO - Bases P-Pbi\d_atribuição_Grupo.xlsx"

# Verificar se o arquivo de saída já existe
if os.path.exists(caminho_arquivo_saida):
    os.remove(caminho_arquivo_saida)  # Remover o arquivo existente

# Carregar o arquivo Excel e selecionar as colunas desejadas
try:
    df = pd.read_excel(caminho_arquivo_entrada, usecols=["NOME", "FUNÇÃO",'MATRÍCULA'])
except FileNotFoundError:
    print(f"Arquivo '{caminho_arquivo_entrada}' não encontrado.")
    exit(1)
except Exception as e:
    print(f"Erro ao carregar arquivo: {e}")
    exit(1)

# Verificar se o DataFrame foi carregado com sucesso
if df.empty:
    print("O DataFrame está vazio. Verifique se as colunas especificadas existem no arquivo.")
    exit(1)

# Salvar o DataFrame resultante como um novo arquivo Excel
try:
    df.to_excel(caminho_arquivo_saida, index=False)
    print(f"Arquivo 'd_atribuição_Grupo.xlsx' salvo com sucesso em '{caminho_arquivo_saida}'")
except Exception as e:
    print(f"Erro ao salvar arquivo: {e}")
    exit(1)
