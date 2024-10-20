import os
import openai
import pandas as pd
from openai.embeddings_utils import get_embedding, cosine_similarity

# Defina sua chave de API da OpenAI
openai.api_key = 'sk-YxOdNbvzbDTPrgoAy'

# Função para carregar a planilha
def carregar_excel(arquivo_excel):
    try:
        print(f"Carregando o arquivo Excel: {arquivo_excel}")
        df = pd.read_excel(arquivo_excel)
        print("Arquivo carregado com sucesso.")
        return df
    except FileNotFoundError:
        print(f"Arquivo {arquivo_excel} não encontrado.")
        exit(1)
    except Exception as e:
        print(f"Ocorreu um erro ao carregar o arquivo: {e}")
        exit(1)

# Função para verificar dados
def verificar_dados(df, coluna):
    nulos = df[coluna].isnull().sum()
    if nulos > 0:
        print(f"Atenção: A coluna '{coluna}' contém {nulos} valores nulos.")
    nao_string = df[~df[coluna].apply(lambda x: isinstance(x, str))]
    if not nao_string.empty:
        print(f"Atenção: A coluna '{coluna}' contém valores não string. Eles serão convertidos para string.")

# Função para calcular embeddings dos textos
def calcular_embeddings(df, coluna):
    print(f"Calculando embeddings para a coluna '{coluna}'...")
    def get_embedding_safe(x):
        try:
            if pd.notnull(x):
                return get_embedding(str(x), engine='text-embedding-ada-002')
            else:
                return None
        except Exception as e:
            print(f"Erro ao calcular o embedding para o valor '{x}': {e}")
            return None
    df[f'{coluna}_embedding'] = df[coluna].apply(get_embedding_safe)
    print(f"Embeddings calculados para a coluna '{coluna}'.")
    return df

def verificar_semelhanca(df_venda, df_compra, threshold=0.85):
    print("Iniciando a verificação de semelhança entre itens de venda e compra...")
    resultados = []

    for idx, row_venda in df_venda.iterrows():
        max_sim = 0
        melhor_correspondencia = None
        codigo_compra = None

        print(f"Processando Item Venda [{idx+1}/{len(df_venda)}]: '{row_venda['Itens Venda']}'")
        venda_embedding = row_venda['Itens Venda_embedding']

        if venda_embedding is None:
            print(f"Embedding do Item Venda '{row_venda['Itens Venda']}' é None. Pulando este item.")
            resultados.append({
                'Item Venda': row_venda['Itens Venda'],
                'Item Compra': None,
                'Código Item Compra': None,
                'Similaridade': None
            })
            continue

        for jdx, row_compra in df_compra.iterrows():
            compra_embedding = row_compra['Itens Compra_embedding']
            if compra_embedding is None:
                continue  # Pula este item de compra se o embedding for None

            sim = cosine_similarity(venda_embedding, compra_embedding)
            if sim > max_sim:
                max_sim = sim
                melhor_correspondencia = row_compra['Itens Compra']
                codigo_compra = row_compra['Código Compra']

        if max_sim >= threshold:
            print(f"Melhor correspondência encontrada: '{melhor_correspondencia}' com similaridade de {max_sim:.4f}")
        else:
            print(f"Nenhuma correspondência acima do threshold para o item '{row_venda['Itens Venda']}'.")

        resultados.append({
            'Item Venda': row_venda['Itens Venda'],
            'Item Compra': melhor_correspondencia,
            'Código Item Compra': codigo_compra,
            'Similaridade': max_sim if melhor_correspondencia else None
        })

    print("Verificação de semelhança concluída.")
    return resultados

# Função para salvar os resultados em um novo Excel
def salvar_resultados(resultados, nome_arquivo):
    print(f"Salvando resultados no arquivo '{nome_arquivo}'...")
    df_resultados = pd.DataFrame(resultados)
    df_resultados.to_excel(nome_arquivo, index=False)
    print("Resultados salvos com sucesso.")

# Fluxo principal
arquivo_excel = 'C:\\Users\\Leandro\\Desktop\\itens_venda_compra.xlsx'
df = carregar_excel(arquivo_excel)

# Separando os dataframes de venda e compra
print("Preparando os dados...")
df_venda = df[['Itens Venda']].drop_duplicates().reset_index(drop=True)
df_compra = df[['Itens Compra', 'Código Compra']].drop_duplicates().reset_index(drop=True)
print("Dados preparados.")

# Verificar dados
verificar_dados(df_venda, 'Itens Venda')
verificar_dados(df_compra, 'Itens Compra')

# Calculando embeddings
df_venda = calcular_embeddings(df_venda, 'Itens Venda')
df_compra = calcular_embeddings(df_compra, 'Itens Compra')

# Verificando semelhanças
resultados = verificar_semelhanca(df_venda, df_compra)

# Salvando os resultados
salvar_resultados(resultados, 'C:\\Users\\Leandro\\Desktop\\resultado_semelhanca.xlsx')

print("Processamento completo. Arquivo 'resultado_semelhanca.xlsx' foi gerado com os resultados.")
