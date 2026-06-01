import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

# Configurações
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

# Caminhos dos arquivos
custo_path = r"C:\Users\win11\Downloads\Custos de produtos - Maio.xlsx"
margem_path = r"C:\Users\win11\Downloads\260523_MRG - wapp.xlsx"
output_path = str(Path.home() / "Downloads" / "Averiguar_Custos (MAR x CUS).xlsx")

def load_data(file_path, file_type):
    try:
        if file_type == 'custo':
            # Tentar ler a aba "Base" (pode ser a única ou uma entre várias)
            df = pd.read_excel(file_path, sheet_name="Base", header=0)
            df = df[['DATA', 'PRODUTO', 'DESCRICAO', 'CUSTO']].copy()
            # Converter para string removendo .0 de números e espaços
            df['PRODUTO'] = df['PRODUTO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True).dt.date
            df['CUSTO'] = pd.to_numeric(
                df['CUSTO'].astype(str).str.replace(',', '.'), 
                errors='coerce'
            )
            df = df.rename(columns={'CUSTO': 'CUSTO_JULHO', 'DESCRICAO': 'DESCRICAO_JULHO'})
            
        elif file_type == 'margem':
            # Ler a aba FEC_PQ com cabeçalho na linha 10 (índice 9)
            df = pd.read_excel(margem_path, sheet_name="FEC_PQ", header=9)
            
            # Filtrar apenas CF='ESP' (ou outros valores relevantes)
            # Verificar se a coluna 'CF' existe
            if 'CF' in df.columns:
                df = df[df['CF'] == 'ESP'].copy()
            else:
                print("ATENÇÃO: Coluna 'CF' não encontrada em FEC_PQ")
                # Tentar encontrar coluna similar
                colunas_cf = [col for col in df.columns if 'CF' in col.upper()]
                if colunas_cf:
                    print(f"Usando coluna alternativa: {colunas_cf[0]}")
                    df = df[df[colunas_cf[0]] == 'ESP'].copy()
            
            # Converter para string removendo .0 de números e espaços
            # Verificar qual coluna de produto está disponível
            if 'CODPRODUTO' in df.columns:
                df['CODPRODUTO'] = df['CODPRODUTO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                col_produto = 'CODPRODUTO'
            elif 'PRODUTO' in df.columns:
                df['PRODUTO'] = df['PRODUTO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                col_produto = 'PRODUTO'
            else:
                # Tentar encontrar coluna de produto por nome similar
                col_produto_candidates = [col for col in df.columns if 'PRODUTO' in col.upper() or 'COD' in col.upper()]
                if col_produto_candidates:
                    col_produto = col_produto_candidates[0]
                    print(f"Usando coluna de produto alternativa: {col_produto}")
                    df[col_produto] = df[col_produto].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                else:
                    raise ValueError("Nenhuma coluna de produto encontrada em FEC_PQ")
            
            # Selecionar colunas necessárias (adaptado para FEC_PQ)
            colunas_disponiveis = df.columns.tolist()
            
            # Mapeamento de colunas
            colunas_necessarias = {
                'PRODUTO': col_produto,
                'DATA': 'DATA',
                'DESCRICAO': 'DESCRICAO',
                'CUSTO': 'CUSTO',
                'CF': 'CF'
            }
            
            # Verificar se as colunas existem e criar mapping real
            colunas_finais = []
            for novo_nome, nome_original in colunas_necessarias.items():
                if nome_original in df.columns:
                    colunas_finais.append(nome_original)
                else:
                    # Tentar encontrar coluna similar
                    encontrada = False
                    for col in colunas_disponiveis:
                        if nome_original.upper() in col.upper():
                            colunas_necessarias[novo_nome] = col
                            colunas_finais.append(col)
                            encontrada = True
                            print(f"Usando '{col}' como '{novo_nome}'")
                            break
                    
                    if not encontrada and novo_nome != 'CF':  # CF pode ser opcional
                        print(f"ATENÇÃO: Coluna '{nome_original}' não encontrada para mapear '{novo_nome}'")
            
            # Selecionar apenas as colunas que existem
            df = df[colunas_finais].copy()
            
            # Renomear para nomes padrão
            rename_dict = {v: k for k, v in colunas_necessarias.items() if v in df.columns}
            df = df.rename(columns=rename_dict)
            
            # Garantir que temos as colunas necessárias
            if 'DATA' not in df.columns:
                print("ERRO: Coluna 'DATA' não encontrada")
                return None
            
            if 'PRODUTO' not in df.columns:
                print("ERRO: Coluna 'PRODUTO' não encontrada")
                return None
            
            # Converter dados
            df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True).dt.date
            df['CUSTO'] = pd.to_numeric(df['CUSTO'], errors='coerce')
            
            if 'CF' in df.columns:
                df['CF'] = df['CF'].astype(str).str.strip()
            else:
                # Se não tiver CF, criar coluna vazia
                df['CF'] = ''
                print("ATENÇÃO: Coluna 'CF' não encontrada, criando coluna vazia")
        
        return df.dropna(subset=['PRODUTO', 'DATA'])
    
    except Exception as e:
        print(f"\nERRO CRÍTICO ao carregar {file_type}: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

# Carregar os dados
print("Carregando dados...")
print("\n--- Carregando arquivo de custo ---")
df_custo = load_data(custo_path, 'custo')

print("\n--- Carregando arquivo de margem (aba FEC_PQ) ---")
df_margem = load_data(margem_path, 'margem')

if df_custo is None or df_margem is None:
    print("\nFalha crítica no carregamento. Verifique os erros acima.")
    exit()

# Verificação final dos dados
print("\nVERIFICAÇÃO FINAL:")
print(f"Total de registros de custo: {len(df_custo)}")
print(f"Total de registros de margem: {len(df_margem)}")
print(f"Colunas em df_margem: {df_margem.columns.tolist()}")
print("\nAmostra dos dados de margem (primeiras 5 linhas):")
print(df_margem.head(10))

if 'CF' in df_margem.columns:
    print(f"\nValores únicos em CF: {df_margem['CF'].unique()}")
    print(f"Registros com CF='ESP': {len(df_margem[df_margem['CF'] == 'ESP'])}")

# Merge dos dados
print("\n--- Realizando merge dos dados ---")
print(f"Merge baseado em: PRODUTO e DATA")
print(f"Tamanho df_margem: {len(df_margem)}")
print(f"Tamanho df_custo: {len(df_custo)}")

merged = pd.merge(
    df_margem,
    df_custo,
    on=['PRODUTO', 'DATA'],
    how='left'
)

print(f"Tamanho após merge: {len(merged)}")

# Processamento adicional
# Combinar descrições
if 'DESCRICAO' in merged.columns:
    merged['DESCRICAO'] = merged['DESCRICAO'].fillna(merged.get('DESCRICAO_JULHO', ''))
elif 'DESCRICAO_JULHO' in merged.columns:
    merged['DESCRICAO'] = merged['DESCRICAO_JULHO']
else:
    merged['DESCRICAO'] = ''

# Classificar status
merged['STATUS'] = np.where(
    merged['CUSTO_JULHO'].isna(),
    'NÃO ENCONTRADO',
    np.where(
        abs(merged['CUSTO'] - merged['CUSTO_JULHO']) <= (0.01 * merged['CUSTO']),
        'IGUAL',
        'DIFERENTE'
    )
)

# Ordenação final
result_cols = ['CF', 'PRODUTO', 'DESCRICAO', 'DATA', 'CUSTO', 'CUSTO_JULHO', 'STATUS']
# Garantir que todas as colunas existam
result_cols_existentes = [col for col in result_cols if col in merged.columns]
final_result = merged[result_cols_existentes].sort_values(['STATUS', 'PRODUTO', 'DATA'])

# Salvar resultados
print(f"\nSalvando resultados em {output_path}")
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for status in ['IGUAL', 'DIFERENTE', 'NÃO ENCONTRADO']:
        sheet_name = status if status != 'NÃO ENCONTRADO' else 'NAO_ENCONTRADO'
        if status in final_result['STATUS'].unique():
            df_to_save = final_result[final_result['STATUS'] == status]
            print(f"- Salvando {len(df_to_save)} registros em {sheet_name}")
            df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)

# Estatísticas finais
print("\n=== RESUMO FINAL ===")
print(f"Total de registros processados: {len(final_result)}")
print(f"- IGUAIS: {len(final_result[final_result['STATUS'] == 'IGUAL'])}")
print(f"- DIFERENTES: {len(final_result[final_result['STATUS'] == 'DIFERENTE'])}")
print(f"- NÃO ENCONTRADOS: {len(final_result[final_result['STATUS'] == 'NÃO ENCONTRADO'])}")

print("\nProcesso concluído com sucesso! Verifique o arquivo gerado.")