import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

# Configurações
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

# Caminhos dos arquivos
custo_path = r"C:\Users\win11\Downloads\Custos de produtos - Janeiro.xlsx"
margem_path = r"C:\Users\win11\Downloads\MRG_260131 - wapp - v2.xlsx"
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
            df = pd.read_excel(margem_path, sheet_name="Base (3,5%)", header=8)
            
            # Filtrar apenas CF='ESP' (ou outros valores relevantes)
            df = df[df['CF'] == 'ESP'].copy()
            
            # Converter para string removendo .0 de números e espaços
            df['CODPRODUTO'] = df['CODPRODUTO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            df = df[['CODPRODUTO', 'DATA', 'DESCRICAO', 'CUSTO', 'CF']].copy()
            df.columns = ['PRODUTO', 'DATA', 'DESCRICAO_MARGEM', 'CUSTO_MARGEM', 'CF']
            
            df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True).dt.date
            df['CUSTO_MARGEM'] = pd.to_numeric(df['CUSTO_MARGEM'], errors='coerce')
            df['CF'] = df['CF'].astype(str).str.strip()
        
        return df.dropna(subset=['PRODUTO', 'DATA'])
    
    except Exception as e:
        print(f"\nERRO CRÍTICO ao carregar {file_type}: {str(e)}")
        return None

# Carregar os dados
print("Carregando dados...")
df_custo = load_data(custo_path, 'custo')
df_margem = load_data(margem_path, 'margem')

if df_custo is None or df_margem is None:
    print("\nFalha crítica no carregamento. Verifique os erros acima.")
    exit()

# Verificação final dos dados
print("\nVERIFICAÇÃO FINAL:")
print(f"Total de registros de custo: {len(df_custo)}")
print(f"Total de registros de margem: {len(df_margem)}")
print("\nAmostra dos dados de margem (com CF):")
print(df_margem[['PRODUTO', 'CF']].head(10))

# Merge dos dados
merged = pd.merge(
    df_margem,
    df_custo,
    on=['PRODUTO', 'DATA'],
    how='left'
)

# Processamento adicional
merged['DESCRICAO'] = merged['DESCRICAO_MARGEM'].fillna(merged['DESCRICAO_JULHO'])

merged['STATUS'] = np.where(
    merged['CUSTO_JULHO'].isna(),
    'NÃO ENCONTRADO',
    np.where(
        abs(merged['CUSTO_MARGEM'] - merged['CUSTO_JULHO']) <= (0.01 * merged['CUSTO_MARGEM']),
        'IGUAL',
        'DIFERENTE'
    )
)

# Ordenação final
result_cols = ['CF', 'PRODUTO', 'DESCRICAO', 'DATA', 'CUSTO_MARGEM', 'CUSTO_JULHO', 'STATUS']
final_result = merged[result_cols].sort_values(['STATUS', 'PRODUTO', 'DATA'])

# Salvar resultados
print(f"\nSalvando resultados em {output_path}")
with pd.ExcelWriter(output_path) as writer:
    for status in ['IGUAL', 'DIFERENTE', 'NÃO ENCONTRADO']:
        sheet_name = status if status != 'NÃO ENCONTRADO' else 'NAO_ENCONTRADO'
        if status in final_result['STATUS'].unique():
            df_to_save = final_result[final_result['STATUS'] == status]
            print(f"- Salvando {len(df_to_save)} registros em {sheet_name}")
            df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)

print("\nProcesso concluído com sucesso! Verifique o arquivo gerado.")