import pandas as pd
import numpy as np
from pathlib import Path

# Definindo os caminhos dos arquivos
csv_path = r"C:\Users\win11\Downloads\ev280725.csv"
xlsx_path = r"Z:\ANDRIELLY\CONTROLE DE NOTAS.xlsx"
output_path = str(Path.home() / "Downloads" / "Averiguar_Custos.xlsx")

# Lendo e preparando o arquivo CSV
df_csv = pd.read_csv(csv_path, sep=';', header=2, encoding='latin1')
df_csv.columns = ['PRODUTO', 'DESCRICAO', 'GRUPO', 'PCS', 'KGS', 'CUSTO', 'TOTAL']
df_csv['CUSTO'] = pd.to_numeric(df_csv['CUSTO'].str.replace(',', '.'), errors='coerce')  # Convertendo para numérico

# Lendo e preparando o arquivo XLSX
xls = pd.ExcelFile(xlsx_path)
df_xlsx_all = pd.DataFrame()

for sheet in xls.sheet_names:
    df_sheet = pd.read_excel(xls, sheet_name=sheet)
    df_sheet.columns = [col.upper().strip() for col in df_sheet.columns]
    
    # Padronizando nome da coluna de custo
    if 'NEGOCIADO' in df_sheet.columns and 'CUSTO UNITÁRIO' not in df_sheet.columns:
        df_sheet = df_sheet.rename(columns={'NEGOCIADO': 'CUSTO UNITÁRIO'})
    
    # Convertendo valores para numérico
    if 'CUSTO UNITÁRIO' in df_sheet.columns:
        df_sheet['CUSTO UNITÁRIO'] = pd.to_numeric(df_sheet['CUSTO UNITÁRIO'], errors='coerce')
    
    df_xlsx_all = pd.concat([df_xlsx_all, df_sheet], ignore_index=True)

# Processando as datas e ordenando
df_xlsx_all['DATA'] = pd.to_datetime(df_xlsx_all['DATA'], dayfirst=True, errors='coerce').dt.date
df_xlsx_all = df_xlsx_all.sort_values('DATA', ascending=False)

# Mantendo apenas o registro mais recente de cada produto
df_xlsx_unique = df_xlsx_all.drop_duplicates(subset=['PRODUTO'], keep='first')

# Juntando os dataframes
result = pd.merge(df_csv, df_xlsx_unique[['PRODUTO', 'CUSTO UNITÁRIO', 'DATA']], 
                 on='PRODUTO', how='left')

# Definindo tolerância para comparação
TOLERANCIA = 0.01  # 1% de diferença

# Classificando os resultados
result['STATUS'] = 'NÃO ENCONTRADO'
result.loc[~result['CUSTO UNITÁRIO'].isna(), 'STATUS'] = 'DIFERENTE'
result.loc[np.isclose(result['CUSTO'], result['CUSTO UNITÁRIO'], rtol=TOLERANCIA, equal_nan=True), 'STATUS'] = 'IGUAL'

# Criando as tabelas finais
tabela1 = result[result['STATUS'] == 'IGUAL'][['PRODUTO', 'DESCRICAO', 'CUSTO', 'CUSTO UNITÁRIO', 'DATA']]
tabela1.columns = ['Código', 'Descrição', 'Custo Estoque', 'Custo Nota', 'Data Nota']

tabela2 = result[result['STATUS'] == 'DIFERENTE'][['PRODUTO', 'DESCRICAO', 'CUSTO', 'CUSTO UNITÁRIO', 'DATA']]
tabela2.columns = ['Código', 'Descrição', 'Custo Estoque', 'Custo Nota', 'Data Nota']

tabela3 = result[result['STATUS'] == 'NÃO ENCONTRADO'][['PRODUTO', 'DESCRICAO', 'CUSTO']]
tabela3.columns = ['Código', 'Descrição', 'Custo Estoque']

# Salvando os resultados
with pd.ExcelWriter(output_path) as writer:
    tabela1.to_excel(writer, sheet_name='Custos_Iguais', index=False)
    tabela2.to_excel(writer, sheet_name='Custos_Diferentes', index=False)
    tabela3.to_excel(writer, sheet_name='Produtos_Nao_Encontrados', index=False)

print(f"Relatório gerado com sucesso em: {output_path}")
print("\nResumo:")
print(f"Produtos com custos iguais: {len(tabela1)}")
print(f"Produtos com custos diferentes: {len(tabela2)}")
print(f"Produtos não encontrados: {len(tabela3)}")