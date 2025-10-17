import pandas as pd
import numpy as np
from pathlib import Path

# Definindo os caminhos dos arquivos
csv_path = r"C:\Users\win11\OneDrive\Documentos\Custos Médios\2025\Outubro\ev171025.csv"
#"Z:\ANDRIELLY\CONTROLE DE NOTAS.xlsx"
xlsx_path = r"S:\hor\arquivos\mario\CONTROLE DE NOTAS ATUALIZADO.xlsx"
output_path = str(Path.home() / "Downloads" / "Averiguar_Custos (EV x NOTA).xlsx")

# Lendo e preparando o arquivo CSV
df_csv = pd.read_csv(csv_path, sep=';', header=2, encoding='latin1')
df_csv.columns = ['PRODUTO', 'DESCRICAO', 'GRUPO', 'PCS', 'KGS', 'CUSTO', 'TOTAL']
df_csv['CUSTO'] = pd.to_numeric(df_csv['CUSTO'].str.replace(',', '.'), errors='coerce')

# Lista de produtos com valores de referência especiais (originais)
produtos_especiais_originais = {
    # Big bacon
    '700': 23.4, 
    # Paleta
    '845': 16.09, '809': 16.09, '1452': 16.09, '1428': 16.09,
    # Costela
    '1446': 14.13, '755': 14.13, '848': 14.13, '1433': 14.13, '1095': 14.13,
    # Lingua
    '1448': 7.87, '817': 7.87, '849': 7.87, '1430': 7.87, 
    # Lombo
    '846': 17.85, '878': 17.85, '1432': 17.85, '1451': 17.85, 
    # Orelha
    '1426': 5.06, '1447': 5.06, '850': 5.06, '746': 5.06,
    # Pé
    '1427': 3.54, '836': 3.54, '852': 3.54, '1450': 3.54, 
    # Ponta
    '1425': 9.84, '750': 9.84, 
    # Rabo
    '851': 15.70, '1449': 15.70, '1429': 15.70, '748': 15.70
}

# Lista de produtos para verificação em "Não Encontrados"
produtos_verificar_nao_encontrados = {
    '1721': 11.8, '1844': 23.43, '1833': 19.5, '1639': 20.55,
    '1690': 15.25,'1567': 10, '1816': 11.98, '1766': 23.2,
    '1856': 12, '1720': 24.33, '1817': 13, '1945': 6.99,
    '1177': 13, '1750': 3.83, '1484': 19.76, '1788': 18.36,
    '1179': 17, '1354': 16, '1673': 25.7, '1795': 29.36, '1546': 10.33,
    '1881': 14.7, '1211': 42.43, '1713': 19.99, '1131': 42.26,
    '1893': 30.9, '1691': 9.1, '807': 16.80, '1667': 6.98, '1873': 7.9,
    '1752': 18.88, '1819': 38.2, '1597': 3.9, '1675': 6,
    '1510': 10.4, '1781': 11.49, '1711': 35, '1796': 7.25,
    '1420': 14.3, '1793': 3, '1547': 40, '1575': 20.83, '1828': 24.69,
    '1826': 24.98, '1116': 23.06, '1759': 10, '1496': 34.95,
    '1717': 8.75, '1621': 8.42, '1624': 1.94, '822': 13.69,
    '1969': 14.2, '1970': 14.2, '1827': 8.98, '1407': 6,
    '1434': 16, '1444': 20, '1335':20.5, '1218': 30.5, '1648': 3.9,
    '902':  9.9, '1927': 51.30, '1265': 26.30, '1708': 1.99, '1282': 8.9
}

# Juntando todos os valores de referência
todos_valores_referencia = {**produtos_especiais_originais, **produtos_verificar_nao_encontrados}

# Lendo e preparando o arquivo XLSX
xls = pd.ExcelFile(xlsx_path)
df_xlsx_all = pd.DataFrame()

for sheet in xls.sheet_names:
    df_sheet = pd.read_excel(xls, sheet_name=sheet)
    df_sheet.columns = [col.upper().strip() for col in df_sheet.columns]
    
    if 'NEGOCIADO' in df_sheet.columns and 'CUSTO UNITÁRIO' not in df_sheet.columns:
        df_sheet = df_sheet.rename(columns={'NEGOCIADO': 'CUSTO UNITÁRIO'})
    
    if 'CUSTO UNITÁRIO' in df_sheet.columns:
        df_sheet['CUSTO UNITÁRIO'] = pd.to_numeric(df_sheet['CUSTO UNITÁRIO'], errors='coerce')
    
    df_xlsx_all = pd.concat([df_xlsx_all, df_sheet], ignore_index=True)

# Processando as datas e ordenando
df_xlsx_all['DATA'] = pd.to_datetime(df_xlsx_all['DATA'], dayfirst=True, errors='coerce').dt.date
df_xlsx_all = df_xlsx_all.sort_values('DATA', ascending=False)

# Mantendo apenas o registro mais recente de cada produto
df_xlsx_unique = df_xlsx_all.drop_duplicates(subset=['PRODUTO'], keep='first')

# Juntando os dataframes
result = pd.merge(
    df_csv, 
    df_xlsx_unique[['PRODUTO', 'CUSTO UNITÁRIO', 'DATA']], 
    on='PRODUTO', 
    how='left'
)

# Adicionando coluna com todos os valores de referência
result['VALOR_REFERENCIA'] = result['PRODUTO'].astype(str).map(todos_valores_referencia)

# Definindo tolerância para comparação
TOLERANCIA = 0.01  # 1% de diferença

# Classificando os resultados
result['STATUS'] = 'NÃO ENCONTRADO'
result.loc[~result['CUSTO UNITÁRIO'].isna(), 'STATUS'] = 'DIFERENTE'
result.loc[np.isclose(result['CUSTO'], result['CUSTO UNITÁRIO'], rtol=TOLERANCIA, equal_nan=True), 'STATUS'] = 'IGUAL'

# Classificação especial para produtos especiais originais
for produto, valor_ref in produtos_especiais_originais.items():
    mask = result['PRODUTO'].astype(str) == produto
    if any(mask):
        if result.loc[mask, 'CUSTO'].iloc[0] >= valor_ref:
            result.loc[mask, 'STATUS'] = 'IGUAL'
        else:
            result.loc[mask, 'STATUS'] = 'DIFERENTE'

# Adicionando coluna de COMPARAÇÃO apenas para produtos não encontrados que estão na lista de verificação
result['COMPARACAO'] = ''
mask_nao_encontrados_verificar = (
    (result['STATUS'] == 'NÃO ENCONTRADO') & 
    (result['PRODUTO'].astype(str).isin(produtos_verificar_nao_encontrados.keys()))
)

for idx in result[mask_nao_encontrados_verificar].index:
    custo = result.at[idx, 'CUSTO']
    valor_ref = result.at[idx, 'VALOR_REFERENCIA']
    
    if pd.isna(custo) or pd.isna(valor_ref):
        continue
    
    if np.isclose(custo, valor_ref, rtol=0.01):
        result.at[idx, 'COMPARACAO'] = 'IGUAL'
    elif custo > valor_ref:
        result.at[idx, 'COMPARACAO'] = 'MAIOR'
    else:
        result.at[idx, 'COMPARACAO'] = 'MENOR'

# Criando as tabelas finais
# Mantendo VALOR_REFERENCIA em todas as abas para produtos especiais originais
tabela1 = result[result['STATUS'] == 'IGUAL'][['PRODUTO', 'DESCRICAO', 'CUSTO', 'VALOR_REFERENCIA', 'CUSTO UNITÁRIO', 'DATA']]
tabela1.columns = ['Código', 'Descrição', 'Custo Estoque', 'Valor Referência', 'Custo Nota', 'Data Nota']

tabela2 = result[result['STATUS'] == 'DIFERENTE'][['PRODUTO', 'DESCRICAO', 'CUSTO', 'VALOR_REFERENCIA', 'CUSTO UNITÁRIO', 'DATA']]
tabela2.columns = ['Código', 'Descrição', 'Custo Estoque', 'Valor Referência', 'Custo Nota', 'Data Nota']

tabela3 = result[result['STATUS'] == 'NÃO ENCONTRADO'][['PRODUTO', 'DESCRICAO', 'CUSTO', 'VALOR_REFERENCIA', 'COMPARACAO']]
tabela3.columns = ['Código', 'Descrição', 'Custo Estoque', 'Valor Referência', 'Comparação']

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