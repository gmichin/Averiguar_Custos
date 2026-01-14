import pandas as pd
import numpy as np
from pathlib import Path

# Definindo os caminhos dos arquivos
csv_path = r"C:\Users\win11\OneDrive\Documentos\Custos Médios\2026\Janeiro\ev140126.csv"
xlsx_path = r"S:\hor\arquivos\mario\CONTROLE DE NOTAS ATUALIZADO.xlsx"
output_path = str(Path.home() / "Downloads" / "Averiguar_Custos (EV x NOTA).xlsx")

# Lendo e preparando o arquivo CSV
df_csv = pd.read_csv(csv_path, sep=';', header=2, encoding='latin1')
df_csv.columns = ['PRODUTO', 'DESCRICAO', 'GRUPO', 'PCS', 'KGS', 'CUSTO', 'TOTAL']
df_csv['CUSTO'] = pd.to_numeric(df_csv['CUSTO'].str.replace(',', '.'), errors='coerce')

# Lista de produtos com valores de referência especiais (originais)
produtos_especiais_originais = {
    '700': 21.35,  # Big bacon
    '845': 15.43, '809': 15.43, '1452': 15.43, '1428': 15.43,  # Paleta
    '1446': 14.10, '755': 14.10, '848': 14.10, '1433': 14.10, '1095': 14.10,  # Costela
    '1448': 7.9, '817': 7.9, '849': 7.9, '1430': 7.9,  # Lingua
    '846': 17.91, '878': 17.91, '1432': 17.91, '1451': 17.91,  # Lombo  
    '1426': 4.05, '1447': 4.05, '850': 4.05, '746': 4.05,  # Orelha
    '1427': 2.93, '836': 2.93, '852': 2.93, '1450': 2.93,  # Pé
    '1425': 9.68, '750': 9.68,  # Ponta
    '851': 14.26, '1449': 14.26, '1429': 14.26, '748': 14.26  # Rabo
}

# Lista de produtos para verificação em "Não Encontrados"
produtos_verificar_nao_encontrados = {
    '1721': 11.8, '1844': 23.43, '1833': 19.5, '1639': 20.55,
    '1690': 15.25, '1567': 10, '1816': 11.98, '1766': 23.2,
    '1856': 12, '1720': 24.33, '1817': 13, '1945': 6.99,
    '1177': 13, '1750': 3.83, '1484': 19.76, '1788': 18.36,
    '1179': 17, '1354': 16, '1673': 25.7, '1546': 10.33,
    '1881': 14.7, '1211': 42.43, '1713': 19.99, '1131': 42.26,
    '1893': 30.9, '807': 16.80, '1667': 6.98, '1873': 7.9,
    '1752': 18.88, '1819': 38.2, '1597': 3.9, '1675': 6, '1481': 18,
    '1510': 10.4, '1781': 11.49, '1711': 35, '1796': 7.25, 
    '1420': 14.3, '1793': 3, '1575': 11.90, '1828': 20.50,
    '1826': 24.98, '1759': 10, '1496': 35.9, '1909': 19,
    '1717': 8.75, '1621': 8.42, '822': 13.69, '1677': 8.15,
    '1969': 14.2, '1970': 14.2, '1827': 8.98, '1407': 6,
    '1434': 16, '1444': 20, '1335': 20.5, '1218': 30.5, '198': 17, 
    '902': 9.9, '1927': 51.30, '1265': 26.30, '1708': 1.99, '1282': 8.9
}

# Juntando todos os valores de referência
todos_valores_referencia = {**produtos_especiais_originais, **produtos_verificar_nao_encontrados}

# Lendo e preparando o arquivo XLSX
print("Lendo arquivo XLSX...")
xls = pd.ExcelFile(xlsx_path)
df_xlsx_all = []

# Colunas que realmente precisamos
colunas_desejadas = ['PRODUTO', 'CUSTO UNITÁRIO', 'NEGOCIADO', 'DATA']

for sheet in xls.sheet_names:
    print(f"Processando aba: {sheet}")
    
    try:
        # Ler apenas algumas linhas primeiro para verificar as colunas
        df_sample = pd.read_excel(xls, sheet_name=sheet, nrows=10)
        df_sample.columns = [str(col).upper().strip() for col in df_sample.columns]
        
        # Verificar quais colunas estão presentes
        colunas_presentes = [col for col in colunas_desejadas if any(col in col_name for col_name in df_sample.columns)]
        
        if not colunas_presentes:
            print(f"  Aviso: Nenhuma coluna relevante encontrada na aba {sheet}")
            continue
        
        # Ler a aba inteira, mas apenas as colunas necessárias
        df_sheet = pd.read_excel(xls, sheet_name=sheet, usecols=lambda x: any(col in str(x).upper() for col in colunas_desejadas))
        df_sheet.columns = [str(col).upper().strip() for col in df_sheet.columns]
        
        # Renomear colunas se necessário
        for col in df_sheet.columns:
            if 'NEGOCIADO' in col and 'CUSTO UNITÁRIO' not in df_sheet.columns:
                df_sheet = df_sheet.rename(columns={col: 'CUSTO UNITÁRIO'})
                break
        
        # Garantir que temos as colunas necessárias
        if 'PRODUTO' not in df_sheet.columns:
            print(f"  Aviso: Coluna 'PRODUTO' não encontrada na aba {sheet}")
            continue
        
        if 'CUSTO UNITÁRIO' not in df_sheet.columns:
            print(f"  Aviso: Coluna 'CUSTO UNITÁRIO' não encontrada na aba {sheet}")
            continue
        
        # Converter tipos de dados
        df_sheet['PRODUTO'] = pd.to_numeric(df_sheet['PRODUTO'], errors='coerce').astype('Int64')
        df_sheet['CUSTO UNITÁRIO'] = pd.to_numeric(df_sheet['CUSTO UNITÁRIO'], errors='coerce')
        
        # Processar data se existir
        if 'DATA' in df_sheet.columns:
            df_sheet['DATA'] = pd.to_datetime(df_sheet['DATA'], dayfirst=True, errors='coerce')
            df_sheet['DATA'] = df_sheet['DATA'].dt.date
        
        # Manter apenas colunas essenciais
        colunas_manter = ['PRODUTO', 'CUSTO UNITÁRIO']
        if 'DATA' in df_sheet.columns:
            colunas_manter.append('DATA')
        
        df_sheet = df_sheet[colunas_manter]
        df_xlsx_all.append(df_sheet)
        
        print(f"  Processado: {len(df_sheet)} linhas")
        
    except Exception as e:
        print(f"  Erro ao processar aba {sheet}: {str(e)}")
        continue

# Concatenar todos os dataframes
if df_xlsx_all:
    df_xlsx_all = pd.concat(df_xlsx_all, ignore_index=True)
    print(f"Total de registros lidos do XLSX: {len(df_xlsx_all)}")
    
    # Ordenar por data (se existir) e remover duplicados mantendo o mais recente
    if 'DATA' in df_xlsx_all.columns:
        df_xlsx_all = df_xlsx_all.sort_values('DATA', ascending=False)
        df_xlsx_unique = df_xlsx_all.drop_duplicates(subset=['PRODUTO'], keep='first')
    else:
        df_xlsx_unique = df_xlsx_all.drop_duplicates(subset=['PRODUTO'], keep='first')
    
    print(f"Produtos únicos no XLSX: {len(df_xlsx_unique)}")
else:
    print("Nenhum dado válido foi lido do arquivo XLSX!")
    df_xlsx_unique = pd.DataFrame(columns=['PRODUTO', 'CUSTO UNITÁRIO', 'DATA'])

# Juntando os dataframes
print("Juntando dados...")
result = pd.merge(
    df_csv, 
    df_xlsx_unique[['PRODUTO', 'CUSTO UNITÁRIO', 'DATA']], 
    on='PRODUTO', 
    how='left'
)

# Adicionando coluna com todos os valores de referência
result['PRODUTO_STR'] = result['PRODUTO'].astype(str)
result['VALOR_REFERENCIA'] = result['PRODUTO_STR'].map(todos_valores_referencia)

# Definindo tolerância para comparação
TOLERANCIA = 0.01  # 1% de diferença

# Classificando os resultados
result['STATUS'] = 'NÃO ENCONTRADO'
result.loc[~result['CUSTO UNITÁRIO'].isna(), 'STATUS'] = 'DIFERENTE'

# Verificar se são iguais (dentro da tolerância)
mask_iguais = ~result['CUSTO UNITÁRIO'].isna() & ~result['CUSTO'].isna()
result.loc[mask_iguais & np.isclose(result['CUSTO'], result['CUSTO UNITÁRIO'], rtol=TOLERANCIA), 'STATUS'] = 'IGUAL'

# Classificação especial para produtos especiais originais
for produto, valor_ref in produtos_especiais_originais.items():
    mask = result['PRODUTO_STR'] == produto
    if any(mask):
        custo_valor = result.loc[mask, 'CUSTO'].iloc[0]
        if not pd.isna(custo_valor) and custo_valor >= valor_ref:
            result.loc[mask, 'STATUS'] = 'IGUAL'
        else:
            result.loc[mask, 'STATUS'] = 'DIFERENTE'

# Adicionando coluna de COMPARAÇÃO apenas para produtos não encontrados que estão na lista de verificação
result['COMPARACAO'] = ''
mask_nao_encontrados_verificar = (
    (result['STATUS'] == 'NÃO ENCONTRADO') & 
    (result['PRODUTO_STR'].isin(produtos_verificar_nao_encontrados.keys()))
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
print("Criando tabelas finais...")

tabela1 = result[result['STATUS'] == 'IGUAL'][['PRODUTO', 'DESCRICAO', 'CUSTO', 'VALOR_REFERENCIA', 'CUSTO UNITÁRIO', 'DATA']]
tabela1.columns = ['Código', 'Descrição', 'Custo Estoque', 'Valor Referência', 'Custo Nota', 'Data Nota']

tabela2 = result[result['STATUS'] == 'DIFERENTE'][['PRODUTO', 'DESCRICAO', 'CUSTO', 'VALOR_REFERENCIA', 'CUSTO UNITÁRIO', 'DATA']]
tabela2.columns = ['Código', 'Descrição', 'Custo Estoque', 'Valor Referência', 'Custo Nota', 'Data Nota']

tabela3 = result[result['STATUS'] == 'NÃO ENCONTRADO'][['PRODUTO', 'DESCRICAO', 'CUSTO', 'VALOR_REFERENCIA', 'COMPARACAO']]
tabela3.columns = ['Código', 'Descrição', 'Custo Estoque', 'Valor Referência', 'Comparação']

# Salvando os resultados
print("Salvando resultados...")
with pd.ExcelWriter(output_path) as writer:
    tabela1.to_excel(writer, sheet_name='Custos_Iguais', index=False)
    tabela2.to_excel(writer, sheet_name='Custos_Diferentes', index=False)
    tabela3.to_excel(writer, sheet_name='Produtos_Nao_Encontrados', index=False)

print(f"\nRelatório gerado com sucesso em: {output_path}")
print("\nResumo:")
print(f"Produtos com custos iguais: {len(tabela1)}")
print(f"Produtos com custos diferentes: {len(tabela2)}")
print(f"Produtos não encontrados: {len(tabela3)}")
print(f"Total de produtos processados: {len(result)}")