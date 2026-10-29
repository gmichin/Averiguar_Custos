import pandas as pd
import numpy as np
from pathlib import Path

# Definindo os caminhos dos arquivos
csv_path = r"C:\Users\win11\VOG\AD - Documentos\Custos Médios\2026\Outubro\ev291026.csv"
xlsx_path = r"Z:\hor\arquivos\mario\CONTROLE DE NOTAS ATUALIZADO ATUAL.xlsx"
output_path = str(Path.home() / "Downloads" / "Averiguar_Custos (EV x NOTA).xlsx")

# Lendo e preparando o arquivo CSV
df_csv = pd.read_csv(csv_path, sep=';', header=2, encoding='latin1')
df_csv.columns = ['PRODUTO', 'DESCRICAO', 'GRUPO', 'PCS', 'KGS', 'CUSTO', 'TOTAL']
df_csv['CUSTO'] = pd.to_numeric(df_csv['CUSTO'].str.replace(',', '.'), errors='coerce')

# Lista de produtos com valores de referência especiais (originais)
produtos_especiais_originais = {
    '700': 22.72,  # Big bacon
    '845': 9.87, '809': 9.52, '1452': 11.35, '1428': 11.35,  # Paleta
    '1446': 11.06, '755': 9.41, '848': 9.56, '1433': 11.06, '1095': 9.5,  # Costela
    '1448': 7.26, '817': 6.33, '849': 5.68, '1430': 7.26,  # Lingua
    '846': 13.67, '878': 13.6, '1432': 15.40, '1451': 15.40,  # Lombo
    '1426': 3.83, '1447': 3.83, '850': 2.78, '746': 2.66,  # Orelha
    '1427': 4.07, '836': 2.52, '852': 3.01, '1450': 4.07,  # Pé
    '1425': 9.05, '750': 7.28,  # Ponta
    '851': 11.05, '1449': 12.39, '1429': 12.39, '748': 8.9  # Rabo
}

# Lista de produtos para verificação em "Não Encontrados"
produtos_verificar_nao_encontrados = {
    '1721': 11.8, '1844': 23.43, '1833': 19.5, '1639': 20.55, '6666': 27, 
    '1690': 12.54, '1816': 11.98, '1766': 23.2, '4091': 9.3, '3087': 6.26, 
    '1856': 12, '1720': 24.33, '1817': 13, '1945': 6.99, '2055': 4.5,
    '1177': 13, '1750': 3.83, '1484': 19.76, '17    88': 18.36, '2031': 5.5,
    '1179': 17, '1673': 25.7, '1546': 10.33, '2030': 5.51, '2083': 15.8,
    '1881': 14.7, '1211': 43.7, '1713': 19.99, '1131': 42.26, '3099': 4.5, 
    '1893': 30.9, '807': 16.80, '1667': 6.98, '1873': 7.9, '2034': 14,
    '1752': 18.88, '1819': 38.2, '1597': 3.9, '1675': 6, '2033': 8.7,
    '1510': 10.4, '1711': 35, '1796': 7.25, '2032': 5.5, '2033': 8.7,
    '1420': 14.3, '1793': 3, '1575': 11.90, '1828': 20.50, '2044': 25.99,
    '1826': 24.98, '1759': 10, '2054': 5.5, '1067': 5.84, '2089': 26,
    '1717': 8.75, '1621': 8.42, '822': 14.07, '1677': 8.15, '2053': 5.5,
    '1969': 14.2, '1970': 14.2, '1827': 8.98, '1342': 4.84, '5066': 40.01,
    '1444': 20, '1335': 29, '1218': 30.5, '198': 17, '906': 10, '4124': 60,
    '902': 9.9, '1927': 51.30, '1265': 26.30, '1708': 1.99, '1282': 8.9,
    '3070': 24.65, '1303': 11.9, '925': 9.5, '2989': 7.2, '2082': 4.3,
    '2981': 33, '8599': 22, '3031': 8.97, '6667': 43.5, '1739': 14.92,
    '1039': 4.2, '2087': 26, '3026': 31, '3021': 34.26, '749': 19.9,
    '2094': 29.9, '3003': 41.9, '3004': 7.21, '3002': 29.9, '3089': 2.1, 
    '3097': 2.5, '3098': 4.5, '1669': 7, '2056': 5.58, '4123': 28.5,
    '3088': 6.26, '3091': 2.4, '3040': 37, '4131': 36.4, '4130': 5.19,
    '1090': 8.5, '3027': 8.7, '4000': 12.5, '4031': 17.5, '3040': 37,
    '4035': 16.65, '4032': 7.5, '7002': 43.5, '4091': 9.3, '1089': 12,
    '4029': 17, '4557': 14.12, '2930': 28.58, '4014': 103.2, '4040': 25,
    '4022': 10, '1298': 14,'4047': 12.5, '4009': 38, '2930': 28.58,
    '3097': 2.5, '3098': 4.5,'2056': 5.58, '4022': 10, '6066': 27.42,
    '7002': 43.5, '3040': 37, '3040': 37, '3087': 6.26, '4047': 12.5, 
    '3088': 6.26, '  3091': 2.4, '4128': 2.97, '4088': 34, '4135': 10.48,
    '1090': 8.88, '3027': 8.7, '4000': 12.5, '4031': 17.5, '4122': 28.9,
    '4035': 16.65, '4032': 7.5, '1298': 14, '4040': 25, '6085': 7.5,
    '4029': 17, '4557': 14.12, '1859': 11.15, '758': 6.8, '757': 7.2,
    '4023': 53.99, '1063': 14.84, '4051': 21.3, '4054': 36.99, '4057': 71,
    '2521': 13.48, '4071': 9.07, '4072': 14.91, '4086': 27.92, '4115': 27.5,
    '4103': 9.5, '4104': 9.5, '4095': 2.97, '4096': 2.97, '4111': 10,
    '4097': 2.97, '4098': 2.97, '4099': 2.97, '4100': 2.97, '4101': 2.97,
    '4102': 2.97, '4110': 10.69, '1590': 15.9, '4112': 4.5, '4087': 33,
    '4116': 29, '4092': 58, '4117': 43, '1181': 32, '4090': 22.5, 
    '4114': 27.5, '4113': 10, '4052': 21, '702': 8.9, '6086': 6.8,
    '4120': 11.8, '4202': 5.2, '4105': 25.5, '4126': 13.98, '6070': 92,
    '4121': 6.99, '7125': 6.8, '4065': 2.5, '6045': 29.5, '6067': 35.5,
    '4132': 43.6, '6069': 30, '6071': 34.5, '5070': 42, '6090': 2.42, 
    '4133': 31.5,'6097': 2, '6065': 11.85, '7071': 16, '4118': 33.5,
    '4170': 17, '4155': 4.92, '4140': 21, '828': 8, '7078': 3.59,
    '6056': 26.5, '6098': 2.45, '6096': 2.42, '1489': 6, '7077': 10, 
    '6091': 2.03, '6092': 2.45, '6093': 2, '6094': 1.25, '6095': 2.45, 
    '6082': 6.5, '6080': 6.5, '6082': 6.5, '4172': 5.6, '7055': 44.5,
    '4171': 9.95, '6081': 6.5, '6099': 15, '2637': 30, '7073': 14,
    '7074': 15, '946': 15.59, '827': 6, '825': 6, '5092': 26.25,
    '5067': 7.07, '4191': 26.5, '4106': 29.5, '7000': 11, '7005': 17.5,
    '7006': 16, '4200': 33, '7004': 8.9, '7003': 8, '4150': 4.99
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