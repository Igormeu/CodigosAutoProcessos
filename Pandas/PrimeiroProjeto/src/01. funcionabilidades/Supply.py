import pandas as pd

# Caminho do arquivo
caminho_arquivos = caminho_arquivos = "//apolo/Governanca/PROCESSOS/MAPEAMENTO DE PROCESSOS/ACOMP KR E KPI/2025/SUPPLY/OKR e KPI - SUP 2025 (PREENCHIDA).xlsx"

# Lê todas as planilhas do Excel
df = pd.read_excel(caminho_arquivos, sheet_name=None)

# Seleciona as planilhas de interesse
okrs = df['Planilha de OKR - LOG']
nome_okrs = df['Planilha4']

# Converte o DataFrame de largo para longo
okrs_longo = okrs.melt(id_vars=["Mês abrv."], var_name="OKRs", value_name="Valor")

# Converte de longo para largo com os meses como colunas
okrs_formatado = okrs_longo.pivot(index="OKRs", columns="Mês abrv.", values="Valor")

# Define a ordem correta dos meses
ordem_meses = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

# Reordena as colunas conforme a ordem definida
okrs_formatado = okrs_formatado.reindex(columns=ordem_meses)

# Reseta o índice para que "OKRs" volte a ser uma coluna
okrs_formatado.reset_index(inplace=True)

# Realiza o mapeamento entre 'Cod. Do OKR' e 'OKR' de forma eficiente
cod_to_descr = dict(zip(nome_okrs['Cod. Do OKR'], nome_okrs['OKR']))

# Substitui os valores na coluna 'OKRs' pelo mapeamento
okrs_formatado['OKRs'] = okrs_formatado['OKRs'].map(cod_to_descr)

# Exibe o DataFrame final
arquivo = 'Resultado_Final/OKR SUP(Completo).xlsx'
okrs_formatado.to_excel(arquivo,sheet_name='OKR')
