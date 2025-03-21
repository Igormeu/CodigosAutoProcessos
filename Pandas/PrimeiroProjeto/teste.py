import pandas as pd
from datetime import datetime as dt
from datetime import timedelta 
import Mod1 as fc
from openpyxl import load_workbook


# Variáveis Globais
Hoje = dt.now()
mes_anterior = Hoje.replace(day=1) - timedelta(days=1)
mesAbrev = mes_anterior.strftime("%b")
arquivo_salvo = "Resultado_Final/OKReKPI - Versão 6.xlsx"
OKR_final_arq = r"//apolo/Governanca/PROCESSOS/MAPEAMENTO DE PROCESSOS/CODIGOS E AUTOMACOES/Codigos.PY/Codigos.PY/Tratar_dados_OKR's_KPI's/Resultado_Final/Acompanhamenato de KR's - 2025.xlsx"
KPI_final_arq = r"//apolo/Governanca/PROCESSOS/MAPEAMENTO DE PROCESSOS/CODIGOS E AUTOMACOES/Codigos.PY/Codigos.PY/Tratar_dados_OKR's_KPI's/Resultado_Final/Acompanhamenato de KPI - 2025.xlsx"

Arquivos_Finais = [OKR_final_arq,KPI_final_arq]
# Carregando dados
df = pd.read_excel("Consolidado OKR e KPI's - Versão 2.0.xlsx", sheet_name=None)
okr = pd.read_excel("OKR - Versão 3.xlsx", sheet_name='OKR')
kpi = df['KPI']
listas = [okr, kpi]  
nomes = ["OKR", "KPI"]

# Processando os dados e escrevendo no arquivo
for j in range(len(listas)):
    tabela = listas[j]
    tabela['Departamento'] = tabela['Departamento'].apply(lambda x: "Expansão&Operação" if x == "Expansão/Lojas" else x)
    fc.AtribuirApurar(Arquivos_Finais[j],tabela)
    

