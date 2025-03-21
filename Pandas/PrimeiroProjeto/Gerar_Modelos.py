import pandas as pd
from datetime import datetime as dt
from datetime import timedelta 
import Mod1 as fc
from openpyxl import load_workbook
import os


hoje = dt.now()
mes_anterior = hoje.replace(day=1) - timedelta(days=1)
mes = mes_anterior.strftime('%b')  # Nome abreviado do mês (ex: "Jan")
mes_num = mes_anterior.month  # Número do mês (ex: 1 para Janeiro)
caminho_Salvar = os.path.join(f"//apolo/Governanca/PROCESSOS/MAPEAMENTO DE PROCESSOS/CODIGOS E AUTOMACOES/Codigos.PY/Codigos.PY/Tratar_dados_OKR's_KPI's/Templates", mes)

arquivo_salvo = "Resultado_Final/OKReKPI - Versão 6.xlsx"
NOMES = ["OKR","KPI"]

df = pd.read_excel(arquivo_salvo,sheet_name=None)
OKR = df['OKR']
KPI = df['KPI']

Diretorias = list(OKR['Departamento'].unique())
Nomes_Diretorias = Diretorias.copy()
Nomes_Diretorias[7] = "Expansão&Operação"

if not os.path.exists(caminho_Salvar):
    os.makedirs(caminho_Salvar)
    
for i in Diretorias:
    df_filtrado_OKR = OKR[(OKR['Departamento'] == i)]
    df_filtrado_KPI = KPI[(KPI['Departamento'] == i)]
    ARQUIVO = f"Templates/{mes}/OKR_KPI - {Nomes_Diretorias[Diretorias.index(i)]}_{mes}.xlsx"

    df_filtrado_OKR = df_filtrado_OKR.drop(
        columns= ['Meses Acomp','Comparação','Projetado','início','Período considerado (M)','Modelo de apuração','Descrição','Data de entrega'],
        errors='ignore'
        )
    df_filtrado_KPI = df_filtrado_KPI.drop(
        columns = ['Meses Acomp','Comparação','início','Período considerado (M)','Modelo de apuração','Descrição','Projetado'],
        errors = 'ignore'
        )

    with pd.ExcelWriter(ARQUIVO, engine='openpyxl') as writer:
        df_filtrado_OKR.to_excel(writer, sheet_name='OKR', index=False)
        df_filtrado_KPI.to_excel(writer, sheet_name='KPI', index=False)
        
    for sheet in NOMES:    
        WB = load_workbook(ARQUIVO)
        WS = WB[sheet]
        
        inicio = 6 if sheet == "OKR" else 5
        inicio_mes = (inicio + 3*mes_num)-2
        
        for col_index in range(inicio, 100):
            col_letter = fc.numero_para_coluna(col_index)  # Converte o número da coluna para a letra correspondente
            WS.column_dimensions[col_letter].hidden = True
            if col_index in [inicio_mes, inicio_mes+1, inicio_mes +2]:
                WS.column_dimensions[col_letter].hidden = False

        WB.save(ARQUIVO)
        
    print (f"{ARQUIVO} salvo com sucesso")