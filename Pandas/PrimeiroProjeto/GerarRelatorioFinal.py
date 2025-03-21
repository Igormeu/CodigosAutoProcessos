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
    for index, row in tabela.iterrows():
        periodo_inicio = int(row['início'])
        periodo_fim = int(row['Período considerado (M)'])
        DepartRef = row['Departamento']
        CaminhoArqReal = f"Templates/{mesAbrev}/OKR_KPI - {row['Departamento']}_{mesAbrev}.xlsx"
        dfRealizado = pd.read_excel(CaminhoArqReal, sheet_name=None)
        CaminhoArqProje = f"Templates/Metas_Anuais/OKR_KPI - {row['Departamento']}_Metas.xlsx"
        dfProjetado = pd.read_excel(CaminhoArqProje,sheet_name=None)
        
        for i in range(periodo_inicio, periodo_fim + 1):
            # Coluna de projetados
            coluna_projetado = f'Projetado {i}/2025'
            if coluna_projetado not in tabela.columns:
                tabela[coluna_projetado] = ""

            projetado_value = fc.buscarEntreTabelas(row['ID'], dfProjetado[nomes[j]], 'ID', coluna_projetado)    
            tabela.loc[index, coluna_projetado] = projetado_value

            #Coluna de realizado
            coluna_realizado = f'Realizado {i}/2025'
            if i <= Hoje.month:
                if coluna_realizado not in tabela.columns:
                    tabela[coluna_realizado] = ""
                    
                realizado_value = fc.buscarEntreTabelas(row['ID'], dfRealizado[nomes[j]], "ID", coluna_realizado)    
                tabela.loc[index, coluna_realizado] = float(realizado_value) if realizado_value is not None else 0
            else:
                tabela.loc[index, coluna_realizado] = 0
            
            #Coluna de apurado
            apurado = fc.apuracaoMaM(
            tabela.loc[index, f'Projetado {i}/2025'],
            tabela.loc[index, f'Realizado {i}/2025'],
            tabela.loc[index, 'Comparação'])

            tabela.loc[index, f'Apurado {i}/2025'] = apurado['Resultado']
            
            #Coluna de apurado no tri
            if i % 3 == 0:
                tabela[f'Apurado do {int(i/3)}° Trimestre'] = ""

    # Remove colunas desnecessárias, ignorando erros caso não existam
    tabela = tabela.drop(
        columns=['Data de entrega', 'Comparação','Projetado'], #'início', 'Período considerado (M)'
        errors='ignore'
    )

    # Salvar o arquivo com ambas as planilhas (OKR e KPI)
    with pd.ExcelWriter(arquivo_salvo, mode='a', engine='openpyxl',if_sheet_exists='replace') as writer:
        for h in range(len(listas)):
            listas[h].to_excel(writer, index=False, sheet_name=nomes[h])
            
        print(f"Arquivo salvo corretamente em {arquivo_salvo}")

    wb = load_workbook(arquivo_salvo)
    ws = wb[nomes[j]]

    inicio = 11 if nomes[j] == "OKR" else 14
    
    for i in range(inicio + (mes_anterior.month * 3), 52):
        col_letter = fc.numero_para_coluna(i)  # Converte o número da coluna para a letra correspondente
        ws.column_dimensions[col_letter].hidden = True

    wb.save(arquivo_salvo)
    
    print(f"Arquivo {arquivo_salvo} - {nomes[j]} gerado com sucesso")
    
    try:
    
        fc.recalcular_planilha(Arquivos_Finais[j])

        print("Os modulos funcionaram")
    except Exception as e:
        print(e)

