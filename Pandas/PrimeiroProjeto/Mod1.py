import pandas as pd
import re
import win32com.client
import os
import numpy as np
from openpyxl import load_workbook

def atualizar_arquivo(caminho_arquivo, df_dict):

    try:
        # Fechar o arquivo caso esteja aberto (Win32)
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Workbooks.Open(caminho_arquivo).Close(SaveChanges=True)
            excel.Quit()
        except Exception:
            pass  # Ignorar erro caso não esteja aberto
        
        # Criar um ExcelWriter e sobrescrever todas as planilhas
        with pd.ExcelWriter(caminho_arquivo, mode='a', engine='openpyxl') as writer:
            for sheet_name, df in df_dict.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        print(f"Arquivo atualizado com sucesso em {caminho_arquivo}")
    
    except Exception as e:
        print(f"Erro ao atualizar o arquivo {caminho_arquivo}: {e}")


def extrair_numeros(texto):
    try:
        numeros = re.findall(r'\d+(?:[.,]\d+)?%?|\d+(?:[.,]\d+)?\s?[Rr]\$', texto)
        datas = re.findall(r'(fim do ano|final do semestre|primeiro semestre|trimestre|[1-9]+ a [0-9]+ meses)', texto, re.IGNORECASE)
    except:  # Captura erros de tipo e texto inválido
        numeros = ["NaN"]
        datas = ["Undefined"]
    
    return numeros, datas

def converterInteiro(numero):
    try:
        return int(numero)  
    except:  
        return 12

def apuracao(valor):
    return ""

def numero_para_coluna(num):
    coluna = ''
    while num > 0:
        num, resto = divmod(num - 1, 26)
        coluna = chr(65 + resto) + coluna
    return coluna

def buscarEntreTabelas(ValorProcurado, lista2, campo2, campoProcurado):
    
    resultado = lista2.loc[lista2[campo2] == ValorProcurado, campoProcurado]
    
    if not resultado.empty:
        return resultado.iloc[0]
    else:
        return None

def apuracaoMaM(projetado, realizado, comparacao, acomp=0):
    comparacao = ">=" if comparacao == "" else comparacao
    resultado = {"compr": comparacao, "Resultado": 0, "Status": False}

    # 🛠️ Verifica e converte valores antes da divisão
    try:
        projetado = float(projetado)
        realizado = float(realizado)
    except (ValueError, TypeError):
        resultado["Resultado"] = "Erro"
        resultado["Status"] = False
        return resultado  # Sai da função se os valores forem inválidos
    comparacao = ">=" if comparacao != comparacao  else comparacao
    match comparacao:
        case "<=":
            resultado["Resultado"] = (projetado / realizado) if realizado != 0 else float('inf')
            resultado["Status"] = resultado["Resultado"] >= 1
        case "-=":
            resultado["Resultado"] = (realizado / projetado)-1 if projetado != 0 else float('inf')
            resultado["Status"] = resultado["Resultado"] >= 1
        case ">=":
            resultado["Resultado"] = realizado / projetado if projetado != 0 else float('inf')
            resultado["Status"] = resultado["Resultado"] >= 1
        case "=":
            resultado["Resultado"] = abs(projetado - realizado)
            resultado["Resultado"] = 1 if  resultado["Resultado"] == 0 else -abs(projetado - realizado)
            resultado["Status"] = projetado == realizado
        case "<":
            resultado["Resultado"] = projetado / realizado if realizado != 0 else float('inf')
            resultado["Status"] = resultado["Resultado"] < 1
        case ">":
            resultado["Resultado"] = realizado / projetado if projetado != 0 else float('inf')
            resultado["Status"] = resultado["Resultado"] > 1
        case _:
            resultado["compr"] = ""
            resultado["Resultado"] = None
            resultado["Status"] = None
    
    resultado["Resultado"] = 1 if resultado["Resultado"] > 1 else resultado["Resultado"]
    
    return resultado

def recalcular_planilha(caminho_arquivo):
    try:
        # Abrir o Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Alterar para True se quiser ver o Excel abrindo
        excel.DisplayAlerts = False  # Evita mensagens de alerta do Excel

        # Abrir a planilha
        wb = excel.Workbooks.Open(caminho_arquivo)

        # Atualizar todas as conexões de dados (caso existam)
        wb.RefreshAll()

        # Forçar recálculo de todas as fórmulas
        excel.Calculation = -4105  # xlCalculationAutomatic
        wb.Save()

        # Garantir que todas as células sejam recalculadas
        for sheet in wb.Sheets:
            sheet.Cells.Calculate()  

        excel.CalculateFull()  # Recalcula tudo novamente para garantir
        wb.Save()
        wb.Close(SaveChanges=True)

        # Fechar o Excel completamente
        excel.Quit()

        print(f"Recalculo das fórmulas concluído para {caminho_arquivo}")

    except Exception as e:
        print(f"Erro ao recalcular a planilha {caminho_arquivo}: {e}")

def AtribuirApurar(caminho_arquivo, Base_Acomp):
    df_Base = Base_Acomp
    if isinstance(caminho_arquivo, str) and os.path.exists(caminho_arquivo):
        df_Fim = pd.read_excel(caminho_arquivo, sheet_name=None)
    else:
        print(f"Erro: '{caminho_arquivo}' não é um caminho válido!")
        return  # Para execução

    arquivoSalvo = "teste.xlsx"
    sheets_Fim = list(df_Fim.keys())
    sheets_Fim.pop(0)  # Removendo a primeira sheet (se necessário)

    for sheets in sheets_Fim:
        df_fim_Mes = df_Fim[sheets]
        for index, row in df_fim_Mes.iterrows():
            Meses_apurar = buscarEntreTabelas(row["ID"], df_Base, "ID", "Meses Acomp")
            Apurar = 0

            if Meses_apurar is None:
                Meses_apurar = []
            elif isinstance(Meses_apurar, str):
                Meses_apurar = [int(x.strip()) for x in Meses_apurar.split(',') if x.strip().isdigit()]
            elif isinstance(Meses_apurar, (int, float)):
                if not np.isnan(Meses_apurar):
                    Meses_apurar = [int(Meses_apurar)]
                else:
                    Meses_apurar = []
            elif isinstance(Meses_apurar, list):
                Meses_apurar = [int(x) for x in Meses_apurar if isinstance(x, (int, float)) and not np.isnan(x)]

            if sheets_Fim.index(sheets) + 1 in Meses_apurar:
                Apurar = 1
            elif (Meses_apurar == []):
                Apurar = 1
            df_fim_Mes.loc[index, "Apurar"] = Apurar

            print (row['ID'],sheets_Fim.index(sheets) + 1, Meses_apurar, Apurar)
        print (df_fim_Mes.head())   
        with pd.ExcelWriter(arquivoSalvo, mode='a', engine='openpyxl',if_sheet_exists='replace') as writer:
            df_fim_Mes.to_excel(writer, index=False, sheet_name=sheets)
