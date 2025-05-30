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

def buscarEntreTabelas(ValorProcurado, lista2, campo2, campoProcurado,tableName,mesNum, Proje_Reali):
    
    # Proje_Reali : 0 Projetado e 1 realizado

    #Atualizar todo mês, ou fazer uma tablea com esses dados

    idsLucro = (129, 127, 131, 134, 136, 138, 141, 143)
    idsperLucro = (130, 128, 132, 133, 135, 137, 140, 142)

    PvalueLucro = (1580000,1430000, 1720000,0,0,0,0,0,0,0,0,0)
    Pvaloresper = (0.0704, 0.0053, 0.00752,0,0,0,0,0,0,0,0,0)

    RvalueLucro = (-80000,-1132251.58, 2320000,0,0,0,0,0,0,0,0,0)
    Rvaloresper = (-0.039, -0.0649, 0.1125,0,0,0,0,0,0,0,0,0)

    if ValorProcurado in idsLucro and tableName == "OKR":
        base = PvalueLucro if Proje_Reali == 0 else RvalueLucro
        return base[mesNum-1]

    elif ValorProcurado in idsperLucro and tableName == "OKR":
        base = Pvaloresper if Proje_Reali == 0 else Rvaloresper
        return base[mesNum-1]

    else:
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
        realizado = float(realizado) if realizado is not None else None
    except (ValueError, TypeError):
        resultado["Resultado"] = "Erro"
        resultado["Status"] = False
        return resultado  # Sai da função se os valores forem inválidos
    comparacao = ">=" if comparacao != comparacao  else comparacao

    if realizado is None:
        resultado["Resultado"] = 0
        resultado["Status"] = None
    else:
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

    resultado["Resultado"] = 0 if resultado["Resultado"] < 0 else resultado["Resultado"]
    
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

def enviarEmails (mesRef, mesNome,diretoria):

    # Cria instância do Outlook
    caminho_exec = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    # Configura o e-mail
    match (diretoria):
        case "Industria":
            remetente = "manoel.pontes@frosty.ind.br; jose.carlos@frosty.ind.br"
        case "Comercial":
            remetente = "andressa.silva@frosty.ind.br ;andressa.silva@frosty.ind.br; germano.batista@frosty.ind.br"
        case "Finanças":
            remetente = "akyme.silva@frosty.ind.br; carlos.souza@frosty.ind.br; marcia.lima@frosty.ind.br; gerson.pacheco@frosty.ind.br"
        case "Recursos Humanos":
            remetente = "wilderlandia.uchoa@frosty.ind.br; katiane.silva@frosty.ind.br"
        case "Supply":
            remetente = "rodrigo.miranda@frosty.ind.br; alexandre.almeida@frosty.ind.br; filipe.romao@frosty.ind.br; leonara.martins@frosty.ind.br"
        case "Tecnologia da Informação":
            remetente = "wagner.lima@frosty.ind.br; audizio.filho@frosty.ind.br; odelly.alves@frosty.ind.br"
        case "Marketing":
            remetente = "susane.mamede@frosty.ind.br; christian.borges@frosty.ind.br"
        case "Expansão&Operação":
            remetente = "abel.lucas@frosty.ind.br; fauber.oliveira@frosty.ind.br"
        case _:
            remetente = "igor.santos@frosty.ind.br"

    mail.To = remetente
    # mail.CC = 'lucimario.braz@frosty.ind.br; edgard@frosty.ind.br' if not remetente == "" else ""
    mail.Subject = f"Modelo de OKR e KPI {diretoria} - {mesRef}"
    mail.Importance = 2
    mail.HTMLBody = f"""
    <html>
    <body style="font-family: Arial, sans-serif; font-size: 11pt; color: #000000;">
        <p>Saudações,</p>

        <p>Conforme alinhado desde janeiro, estou enviando a planilha modelo de <b>OKR</b> e <b>KPI</b> para serem preenchidas com as informações referentes ao mês de <b>{mesNome}</b>.</p>

        <p><b>Orientações gerais:</b></p>
        <ol>
        <li><b>Divergências nos valores:</b> Caso observe alguma diferença entre o valor projetado e o valor presente na planilha, por favor, revise a planilha “Metas Gerais” que foi enviada ao setor.</li>
        <li><b>Inclusão de novos OKR's ou KPI's:</b> Se houver algum OKR ou KPI que não conste na lista, solicito que <u>não os insira diretamente nesta planilha</u>. Em vez disso, me avise pessoalmente ou responda a este e-mail.</li>
        </ol>

        <p>A apuração dos resultados do mês anterior já foi realizada e está disponível no <b>BI</b>, na seção 
        <span style="color: #1F4E79;"><a href="https://app.powerbi.com/groups/c3290536-fb7c-4de2-b21f-9d74b57e4d40/reports/d1912d0d-5688-4d1f-a89b-ad3ae11ae340/cdd89694795de4e819e9?experience=power-bi"> "FIN - ReportStatusKeyResults - Oficial"</a></span>, a qual, até o momento do envio deste e-mail, todos os diretores devem ter acesso. Caso algum diretor ainda não tenha acesso, por favor, entre em contato comigo.</p>

        <p><b>Nota:</b> A apuração de Abril ocorreu de forma parcial, dado que algumas diretóris não foram capazes de fornecer os dados de apuação em tempo hábil, denotando em uma gap na coleta dos dados com a virada de sistema.<br>
        Peço que as diretorias que não enviaram seus indicadores por este motivo ou quaalquer outro que tenha ocorrido, envie em conjunto a planilha de Maio a Abril para que possamos ter um acompanhamento real das métricas da organização.</p>

        <p><b>Prazo de entrega:</b> O prazo final para o envio dos resultados é <span style="color: red;"><b>06/06</b></span>.</p>

        <p>Estou à disposição para esclarecer qualquer dúvida que possa surgir sobre o tema.</p>

        <p size="10">Atenciosamente,</p>

        <p size="10"><b>Igor Stenio</b><br>
        Auxiliar de Processos</p>
    </body>
    </html>
    """

    # Adiciona anexo
    caminho_anexo = os.path.join(caminho_exec,f"03. templates/03. realizado_mam/{mesRef}/OKR_KPI - {diretoria}_{mesRef}.xlsx")

    print (caminho_anexo)
    if os.path.exists(caminho_anexo):
        mail.Attachments.Add(caminho_anexo)
        # Envia o e-mail
        mail.Send()
        print('E-mail enviado com sucesso!')
    else:
        print('Anexo não encontrado!')

        