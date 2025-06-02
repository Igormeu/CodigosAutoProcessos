from modulos import modulos as md
import os
import time
import pyautogui as py
import pyperclip
import pandas as pd
from datetime import datetime, timedelta
import subprocess

# # Diretório padrão para downloads
# novo_caminho = os.path.join(os.path.expanduser("~"), "Downloads")

# # Obtém a data do dia anterior e gera o caminho da planilha
# def obter_caminho_planilha():
#     data_anterior = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
#     nome_arquivo = f"{data_anterior}.xlsx"
#     caminho = os.path.join(novo_caminho, nome_arquivo)
#     return caminho

# # Caminho para a planilha atualizada
# caminho_planilha = obter_caminho_planilha()

# # Verifica a existência da planilha
# if not os.path.exists(caminho_planilha):
#     print(f"Planilha atualizada não encontrada: {caminho_planilha}")
#     exit()

# # Carrega a planilha atualizada
# print(f"📂 Carregando planilha: {caminho_planilha}")
# df = pd.read_excel(caminho_planilha, engine='openpyxl')

# Configurações do SAP
sap_path = r"C:\Program Files\SAP\SAP Business One\SAP Business One.exe"

# Inicia o SAP Business One
subprocess.Popen(sap_path)

# Aguardar o SAP Business One abrir
time.sleep(20)  # Ajuste o tempo, se necessário



# Ativar addon
try:
    # Inserir o usuário e a senha
    usuario = ""
    senha = ""

    # Digitar o usuário
    py.write(usuario)
    py.press('TAB')
    time.sleep(2)
    py.write(senha)
    py.press('enter')  

    time.sleep(30)  

    #Remover Por-Up
    x,y = md.coordenadas_imagem("imgs/FecharWelcome.png", 10)
    x = x+20
    coordenadas = x,y
    py.click(coordenadas)

    time.sleep(2)  

    x,y = md.coordenadas_imagem("imgs/FecharFiliais.png", 10)
    x = x+20
    coordenadas = x,y
    py.click(coordenadas)

    time.sleep(2) 

    #Ir até a aba do Addon
    py.keyDown('alt')
    py.press('m')
    py.press('a')
    py.press('a')
    py.press('a')
    py.keyUp('alt')

    #Ativar Addon
    coordenadas = md.coordenadas_imagem("imgs/LinhaDoAddon.png", 10)
    py.click(coordenadas,clicks=2)

    #Iniciar Addon
    x,y = md.coordenadas_imagem("imgs/LigareSairdaTela.png", 10)
    x = x+50
    coordenadas = x,y
    py.click(coordenadas)

    time.sleep(2)

    x,y = md.coordenadas_imagem("imgs/LigareSairdaTela.png", 10)
    x = x-50
    coordenadas = x,y

    py.click(coordenadas)

    time.sleep(2)

except Exception:
    print("Erro ao executar o RPA")



# py.click(x=1083, y=238)
# time.sleep(2)
# py.click(x=1723, y=930)
# time.sleep(90)

# # Vai para o BankPlus
# py.click(x=416, y=15)
# for _ in range(16):
#     py.press('down')
#     time.sleep(0.1)  # Pequena pausa para garantir execução correta
# py.press('enter')