import os
import pyautogui as py
import time
import webbrowser as wb
from tkinter import *

# Variáveis
Chrome = "C://Program Files/Google/Chrome/Application/chrome.exe"
url = "https://trello.com/b/74jdupoI"

#Função para exibir um popup

def popup():
    window = Tk()
    window.title("Alert")
    
    # Dimensões da janela
    window_width = 300
    window_height = 150
    
    # Calcula a posição para centralizar a janela
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    
    # Define a geometria da janela
    window.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")
    
    # Melhorar a interface
    mensagem = Label(window, text="A atualização do BI falhou,\n tente novamente", font=("Arial", 12), pady=20)
    mensagem.pack()
    
    button = Button(window, text="OK", command=window.destroy, font=("Arial", 10), padx=20, pady=5)
    button.pack(pady=10)
    
    window.mainloop()


#Função para fazer o reconhecimento das imgs
def coordenadas_imagem(nome_img,max_tentativas=5):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    url_img1 = os.path.join(script_dir, nome_img)

    try:
        img = py.locateCenterOnScreen(url_img1, confidence=0.9)
    except py.ImageNotFoundException:
        img = None

    if img is not None:
        print("Imagem encontrada:", img)
        return img.x, img.y
    else:
        tentativas = 0 
        while img is None and tentativas < max_tentativas:
            try:
                print("Imagem não encontrada, tentando novamente...")
                img = py.locateCenterOnScreen(url_img1, confidence=0.9)
            except py.ImageNotFoundException:
                img = None
            time.sleep(1)
            tentativas += 1
        if img is not None:
            print("Imagem encontrada:", img)
            return img.x, img.y
        else:
            # popup()
            return None

# Abrindo o Chrome
wb.get(Chrome + ' %s').open(url)

# Esperar até que o navegador esteja completamente carregado
time.sleep(20)

#verificar se a tela está maximizada

maximizar_tela = coordenadas_imagem('maximizar.png',1)

if not(maximizar_tela == None):
    
    py.click(maximizar_tela)

#Evitar bug do chrome

ocorreu_bug = coordenadas_imagem('menu_aberto.png',1)

if ocorreu_bug == None:
    #Abrir menu
    py.press('w')

time.sleep(2)

# Rolar a barra lateral

mover = coordenadas_imagem('Barra_rolagem.png')
py.moveTo(mover)
py.scroll(-300)

time.sleep(2)
# Clicar no botão desejado
clicar = coordenadas_imagem('Imprimir.png')
py.click(clicar)

time.sleep(2)
# Ativar a extensão trelloexport
clicar = coordenadas_imagem('trelloexport.png')
py.click(clicar)

time.sleep(2)
# Marcar os dois checkbox

clicar_x,clicar_y = coordenadas_imagem('Checkbox1.png')

clicar_x -= 35

py.click(clicar_x,clicar_y)

time.sleep(2)

clicar_x,clicar_y = coordenadas_imagem('Checkbox2.png')

clicar_x -= 50

py.click(clicar_x,clicar_y)

time.sleep(2)
# Selecionar os quadros a serem exportados

py.scroll(-200)

py.keyDown('ctrl')

time.sleep(2)

clicar = coordenadas_imagem('borda_process.png')
py.click(clicar)

time.sleep(2)

clicar = coordenadas_imagem('borda_legalizacion.png')
py.click(clicar)

# Clicar em exportar
clicar = coordenadas_imagem('botton_export.png')
py.click(clicar)

resposta_busca = coordenadas_imagem("Sucessfull.png")

#Ver se o export foi rrealizado com sucesso
while (resposta_busca == None):
    time.sleep(15)
    resposta_busca = coordenadas_imagem("Sucessfull.png")

# Sair do Chrome
py.hotkey("ctrl", "shift", "w")
