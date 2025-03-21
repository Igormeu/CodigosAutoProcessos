import webbrowser as wb
import pyautogui as py
import os
import time
from tkinter import *

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


def coordenadas_imagem(nome_img):
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
        max_tentativas = 10  # Limite de tentativas para evitar loop infinito
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
            popup()
            return None

# Variáveis
Chrome = "C://Program Files/Google/Chrome/Application/chrome.exe"
url = "https://trello.com/b/74jdupoI"
caminho = "C:\\Users\\estagiario.expansao\\Downloads"

# Abrindo o Chrome

wb.get(Chrome + ' %s').open(url)

time.sleep(20)

#clicar nos três pontinhos

py.click(1342, 157)

time.sleep(2)

#Rolar a barra lateral

py.moveTo(coordenadas_imagem("tres_pontinhos.png"))

py.scroll(-300)

time.sleep(2)

py.click(1180, 669)

time.sleep(5)

#Ativar a extensão trelloexport

py.click(1161, 599)

time.sleep(2)

#Marcar os dois checkbox

py.click(615, 241)

py.click(615, 265)

#Selecionar os quadros a serem exportados

py.scroll(-200)

time.sleep(2)

py.click(665,330)

time.sleep(2)

py.hold("ctrl")

py.click(678,382)

py.click(718,405)

#Clicar em exportar

py.click(899, 687)

time.sleep(120)

#Sair so chrome

py.hotkey("ctrl","shift","w")


