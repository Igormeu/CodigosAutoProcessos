import os
import pyautogui as py
import time
import webbrowser as wb
from tkinter import *

# Variáveis
Chrome = "C://Program Files/Google/Chrome/Application/chrome.exe"
url = "https://trello.com/b/74jdupoI"
caminho = "C:\\Users\\estagiario.expansao\\Downloads"
clicar = 0

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

clicar_x,clicar_y = coordenadas_imagem('Checkbox1.png')

clicar_x -= 35

py.click(clicar_x,clicar_y)

time.sleep(2)

clicar_x,clicar_y = coordenadas_imagem('Checkbox2.png')

clicar_x -= 50

py.click(clicar_x,clicar_y)


print(clicar)



#615, 241 - Meio: X=673, Y=292
#615, 265

time.sleep(2)