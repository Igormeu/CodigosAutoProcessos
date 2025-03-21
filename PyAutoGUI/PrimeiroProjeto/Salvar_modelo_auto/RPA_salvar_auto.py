import os
import pyautogui as py
import time
from tkinter import *

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
            return None

def popup(mensagem):
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
    mensagem = Label(window, text=f"{mensagem}", font=("Arial", 12), pady=20)
    mensagem.pack()
    
    button = Button(window, text="Ligar novamente", command=executar(), font=("Arial", 10), padx=20, pady=5)
    button.pack(pady=10)
    
    button = Button(window, text="Desligar", command=window.destroy, font=("Arial", 10), padx=20, pady=5)
    button.pack(pady=10)
    
    window.mainloop()

def executar ():
    mensagem = "O código não encontrou a interface do bizagi\nO que você gostria de fazer ?"
     
    a = coordenadas_imagem("interface.png",3)
    
    while a is not None:
        time.sleep(60)
        clicar = coordenadas_imagem("salvar.png",2)
        py.click(clicar)
        a = coordenadas_imagem("interface.png",3)   
    popup(mensagem)

executar()