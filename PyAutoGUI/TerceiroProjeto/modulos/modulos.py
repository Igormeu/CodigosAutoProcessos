import os
import time
import pyautogui as py
import pyperclip
import pandas as pd
from datetime import datetime, timedelta
import subprocess


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