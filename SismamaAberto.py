import time
from pyautogui import moveTo, click, position
from pywinauto.application import Application
import pyautogui

def clicarMenu():
    
    time.sleep(5)
        
app = Application().connect(title="Coordenação Municipal - v.4.18")

# Selecione a janela principal do programa
janela_principal = app.window(title="Coordenação Municipal - v.4.18")

# Clique no menu "Cadastro"
janela_principal.menu_select("Cadastro")
    
    
    # menu = pyautogui.position()
    # pyautogui.moveTo(-646, 267, duration=1)
    # time.sleep(5)
    # pyautogui.doubleClick()
    # pyautogui.click()
    # print('AQUI')
    # print(menu);
    # x_input, y_input = -1100, 195
    # menu = pyautogui.position()
    # print(menu);
    # pyautogui.click(x_input, y_input)

    # # Digita o texto desejado
    # texto = "Olá, mundo!"
    # pyautogui.typewrite(texto)

clicarMenu()
   