import ctypes
import time
from pyautogui import moveTo, click, position
import pyautogui


def procurar_sismama():
    time.sleep(3)
    postion_app_fixo = pyautogui.position()
    print(postion_app_fixo)
    
    # pyautogui.moveTo(-1109,303)
    # pyautogui.click()
    
    # pyautogui.moveTo(-1077,354);
    # pyautogui.click();
    
    menu_principal_position = pyautogui.locateOnScreen('menuCadastro.png')
    
    if menu_principal_position is not None:
        pyautogui.click(menu_principal_position)

        # Aguarde um pequeno período de tempo para o submenu abrir
        time.sleep(1)
    else:
        print("Menu principal não encontrado.")
    
    
procurar_sismama(); 