import ctypes
import time
from pyautogui import moveTo, click, position
import pyautogui

# Constantes da API do Windows
SW_RESTORE = 9

def run_as_admin(program):
    try:
        # Chama a função ShellExecute para executar o programa como administrador
        ctypes.windll.shell32.ShellExecuteW(None, "runas", program, None, None, 1)
        time.sleep(5);
    except Exception as e:
        print(f"Erro ao executar como administrador: {e}")

def find_window(window_title):
    # Encontra a janela principal do executável pelo título da janela
    hwnd = ctypes.windll.user32.FindWindowW(None, window_title)
    return hwnd

def menu_item_position():
    positionMenu = pyautogui.position()
    print('Chegou na função de item position.'); 
    pyautogui.moveTo(-1063,635)
    pyautogui.click();

if __name__ == "__main__":
    # Caminho para o programa que você deseja abrir como administrador
    caminho_programa = r'C:\datasus\SisMamaFB\SisMamaFB.exe'
    
    # Título da janela do aplicativo
    window_title = "Coordenação Municipal - v.4.18"  # Substitua pelo título da janela do seu aplicativo
    
    # Localizando a janela principal do aplicativo
    hwnd = find_window(window_title)

    if hwnd:
        print("Janela encontrada!")

        # Exibir a janela
        ctypes.windll.user32.ShowWindow(hwnd, SW_RESTORE)


    else:
        print("Janela não encontrada.")

    # Executando o programa como administrador
    run_as_admin(caminho_programa)
    menu_item_position()
    
    
