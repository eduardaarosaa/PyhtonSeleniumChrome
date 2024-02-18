import ctypes
import time
from pywinauto import Desktop

# Constantes da API do Windows
WM_COMMAND = 0x0111

def send_menu_command(hwnd, menu_id):
    # Enviar uma mensagem WM_COMMAND para selecionar um item de menu
    command = (menu_id & 0xFFF) | (1 << 16)  # Construir o comando de mensagem WM_COMMAND
    ctypes.windll.user32.SendMessageW(hwnd, WM_COMMAND, command, 0)

if __name__ == "__main__":
    # Obtendo o identificador da janela (HWND)
    hwnd = Desktop().window(title="Coordenação Municipal - v.4.18").handle

    if hwnd:
        # Enviando o comando para selecionar o item de menu com ID 224
        menu_id = 524

        # Enviando o comando para a janela
        send_menu_command(hwnd, menu_id)

        # Aguardando um curto período de tempo para que a ação do menu seja realizada
        time.sleep(0.5)
    else:
        print("Janela não encontrada.")
