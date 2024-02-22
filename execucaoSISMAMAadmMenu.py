import ctypes
import time
from pyautogui import moveTo, click, position
import pyautogui
from pywinauto.application import Application
import pandas as pd 

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
    try:        
            time.sleep(1)
            pyautogui.moveTo(x=368,y=166)
            time.sleep(1)
            pyautogui.click()
            
    except Exception as e:
            print(f"Erro ao selecionar o menu: {e}")
            
def menu_item_seguimento():
    
    try:
        pyautogui.moveTo(x=397,y=216)
        time.sleep(1)
        pyautogui.click()
        time.sleep(5)
        menu_item = pyautogui.position();
        print(menu_item);
            
    except Exception as e:
        print(f"Erro ao selecionar o item seguimento do menu: {e}")    

def novo_cadastro_paciente():
    
    try:
        pyautogui.moveTo(x=1017,y=89)
        time.sleep(1)
        pyautogui.click()
            
    except Exception as e:
        print(f"Erro ao selecionar o item novo em cadastro de paciente: {e}")  
          
def excel_dados_paciente():
    try: 
        caminho_arquivo_excel = 'Sismama.xlsx'
        dados_excel = pd.read_excel(caminho_arquivo_excel)
        # print(dados_excel.head())
        return dados_excel.head()
    
    except Exception as e: 
        print(f"Erro ao procurar o Excel: {e}")  

def dados_processamento_excel(dados_excel):
    try:
       cadastro_dados_paciente(dados_excel)
       cadastro_dados_residenciais(dados_excel)
       cadastro_dados_risco_elevado(dados_excel)
      
    except Exception as e: 
        print(f"Erro ao processar os dados do excel: {e}")  
        
def cadastro_dados_paciente():
    try: 
        #campos
        cartao_sus = dados_excel['cartao sus']
        nome = dados_excel['nome']
        sexo = dados_excel['sexo']
        mae = dados_excel['mae']
        identidade = dados_excel['identidade']
        cpf = dados_excel['cpf']
        idade = dados_excel['idade']
    
    except Exception as e: 
        print(f"Erro ao processar os dados do paciente no novo cadastro: {e}")  
        
def cadastro_dados_residenciais():
    try: 
        #campos
        endereco = dados_excel['endereço']
        numero = dados_excel['numero']
        uf = dados_excel['uf']
        municipio = dados_excel['municipio']
    
    except Exception as e: 
        print(f"Erro ao processar os dados de endereço no novo cadastro do paciente: {e}")  
        
def cadastro_dados_risco_elevado():
    try: 
        #campos
          risco_elevado = dados_excel['risco']
    
    except Exception as e: 
        print(f"Erro ao processar os dados de risco elevado no novo cadastro do paciente: {e}")  
   
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
    # menu_item_position()
    # menu_item_seguimento()
    # novo_cadastro_paciente()
    dados_excel = excel_dados_paciente()
    dados_processamento_excel(dados_excel)
    
    
