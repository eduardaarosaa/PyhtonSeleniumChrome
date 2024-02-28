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
      
    except Exception as e:
        print(f"Erro ao selecionar o item seguimento do menu: {e}")    

def novo_cadastro_paciente():
    
    try:
        pyautogui.moveTo(x=1017,y=89)
        time.sleep(3)
        pyautogui.click()
        time.sleep(5)
        
        print('aquiiiiiiiii')
        
        posicao = pyautogui.position()
        print(posicao)
        
    except Exception as e:
        print(f"Erro ao selecionar o item novo em cadastro de paciente: {e}")  
          
def excel_dados_paciente():
    try: 
        caminho_arquivo_excel = 'Sismama.csv'
        dados_excel = pd.read_csv(caminho_arquivo_excel)
        print(dados_excel.head())
        return dados_excel.head()
    
    except Exception as e: 
        print(f"Erro ao procurar o Excel: {e}")  

def dados_processamento_excel(dados_excel):
    try:
       time.sleep(3)
       cadastro_dados_paciente(dados_excel)
       cadastro_dados_residenciais(dados_excel)
       cadastro_dados_risco_elevado(dados_excel)
       cadastro_aba_mamografia(dados_excel)
      
    except Exception as e: 
        print(f"Erro ao processar os dados do excel: {e}")  
        
def tratativa_risco(risco):
    try:
        #não há risco 
        if risco == 'não':  
            pyautogui.moveTo( 526,372)
            pyautogui.click()
        elif risco == 'n/s':
            #não sei
             pyautogui.moveTo(705,369)
             pyautogui.click()
        else:
             #Tem risco - SIM
             pyautogui.moveTo(340,371)
             pyautogui.click()
             
        time.sleep(0.5)

    except Exception as e: 
        print(f"Erro ao processar os dados do excel - Risco elevado para câncer de mama: {e}")  
        
def cadastro_dados_paciente(dados_excel):
    try: 
        #campos
        cartao_sus = dados_excel.loc[0,'cartao sus']
        nome = dados_excel.loc[0, 'nome']
        sexo = dados_excel.loc[0, 'sexo']
        nome_mae = dados_excel.loc[0, 'mae']
        orgao_emissor = dados_excel.loc[0, 'orgao emissor']
        identidade = dados_excel.loc[0, 'identidade']
        cpf = dados_excel.loc[0, 'cpf']
        idade = dados_excel.loc[0, 'idade']
        nacionalidade = dados_excel.loc[0, 'nacionalidade']
        raca_cor = dados_excel.loc[0, 'raça']
        
        #cartao sus
        pyautogui.click(418,100)
        pyautogui.write(str(cartao_sus), interval=0.1)
        
        time.sleep(5)

        posicao = pyautogui.position()
        print(posicao)
        
        time.sleep(0.5)
        
        #nome
        pyautogui.click(581,119)
        pyautogui.write(nome, interval=0.1)
        
        time.sleep(0.5)
        
        tratativa_sexo(sexo)
        
        time.sleep(0.5)
        
        #nome_mae
        pyautogui.click(463,149)
        pyautogui.write(nome_mae, interval=0.1)
        
        time.sleep(0.5)
    
        #orgao_emissor
        pyautogui.click(611,168) #confirmar a coordenada
        pyautogui.write(orgao_emissor, interval=0.1) 
        
        time.sleep(0.5)
        
        #identidade
        pyautogui.click(440,165)
        pyautogui.write(str(identidade), interval=0.1) 
        
        time.sleep(0.5)
        
        #cpf
        pyautogui.click(421,187)
        pyautogui.write(str(cpf), interval=0.1) 
        
        time.sleep(3)
    
        #idade
        pyautogui.click(699,191)
        pyautogui.write(str(idade), interval=0.1) 
        
        time.sleep(0.5)
        
        #nacionalidade
        pyautogui.click(566,214) # SELECT VER COMO CLICA DE FORMA DINAMICA
        pyautogui.write(str(nacionalidade), interval=0.1) 
        pyautogui.press('enter')
        
        #raca_cor 
        tratativa_raca(raca_cor)
  
        
        
    except Exception as e: 
        print(f"Erro ao processar os dados do paciente no novo cadastro: {e}")  
        
def cadastro_dados_residenciais(dados_excel):
    try: 
        #campos
        endereco = dados_excel.loc[0, 'endereço']
        numero = dados_excel.loc[0, 'numero']
        uf = dados_excel.loc[0, 'uf']
        municipio = dados_excel.loc[0, 'municipio']
        telefone = dados_excel.loc[0, 'telefone']
        
        time.sleep(0.5)
        
        #endereco
        pyautogui.click(429,268)
        pyautogui.write(endereco, interval=0.1)
        
        time.sleep(0.5)
 
        #numero
        pyautogui.click(753,269)
        pyautogui.write(str(numero), interval=0.1)
        
        time.sleep(0.5)
        
        #uf 
        pyautogui.click(849,298)
        pyautogui.write(uf, interval=0.1)
        pyautogui.press('enter') 
        
        time.sleep(0.5)
        
        #municipio 
        pyautogui.click(472,310)
        pyautogui.write(municipio, interval=0.1)
        pyautogui.press('enter') 
        
        #telefone
        pyautogui.click(787,315) #confirmar coordenada
        pyautogui.write(telefone, interval=0.1)
    
    except Exception as e: 
        print(f"Erro ao processar os dados de endereço no novo cadastro do paciente: {e}")  
        
def cadastro_dados_risco_elevado(dados_excel):
    try: 
        # Campos
        risco_elevado = dados_excel.loc[0, 'risco']
        
        tratativa_risco(risco_elevado)

    except Exception as e: 
        print(f"Erro ao processar os dados de risco elevado no novo cadastro do paciente: {e}")

        
def tratativa_sexo(sexo):
    try: 
        #Feminino
        if sexo == 'F' :
            pyautogui.moveTo(716,100)
            pyautogui.click() 
        else:
            #masculino
            pyautogui.moveTo(640,100)
            pyautogui.click() 
    except Exception as e: 
        print(f"Erro ao tratar os dados do paciente no campo sexo {e}")  
        
def tratativa_raca(raca_cor):
    try: 
        if raca_cor == 'branca' :
            pyautogui.moveTo(815,175)
            pyautogui.click() 
        elif raca_cor == 'preta':
            pyautogui.moveTo(813,202)
            pyautogui.click() 
        elif raca_cor == 'parda':
            pyautogui.moveTo(812,225)
            pyautogui.click() 
        elif raca_cor == 'amarela':
            pyautogui.moveTo(891,177)
            pyautogui.click() 
        else :
            #indigena
            pyautogui.moveTo(892,199)
            pyautogui.click() 
    except Exception as e: 
        print(f"Erro ao tratar os dados do paciente no campo: sexo {e}")
        
import pyautogui

def cadastro_aba_mamografia(dados_excel):
    try:
        # Campos
        cnes = dados_excel.loc[0, 'cnes']
        servico_radiologico = dados_excel.loc[0, 'serviço radiologico']
        localizacao_mama = dados_excel.loc[0, 'localizacao mama']
        tipo_mamografia = dados_excel.loc[0, 'tipo mamografia']
        numero_exame = dados_excel.loc[0, 'numero exame']
        mama_direita = dados_excel.loc[0, 'mama direita']
        mama_esquerda = dados_excel.loc[0, 'mama esquerda']
        categoria_direita = dados_excel.loc[0, 'categoria direita']
        categoria_esquerda = dados_excel.loc[0, 'categoria esquerda']
        
        #CNES
        pyautogui.click(408, 447)
        pyautogui.write(str(cnes), interval=0.1)
        pyautogui.press('enter') 
          
        # Serviço radiologico
        pyautogui.click(936, 438)
        pyautogui.click(936, 438)
        # pyautogui.write(str(servico_radiologico), interval=0.1)
        # pyautogui.press('enter') 
        
        tratativa_localizacao_mamografia(localizacao_mama)
        tratativa_tipo_mamografia(tipo_mamografia)
        
        # número do exame 
        pyautogui.click(921,493) 
        pyautogui.write(str(numero_exame), interval=0.1)
        
        # mama direita
        pyautogui.click(450, 532)
        pyautogui.write(str(mama_direita), interval=0.1)
        pyautogui.press('enter') 
        
        # mama esquerda
        pyautogui.click(797, 536) 
        pyautogui.write(str(mama_esquerda), interval=0.1)
        pyautogui.press('enter') 
        
        tratativa_categoria_mamografia_direta(categoria_direita)
        tratativa_categoria_mamografia_esquerda(categoria_esquerda)
  
        
    
    except Exception as e: 
        print(f"Erro ao tratar os dados do paciente - Aba mamografia {e}")
        
def tratativa_localizacao_mamografia(localizacao_mama):
    try:
        if(localizacao_mama == 'direita'):
            pyautogui.moveTo(348,494)
            pyautogui.click()
        elif(localizacao_mama == 'esquerda'):
             pyautogui.moveTo(449,501)
             pyautogui.click()
        else:
            #ambas
             pyautogui.moveTo(547,497)
             pyautogui.click()
    
    except Exception as e: 
        print(f"Erro ao tratar os dados do paciente - Aba mamografia {e}")
        
def tratativa_tipo_mamografia(tipo):
    try:
        if(tipo == 'diagnostica'):
            pyautogui.moveTo(667,495)
            pyautogui.click()
        elif(tipo == 'rastreamento'):
             pyautogui.moveTo(755,498)
             pyautogui.click()
    
    except Exception as e: 
        print(f"Erro ao tratar os dados do paciente - Aba mamografia (Tipo mamografia) {e}")

def tratativa_categoria_mamografia_direta(categoria):
    try:
        if(categoria == 0):
            pyautogui.moveTo(445,592)
            pyautogui.click()
        elif(categoria == 1):
             pyautogui.moveTo(443,610)
             pyautogui.click()
        elif(categoria == 2):
             pyautogui.moveTo(446,626)
             pyautogui.click()
        elif(categoria == 3):
             pyautogui.moveTo(444,643)
             pyautogui.click()
        elif(categoria == 4):
             pyautogui.moveTo(445,657)
             pyautogui.click()
        elif(categoria == 5):
             pyautogui.moveTo(445,675)
             pyautogui.click()
        else:
             pyautogui.moveTo(443,688)
             pyautogui.click()
    
    except Exception as e: 
        print(f"Erro ao tratar os dados do paciente - Aba mamografia (Categoria Mama direita) {e}")


def tratativa_categoria_mamografia_esquerda(categoria):
    try:
        if(categoria == 0):
            pyautogui.moveTo(854,594)
            pyautogui.click()
        elif(categoria == 1):
             pyautogui.moveTo(854,609)
             pyautogui.click()
        elif(categoria == 2):
             pyautogui.moveTo(856,626)
             pyautogui.click()
        elif(categoria == 3):
             pyautogui.moveTo(854,640)
             pyautogui.click()
        elif(categoria == 4):
             pyautogui.moveTo(853,661)
             pyautogui.click()
        elif(categoria == 5):
             pyautogui.moveTo(853,674)
             pyautogui.click()
        else:
             pyautogui.moveTo(853,691)
             pyautogui.click()
    
    except Exception as e: 
        print(f"Erro ao tratar os dados do paciente - Aba mamografia (Categoria Mama esquerda) {e}")
        
        
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
    menu_item_seguimento()
    novo_cadastro_paciente()
    dados_excel = excel_dados_paciente()
    dados_processamento_excel(dados_excel)
    
    
