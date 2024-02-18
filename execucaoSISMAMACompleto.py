import subprocess
from pywinauto.application import Application

def abrir_programa_e_navegar(caminho_programa):
    try:
        # Abrir o programa usando subprocess com privilégios elevados
        subprocess.Popen(['runas', '/user:EDUARDA', caminho_programa], shell=True)
        print(f'O programa {caminho_programa} foi aberto como administrador com sucesso!')

        # Conectar-se ao aplicativo usando pywinauto
        app = Application(backend="uia").connect(path=caminho_programa)

        # Referenciar a janela principal do programa
        janela_principal = app.window()

        # Exemplo de interação: clique em um menu chamado "Cadastro"
        # Substitua 'Cadastro' com o nome exato do seu menu
        menu_cadastro = janela_principal.child_window(title="Cadastro", control_type="MenuItem")
        menu_cadastro.click_input()

        # Adicione mais interações conforme necessário

    except FileNotFoundError:
        print(f'O programa {caminho_programa} não foi encontrado.')
    except Exception as e:
        print(f'Ocorreu um erro ao tentar abrir o programa {caminho_programa} como administrador: {e}')

if __name__ == "__main__":
    caminho_programa = r'C:\datasus\SisMamaFB\SisMamaFB.exe'
    abrir_programa_e_navegar(caminho_programa)
12
