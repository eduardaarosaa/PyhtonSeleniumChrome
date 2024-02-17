import subprocess

def abrir_programa(caminho_programa):
    try:
        subprocess.Popen(caminho_programa)
        print(f'O programa {caminho_programa} foi aberto com sucesso!')
    except FileNotFoundError:
        print(f'O programa {caminho_programa} não foi encontrado.')
    except Exception as e:
        print(f'Ocorreu um erro ao tentar abrir o programa {caminho_programa}: {e}')

if __name__ == "__main__":
    caminho_programa = r'C:\Windows\system32\mspaint.exe'
    abrir_programa(caminho_programa)
