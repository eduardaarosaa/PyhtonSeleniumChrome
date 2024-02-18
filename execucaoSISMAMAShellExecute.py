import ctypes
import sys

# Define constantes para a função ShellExecute
SEE_MASK_NOCLOSEPROCESS = 0x00000040
SW_SHOWNORMAL = 1

def run_as_admin(program):
    try:
        # Chama a função ShellExecute para executar o programa como administrador
        ctypes.windll.shell32.ShellExecuteW(None, "runas", program, None, None, SW_SHOWNORMAL)
    except Exception as e:
        print(f"Erro ao executar como administrador: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # Caminho para o programa que você deseja abrir como administrador
    caminho_programa = r'C:\datasus\SisMamaFB\SisMamaFB.exe'
    run_as_admin(caminho_programa)
