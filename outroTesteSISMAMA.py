from pywinauto.application import Application
import time

# Caminho para o executável do aplicativo
caminho_programa = r'C:\datasus\SisMamaFB\SisMamaFB.exe'

# Inicialize o aplicativo
app = Application(backend="uia").start(caminho_programa)
time.sleep(5)  # Espera alguns segundos para que o aplicativo seja carregado

# Conecte-se à janela principal do aplicativo
janela_principal = app.window(title="Título da Janela")

# Acesse o menu principal
menu_principal = janela_principal.child_window(title="Menu Principal", control_type="MenuBar")
menu_principal.click_input()

# Acesse o submenu
submenu = janela_principal.child_window(title="Submenu", control_type="MenuItem")
submenu.click_input()

# Realize outras interações necessárias no submenu

# Feche o aplicativo quando terminar
app.kill()
