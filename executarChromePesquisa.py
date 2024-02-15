from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# Inicializa o driver do Chrome
driver = webdriver.Chrome()

# Abre o Google
driver.get("https://www.google.com")

# Localiza o campo de pesquisa pelo seu ID
campo_pesquisa = driver.find_element("name","q")

# Digita "bolo de chocolate" no campo de pesquisa
campo_pesquisa.send_keys("bolo de chocolate")

# Espera um segundo para a página responder
time.sleep(1)

# Pressiona Enter para realizar a pesquisa
campo_pesquisa.send_keys(Keys.RETURN)

# Espera um pouco para que os resultados da pesquisa apareçam
time.sleep(5)