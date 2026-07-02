from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import time

# Crear el navegador (Chrome)
driver = webdriver.Chrome()

# Abrir una URL
driver.get("https://sapbo.marco.com.pe:40000/ControlCenter/sbo.html?saml2TokenRef=01368ad1-240c-4782-9c30-fc406baf2c6c")
time.sleep(2)

# Encontrar el select
combo = driver.find_element(By.ID, "sbo_company")
# Convertirlo en objeto Select
select = Select(combo)
select.select_by_value("SBO_MARCO_PE")

# Encontrar el input
input_usuario = driver.find_element(By.ID, "sbo_user")
# Escribir texto
input_usuario.send_keys("AsiPla06")

# Encontrar el input
input_contraseña = driver.find_element(By.ID, "sbo_password")
# Escribir texto
input_contraseña.send_keys("MP1234")

boton = driver.find_element(By.ID, "logon_sbo_btn")
boton.click()

# Cerrar navegador
driver.quit()

