from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd
import sys

def es_numero_valido(numero):
    numero = numero.strip()
    return numero.isdigit() and len(numero) == 9

try:
    # Lee el archivo Excel
    df = pd.read_excel('C:/Omar/Personal/Infopucp/Código/Automatización de mensajes/FichaInscripcion.xls', usecols="Z", skiprows=9)
    dm = pd.read_excel('C:/Omar/Personal/Infopucp/Código/Automatización de mensajes/FichaInscripcion.xls', usecols="C", skiprows=9)
    assert df is not None, "No se pudo leer el archivo Excel."
    
    # Procesa los números de teléfono
    numeros = df.iloc[:, 0].dropna().astype(str).apply(lambda x: x.split('.')[0]).tolist()
    numeros_validos = [nuevoNumero for nuevoNumero in numeros if es_numero_valido(nuevoNumero)]
    # messages = dm.iloc[:, 0].dropna().astype(str).tolist()  

    rowNumbers = len(numeros)
    validNumbers = len(numeros_validos)
    assert validNumbers == rowNumbers, "Al menos uno de los números presenta un formato incorrecto."

    # Solicita el mensaje a enviar
    while True:
        mensaje = input("Ingrese el mensaje a enviar: ")
        print("El mensaje es el correcto?, y(Si) n(No): ")
        letra = input()
        if letra.lower() == "y":
            break
    print("Mensaje a enviar: ", mensaje)    

    # Configuración del perfil de Firefox
    options = webdriver.FirefoxOptions()
    profile_path = r'C:\Users\omara\AppData\Roaming\Mozilla\Firefox\Profiles\fztyadlv.default-release-1730728936168'  # Ruta de tu perfil de Firefox
    options.add_argument(f'--profile {profile_path}')
    options.page_load_strategy = 'normal'
    
    # Ruta específica del ejecutable de Firefox
    options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'

    # Inicializa el servicio de Firefox con la ruta de geckodriver
    service = Service('C:/GeckodriverFireFox/geckodriver.exe')
    driver = webdriver.Firefox(service=service, options=options)

    driver.get("https://web.whatsapp.com/")
    wait = WebDriverWait(driver, 30)
    header = wait.until(EC.presence_of_element_located((By.TAG_NAME, "header")))
    assert header is not None, "No se encontró el elemento <header>, la página de WhatsApp Web no se cargó correctamente."
    
    # Enviar mensajes a cada número válido
    position = 0;
    for numero in numeros_validos:
        position += 1
        # mensaje = messages[position]
        driver.get(f"https://web.whatsapp.com/send?phone=51{numero}")
        chat_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[tabindex='-1']")))
        assert chat_box is not None, f"No se pudo abrir el chat con el número {numero}."
        message_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']")))
        assert message_box is not None, "No se encontró el cuadro de mensaje, no se puede enviar el mensaje."
        # Usar ActionChains para enviar el mensaje
        actions = ActionChains(driver)
        actions.click(message_box)
        

        print(mensaje)

        # Escribir el mensaje carácter por carácter
        time.sleep(1)
        for char in mensaje:
            actions.send_keys(char)
        actions.perform()
        
        # Enviar el mensaje
        message_box.send_keys(u'\ue007')  # Envía el mensaje
        time.sleep(2)
except AssertionError as e:
    print("Uno de los asertos fue activado: ", e)
    sys.exit(1)

finally:
    time.sleep(2)
    driver.quit()
    print("Programa concluido")
