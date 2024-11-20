from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import sys

def es_numero_valido(numero):
    numero = numero.strip()
    return numero.isdigit() and len(numero) == 9

try:
    df = pd.read_excel('C:/Omar/Personal/Infopucp/Código/Automatización de mensajes/FichaInscripcion.xls', usecols="Z", skiprows=9)
    assert df is not None, "No se pudo leer el archivo Excel."
    
    # Selecciona la primera columna desde el DataFrame, elimina valores nulos,
    # convierte los valores a texto, elimina cualquier parte decimal y convierte el resultado a una lista
    numeros = df.iloc[:, 0].dropna().astype(str).apply(lambda x: x.split('.')[0]).tolist()
    
    numeros_validos= [nuevoNumero for nuevoNumero in numeros if es_numero_valido(nuevoNumero)]  

    rowNumbers = len(numeros);
    validNumbers = len(numeros_validos);  
    assert validNumbers == rowNumbers, "Al menos uno de los números presenta un formato incorrecto."

    #Una vez que pasamos la validacion entonces pedimos el mensaje a enviar
    while(1):
        mensaje = input("Ingrese el mensaje a enviar: ")
        print("El mensaje es el correcto?, y(Si) n(No): ")
        letra = input()
        if(letra == "y" or letra == "Y"):
            break;
        else:    
            continue;
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-extensions")
    options.add_argument("user-data-dir=C:/Users/omara/AppData/Local/Google/Chrome/User Data")
    options.add_argument("profile-directory=Default")
    options.page_load_strategy = 'normal'
    driver = webdriver.Chrome(options=options)

    driver.get("https://web.whatsapp.com/")
    wait = WebDriverWait(driver, 30)
    header = wait.until(EC.presence_of_element_located((By.TAG_NAME, "header")))
    assert header is not None, "No se encontró el elemento <header>, la página de WhatsApp Web no se cargó correctamente."
    
    for numero in numeros_validos:
        driver.get(f"https://web.whatsapp.com/send?phone=51{numero}")
        
        chat_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[tabindex='-1']")))
        assert chat_box is not None, "No se pudo abrir el chat con el número {numero}."
        
        message_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']")))
        assert message_box is not None, "No se encontró el cuadro de mensaje, no se puede enviar el mensaje."
        message_box.send_keys(mensaje)
        message_box.send_keys(u'\ue007')
        time.sleep(2)
    
except AssertionError as e:
    print("Uno de los asertos fue activado: ", e);
    sys.exit(1) 

finally:
    time.sleep(3)
    print("Programa concluido")
