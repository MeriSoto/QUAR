# Importación de librerías necesarias
import json
import sys

# Selenium para automatización web
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Utilidad para instalar automáticamente el driver de Chrome
from webdriver_manager.chrome import ChromeDriverManager

#--------------------------------------#
# FUNCION PRINCIPAL: buscar_cfr
#--------------------------------------#
# Dado un nombre de buque, accede al sitio web del censo de buques,
# busca el nombre exacto, accede a su ficha, y extrae el CFR y estado.
def buscar_cfr(nombre_buque):
    # Configurar navegador en modo headless (sin interfaz)
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")  # Usar el nuevo modo headless de Chrome
    
    # Inicializar el driver de Chrome con webdriver-manager
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    # Definir tiempo de espera máximo para elementos web
    wait = WebDriverWait(driver, 15)

    # Lista donde se almacenarán los resultados
    resultados = []

    try:
        # Abrir la página de búsqueda del censo
        driver.get("https://servicio.pesca.mapama.es/censo/ConsultaBuqueRegistro/Buques/Search")

        # Esperar a que cargue el campo de búsqueda e ingresar el nombre del buque
        wait.until(EC.presence_of_element_located((By.ID, "text"))).send_keys(nombre_buque)

        # Hacer clic en el botón de buscar
        driver.find_element(By.CSS_SELECTOR, "button.btn").click()

        # Esperar a que aparezcan los resultados de búsqueda
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".list--search-result")))

        # Obtener todos los enlaces de resultados
        enlaces = driver.find_elements(By.CSS_SELECTOR, "h3.list-item--search-result-title a")

        # Filtrar enlaces cuyo texto coincida exactamente con el nombre buscado (ignorando mayúsculas/minúsculas y espacios)
        enlaces_filtrados = [
            a.get_attribute("href")
            for a in enlaces
            if a.text.strip().lower() == nombre_buque.strip().lower()
        ]

        # Recorrer cada enlace de resultados filtrados y extraer datos
        for url in enlaces_filtrados:
            driver.get(url)

            # Esperar a que cargue el panel de información del buque
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "dl.info-fields-panel")))
            campos = driver.find_elements(By.CSS_SELECTOR, "dl.info-fields-panel div")

            # Inicializar valores por defecto
            estado = "-"
            cfr = "-"

            # Buscar los campos "Estado" y "CFR" dentro de la ficha del buque
            for campo in campos:
                etiqueta = campo.find_element(By.TAG_NAME, "dt").text.strip()
                valor = campo.find_element(By.TAG_NAME, "dd").text.strip()
                if etiqueta == "Estado":
                    estado = valor
                elif etiqueta == "CFR":
                    cfr = valor

            # Añadir resultado a la lista
            resultados.append({"cfr": cfr, "estado": estado})

    except Exception as e:
        # En caso de error (por ejemplo, problemas de conexión o estructura inesperada)
        print("[]")  # Devolver lista vacía como JSON
        return

    finally:
        # Cerrar el navegador aunque haya error
        driver.quit()

    # Imprimir resultados en formato JSON para que R pueda capturarlos
    print(json.dumps(resultados, ensure_ascii=False))


#-----------------------------#
# BLOQUE PRINCIPAL DEL SCRIPT
#-----------------------------#
# Ejecutado cuando se llama el script directamente desde la terminal
if __name__ == "__main__":
    if len(sys.argv) > 1:
        buscar_cfr(sys.argv[1])  # Usar el argumento como nombre de buque
    else:
        # Si no se proporciona nombre, devolver un JSON de error
        print(json.dumps([{"cfr": "error", "estado": "No se proporcionó nombre de buque"}]))
