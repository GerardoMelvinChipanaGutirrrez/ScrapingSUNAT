from logic.excel_manager import *
from logic.sunat_scraper import SunatScraper
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

def iniciar_proceso(ruta_entrada, ruta_salida):
    df = leer_excel(ruta_entrada)
    crear_excel_salida(ruta_salida)

    if "razon_social" not in df.columns:
        raise Exception("El archivo no contiene la columna 'razon_social'")

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    scraper = SunatScraper(driver)

    try:
        for razon in df["razon_social"]:
            try:
                print(f"Buscando: {razon}")
                ruc, empresa, reps = scraper.buscar(nombre=razon)

                if reps:
                    for dni, nom, cargo, fecha in reps:
                        agregar_fila(ruta_salida, [
                            razon, ruc, empresa, dni, nom, cargo, fecha
                        ])
                else:
                    agregar_fila(ruta_salida, [razon, ruc, empresa, "", "", "", ""])

            except Exception as e:
                print(f"Error procesando '{razon}': {e}")
                agregar_fila(ruta_salida, [razon, "", "", "", "", "", ""])

    finally:
        driver.quit()
