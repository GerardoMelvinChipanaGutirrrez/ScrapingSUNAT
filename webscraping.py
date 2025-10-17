import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, TimeoutException

# ---------------- CONFIGURACIÓN DEL NAVEGADOR ----------------
opts = Options()
opts.add_argument("--start-maximized")
opts.add_argument("--disable-blink-features=AutomationControlled")
opts.add_argument("--disable-infobars")
opts.add_argument("--disable-gpu")
opts.add_argument("--no-sandbox")
opts.add_argument("--disable-dev-shm-usage")
opts.add_argument("--ignore-certificate-errors")
opts.add_experimental_option("excludeSwitches", ["enable-automation"])
opts.add_experimental_option('useAutomationExtension', False)
opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=opts)
wait = WebDriverWait(driver, 10)

# ---------------- FUNCIONES ----------------
def buscar_en_universidad_peru(nombre):
    driver.get("https://www.universidadperu.com/empresas/universidades-del-peru.php")
    try:
        input_busqueda = wait.until(EC.presence_of_element_located((By.ID, "buscaempresa1")))
        input_busqueda.clear()
        input_busqueda.send_keys(nombre)
        driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()
        time.sleep(2)

        ul = driver.find_element(By.CSS_SELECTOR, "div.clr ul")
        primer_li = ul.find_elements(By.TAG_NAME, "li")[0]
        texto = primer_li.text
        nombre_extraido = texto.split(":")[-1].strip()
        return nombre_extraido
    except Exception as e:
        print(f"❌ No se encontró en UniversidadPeru: {nombre} -> {e}")
        return None


def buscar_en_sunat(nombre=None, ruc=None):
    driver.get("https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp")
    try:
        if nombre:
            wait.until(EC.element_to_be_clickable((By.ID, "btnPorRazonSocial"))).click()
            input_nombre = wait.until(EC.presence_of_element_located((By.ID, "txtNombreRazonSocial")))
            input_nombre.clear()
            input_nombre.send_keys(nombre)
        elif ruc:
            wait.until(EC.element_to_be_clickable((By.ID, "btnPorRuc"))).click()
            input_ruc = wait.until(EC.presence_of_element_located((By.ID, "txtRuc")))
            input_ruc.clear()
            input_ruc.send_keys(ruc)
        else:
            return None

        wait.until(EC.element_to_be_clickable((By.ID, "btnAceptar"))).click()

        # Esperar a que aparezcan los resultados
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "a"))).click()
        time.sleep(1)

        nombre_empresa = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "h4")))[1].text

        # Intentar abrir los representantes legales
        try:
            btn_representante = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Representante(s) Legal(es)')]")))
            btn_representante.click()
            time.sleep(1)

            tds = driver.find_elements(By.TAG_NAME, "td")
            representantes = [td.text for td in tds if td.text.strip() != ""]
        except (NoSuchElementException, TimeoutException):
            representantes = []

        return nombre_empresa, representantes

    except (NoSuchElementException, ElementNotInteractableException, TimeoutException) as e:
        print(f"⚠️ No se pudo buscar {nombre or ruc}: {e}")
        return None
    except Exception as e:
        print(f"❌ Error buscando en SUNAT ({nombre or ruc}): {e}")
        return None


# ---------------- PROCESAMIENTO DEL EXCEL ----------------
try:
    df = pd.read_excel("empresas.xlsx")
except Exception as e:
    print(f"⚠️ Error leyendo el archivo Excel: {e}")
    driver.quit()
    exit()

resultados = []

for razon in df["razon_social"]:
    print(f"🔍 Buscando: {razon}")
    data = buscar_en_sunat(nombre=razon)
    if not data:
        nombre_alternativo = buscar_en_universidad_peru(razon)
        if nombre_alternativo:
            data = buscar_en_sunat(nombre=nombre_alternativo)

    if data:
        nombre_empresa, representantes = data
        resultados.append({
            "razon_social": razon,
            "nombre_encontrado": nombre_empresa,
            "representantes": ", ".join(representantes)
        })
    else:
        resultados.append({
            "razon_social": razon,
            "nombre_encontrado": "No encontrado",
            "representantes": ""
        })

# ---------------- GUARDAR RESULTADOS ----------------
output = pd.DataFrame(resultados)
output.to_excel("resultados_sunat.xlsx", index=False)
print("✅ Proceso completado. Resultados guardados en 'resultados_sunat.xlsx'.")

driver.quit()
