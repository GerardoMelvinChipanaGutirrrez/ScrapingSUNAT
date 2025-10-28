import time
import pandas as pd
import re
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, TimeoutException

# ---------------- CONFIGURACIÓN ----------------
escritorio = r"F:\Programas sin finalizar\ScrapingSUNAT"
ruta_entrada = os.path.join(escritorio, "empresas.xlsx")
ruta_salida = os.path.join(escritorio, "resultados_sunat_detallado.xlsx")

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
opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=opts)
wait = WebDriverWait(driver, 10)


# ---------------- FUNCIÓN PRINCIPAL (DEBUG PASO A PASO) ----------------
def buscar_en_sunat(nombre=None, ruc=None, intentos=2):
    """
    Versión con prints detallados después de dar clic en Aceptar
    para depurar en qué paso falla la navegación/extracción.
    """
    for intento in range(1, intentos + 1):
        try:
            print(f"\n=== Intento {intento} - Buscando: {nombre or ruc} ===")
            driver.get("https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp")
            time.sleep(0.8)

            # --- Selección de búsqueda por nombre o RUC ---
            if nombre:
                wait.until(EC.element_to_be_clickable((By.ID, "btnPorRazonSocial"))).click()
                input_nombre = wait.until(EC.presence_of_element_located((By.ID, "txtNombreRazonSocial")))
                input_nombre.clear()
                input_nombre.send_keys(nombre)
                print(f"  -> Nombre ingresado: {nombre}")
            elif ruc:
                print("Paso: seleccionar búsqueda por RUC")
                wait.until(EC.element_to_be_clickable((By.ID, "btnPorRuc"))).click()
                input_ruc = wait.until(EC.presence_of_element_located((By.ID, "txtRuc")))
                input_ruc.clear()
                input_ruc.send_keys(ruc)
                print(f"  -> RUC ingresado: {ruc}")
            else:
                print("❌ Ningún parámetro (nombre/ruc) proporcionado.")
                return None

            # --- Click en Aceptar ---
            wait.until(EC.element_to_be_clickable((By.ID, "btnAceptar"))).click()
            time.sleep(1.5)

            # --- Esperar lista de resultados ---
            try:
                resultados = wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.list-group a"))
                )
            except TimeoutException:
                print("⚠️ Timeout: no se encontraron elementos 'div.list-group a' después de Aceptar.")
                continue  # reintenta

            # --- Buscar el <a> cuyo segundo <h4> contenga el nombre ---
            elegido = None
            for idx, a in enumerate(resultados, start=1):
                try:
                    h4s = a.find_elements(By.TAG_NAME, "h4")
                    texto_h4 = h4s[1].text.strip().upper() if len(h4s) >= 2 else ""
                    print(f"    Enlace #{idx}: segundo h4 -> '{texto_h4[:80]}'")
                    if nombre and nombre.strip().upper() in texto_h4:
                        elegido = a
                        print(f"    -> Coincidencia encontrada en enlace #{idx}.")
                        break
                except Exception as e:
                    print(f"    ⚠️ Error leyendo h4 del enlace #{idx}: {e}")

            if not elegido:
                print(f"⚠️ No se encontró coincidencia exacta para {nombre or ruc} entre los resultados.")
                continue  # reintenta

            # --- Intentar clic en el enlace elegido ---
            try:
                elegido.click()
            except Exception as e_click:
                print(f"  ⚠️ Clic normal falló: {e_click} -> intento con JavaScript.")
                try:
                    driver.execute_script("arguments[0].click();", elegido)
                    print("  -> Clic via JS ejecutado.")
                except Exception as e_js:
                    print(f"  ❌ Clic via JS también falló: {e_js}")
                    continue  # reintenta

            # --- Esperar la página de detalle (div.list-group-item) ---
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.list-group-item")))
            except TimeoutException:
                print("⚠️ Timeout: detalle no cargó tras clicar el resultado.")
                continue  # reintenta
            time.sleep(0.6)

            # --- Extraer primer div y RUC ---
            primer_divs = driver.find_elements(By.CSS_SELECTOR, "div.list-group-item")
            if not primer_divs:
                print("  ⚠️ No se encontraron div.list-group-item tras cargar el detalle.")
                continue
            try:
                h4s = primer_divs[0].find_elements(By.TAG_NAME, "h4")
                ruc_texto = h4s[1].text.strip() if len(h4s) >= 2 else ""
                print(f"  -> Texto del segundo h4 del primer div: '{ruc_texto[:120]}'")
                match_ruc = re.search(r"(\d{11})", ruc_texto)
                ruc_numero = match_ruc.group(1) if match_ruc else ""
            except Exception as e:
                print(f"  ⚠️ Error al extraer RUC/nombre de primer div: {e}")
                ruc_numero = ""

            # --- Verificar estado activo/habido (divs verdes) ---
            divs_verdes = driver.find_elements(By.CSS_SELECTOR, "div.list-group-item.list-group-item-success")
            if len(divs_verdes) < 2:
                print(f"⚠️ Empresa {nombre or ruc} no parece activa/habida (se considera baja).")
                return "baja", "baja", []

            # --- Obtener representantes legales (botón y tabla) ---
            representantes = []
            try:
                try:
                    btn_rep = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Representante')]")))
                    btn_rep.click()
                    time.sleep(0.8)
                except Exception:
                    print("  -> No se encontró o no se pudo clicar el botón 'Representante' (se intentará leer la tabla si existe).")

                filas_rep = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                for i, fila in enumerate(filas_rep, start=1):
                    try:
                        celdas = fila.find_elements(By.TAG_NAME, "td")
                        if len(celdas) >= 4:
                            dni = celdas[0].text.strip()
                            nombre_rep = celdas[1].text.strip()
                            cargo = celdas[2].text.strip()
                            fecha = celdas[3].text.strip()
                            representantes.append((dni, nombre_rep, cargo, fecha))
                            print(f"    -> Rep #{i}: {dni}, {nombre_rep}, {cargo}, {fecha}")
                        else:
                            print(f"    -> Fila #{i} ignorada (menos de 4 celdas).")
                    except Exception as e_fila:
                        print(f"    ⚠️ Error procesando fila #{i}: {e_fila}")
            except Exception as e_rep:
                print(f"  ⚠️ Error al extraer representantes: {e_rep}")
                representantes = []

            # --- Obtener nombre de la empresa ---
            print("Paso: obtener nombre de la empresa desde el segundo h4 (primer div)")
            nombre_empresa = ""
            try:
                if primer_divs and len(primer_divs[0].find_elements(By.TAG_NAME, "h4")) >= 2:
                    nombre_empresa = primer_divs[0].find_elements(By.TAG_NAME, "h4")[1].text.strip()
                    print(f"  -> Nombre empresa: '{nombre_empresa}'")
            except Exception as e_nom:
                print(f"  ⚠️ Error obteniendo nombre de la empresa: {e_nom}")

            print(f"=== FIN intento {intento} - éxito ===")
            return ruc_numero, nombre_empresa, representantes

        except Exception as e:
            print(f"❌ Error inesperado en intento {intento} para {nombre or ruc}: {e}")
            time.sleep(1.5)
            continue

    print(f"❌ Falló la búsqueda definitiva para {nombre or ruc} después de {intentos} intentos.")
    return None

# ---------------- LECTURA DEL EXCEL ----------------
try:
    df = pd.read_excel(ruta_entrada)
except Exception as e:
    print(f"⚠️ Error leyendo el archivo Excel: {e}")
    driver.quit()
    exit()

# ---------------- CREAR ARCHIVO DE SALIDA VACÍO ----------------
from openpyxl import Workbook, load_workbook

if not os.path.exists(ruta_salida):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"
    ws.append([
        "razon_social", "ruc", "nombre_encontrado",
        "dni_representante", "nombre_representante",
        "cargo", "fecha_designacion"
    ])
    wb.save(ruta_salida)
    print(f"📘 Archivo creado: {ruta_salida}")
else:
    print(f"📘 Archivo existente detectado: {ruta_salida}")

# ---------------- PROCESAR CADA EMPRESA ----------------
for razon in df["razon_social"]:
    print(f"\n🔍 Buscando: {razon}")
    data = buscar_en_sunat(nombre=razon)

    filas_a_agregar = []

    if data:
        ruc, nombre_empresa, reps = data
        if ruc == "baja":
            filas_a_agregar.append([
                razon, "baja", "baja", "", "", "", ""
            ])
        elif reps:
            for dni, nombre_rep, cargo, fecha in reps:
                filas_a_agregar.append([
                    razon, ruc, nombre_empresa, dni, nombre_rep, cargo, fecha
                ])
        else:
            filas_a_agregar.append([
                razon, ruc, nombre_empresa, "", "", "", ""
            ])
    else:
        filas_a_agregar.append([
            razon, "", "No encontrado", "", "", "", ""
        ])

    # ---------------- GUARDAR RESULTADOS EN EL EXCEL ----------------
    try:
        wb = load_workbook(ruta_salida)
        ws = wb.active
        for fila in filas_a_agregar:
            ws.append(fila)
        wb.save(ruta_salida)
        print(f"💾 Datos guardados para: {razon}")
    except Exception as e:
        print(f"⚠️ Error guardando datos para {razon}: {e}")

print(f"\n✅ Proceso completado. Los resultados se fueron guardando en tiempo real en:\n{ruta_salida}")

driver.quit()