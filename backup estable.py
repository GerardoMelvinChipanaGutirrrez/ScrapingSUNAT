# import time
# import pandas as pd
# import re
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, TimeoutException

# # ---------------- CONFIGURACIÓN DEL NAVEGADOR ----------------
# opts = Options()
# opts.add_argument("--start-maximized")
# opts.add_argument("--disable-blink-features=AutomationControlled")
# opts.add_argument("--disable-infobars")
# opts.add_argument("--disable-gpu")
# opts.add_argument("--no-sandbox")
# opts.add_argument("--disable-dev-shm-usage")
# opts.add_argument("--ignore-certificate-errors")
# opts.add_experimental_option("excludeSwitches", ["enable-automation"])
# opts.add_experimental_option('useAutomationExtension', False)
# opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

# service = Service(ChromeDriverManager().install())
# driver = webdriver.Chrome(service=service, options=opts)
# wait = WebDriverWait(driver, 10)

# # ---------------- FUNCIONES ----------------
# def buscar_en_sunat(nombre=None, ruc=None):
#     driver.get("https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp")
#     try:
#         if nombre:
#             wait.until(EC.element_to_be_clickable((By.ID, "btnPorRazonSocial"))).click()
#             input_nombre = wait.until(EC.presence_of_element_located((By.ID, "txtNombreRazonSocial")))
#             input_nombre.clear()
#             input_nombre.send_keys(nombre)
#         elif ruc:
#             wait.until(EC.element_to_be_clickable((By.ID, "btnPorRuc"))).click()
#             input_ruc = wait.until(EC.presence_of_element_located((By.ID, "txtRuc")))
#             input_ruc.clear()
#             input_ruc.send_keys(ruc)
#         else:
#             return None

#         wait.until(EC.element_to_be_clickable((By.ID, "btnAceptar"))).click()

#         # Esperar a que aparezcan los resultados
#         wait.until(EC.presence_of_element_located((By.TAG_NAME, "a"))).click()
#         time.sleep(1)

#         # Obtener el RUC y razón social
#         nombre_empresa = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "h4")))[1].text
#         match_ruc = re.search(r"(\d{11})", nombre_empresa)
#         ruc_numero = match_ruc.group(1) if match_ruc else ""

#         # Intentar abrir los representantes legales
#         try:
#             btn_representante = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Representante(s) Legal(es)')]")))
#             btn_representante.click()
#             time.sleep(1)

#             filas = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
#             representantes = []
#             for fila in filas:
#                 celdas = fila.find_elements(By.TAG_NAME, "td")
#                 if len(celdas) >= 4:
#                     dni = celdas[0].text.strip()
#                     nombre_rep = celdas[1].text.strip()
#                     cargo = celdas[2].text.strip()
#                     fecha = celdas[3].text.strip()
#                     representantes.append((dni, nombre_rep, cargo, fecha))
#         except (NoSuchElementException, TimeoutException):
#             representantes = []

#         return ruc_numero, nombre_empresa, representantes

#     except (NoSuchElementException, ElementNotInteractableException, TimeoutException) as e:
#         print(f"⚠️ No se pudo buscar {nombre or ruc}: {e}")
#         return None
#     except Exception as e:
#         print(f"❌ Error buscando en SUNAT ({nombre or ruc}): {e}")
#         return None


# # ---------------- PROCESAMIENTO DEL EXCEL ----------------
# try:
#     df = pd.read_excel("empresas.xlsx")
# except Exception as e:
#     print(f"⚠️ Error leyendo el archivo Excel: {e}")
#     driver.quit()
#     exit()

# resultados = []

# for razon in df["razon_social"]:
#     print(f"🔍 Buscando: {razon}")
#     data = buscar_en_sunat(nombre=razon)

#     if data:
#         ruc, nombre_empresa, reps = data
#         if reps:
#             for dni, nombre_rep, cargo, fecha in reps:
#                 resultados.append({
#                     "razon_social": razon,
#                     "ruc": ruc,
#                     "nombre_encontrado": nombre_empresa,
#                     "dni_representante": dni,
#                     "nombre_representante": nombre_rep,
#                     "cargo": cargo,
#                     "fecha_designacion": fecha
#                 })
#         else:
#             resultados.append({
#                 "razon_social": razon,
#                 "ruc": ruc,
#                 "nombre_encontrado": nombre_empresa,
#                 "dni_representante": "",
#                 "nombre_representante": "",
#                 "cargo": "",
#                 "fecha_designacion": ""
#             })
#     else:
#         resultados.append({
#             "razon_social": razon,
#             "ruc": "",
#             "nombre_encontrado": "No encontrado",
#             "dni_representante": "",
#             "nombre_representante": "",
#             "cargo": "",
#             "fecha_designacion": ""
#         })

# # ---------------- GUARDAR RESULTADOS ----------------
# output = pd.DataFrame(resultados)
# output.to_excel("resultados_sunat_detallado.xlsx", index=False)
# print("✅ Proceso completado. Resultados guardados en 'resultados_sunat_detallado.xlsx'.")

# driver.quit()

