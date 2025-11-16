import time
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

def buscar_ruc_en_universidad_peru(nombre, driver, wait):
    """
    Busca en Google el RUC de una empresa usando 'universidadperu.com'
    """
    try:
        consulta = f"{nombre} ruc universidad peru"
        print(f"🌐 Buscando en Google: {consulta}")
        driver.get("https://www.google.com")
        time.sleep(1)

        # --- Aceptar cookies si aparece ---
        try:
            btn_aceptar = driver.find_element(By.XPATH, "//button//*[text()='Aceptar todo']/..")
            btn_aceptar.click()
            time.sleep(1)
        except:
            pass

        # --- Buscar en Google ---
        caja = wait.until(EC.presence_of_element_located((By.NAME, "q")))
        caja.clear()
        caja.send_keys(consulta)
        caja.submit()
        time.sleep(2.5)

        # --- Buscar resultado de universidadperu.com ---
        enlaces = driver.find_elements(By.CSS_SELECTOR, "a h3")
        url_universidad_peru = None

        for enlace in enlaces:
            try:
                a_tag = enlace.find_element(By.XPATH, "./ancestor::a")
                href = a_tag.get_attribute("href")
                if "universidadperu.com" in href:
                    url_universidad_peru = href
                    print(f"🔗 Enlace encontrado: {href}")
                    break
            except:
                continue

        if not url_universidad_peru:
            print("⚠️ No se encontró ningún resultado de universidadperu.com.")
            return None

        # --- Entrar a la página encontrada ---
        driver.get(url_universidad_peru)
        time.sleep(2)

        texto = driver.find_element(By.TAG_NAME, "body").text

        # --- Extraer el RUC ---
        match = re.search(r"RUC[:\s]+(\d{11})", texto)
        if match:
            ruc = match.group(1)
            print(f"✅ RUC encontrado: {ruc}")
            return ruc

        # Alternativa: buscar cualquier número de 11 dígitos
        match_alt = re.search(r"\b(\d{11})\b", texto)
        if match_alt:
            ruc = match_alt.group(1)
            print(f"✅ RUC (coincidencia numérica): {ruc}")
            return ruc

        print("⚠️ No se encontró RUC en la página.")
        return None

    except Exception as e:
        print(f"❌ Error en búsqueda: {e}")
        return None
