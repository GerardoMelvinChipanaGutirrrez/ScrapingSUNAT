from .browser import setup_driver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import re

class SunatScraper:
    def __init__(self):
        self.driver, self.wait = setup_driver()

    def buscar_en_sunat(self, nombre=None, ruc=None):
        # ... tu código existente de buscar_en_sunat ...
        pass

    def buscar_ruc_en_universidad_peru(self, nombre):
        # ... tu código existente de buscar_ruc_en_universidad_peru ...
        pass

    def __del__(self):
        if hasattr(self, 'driver'):
            self.driver.quit()