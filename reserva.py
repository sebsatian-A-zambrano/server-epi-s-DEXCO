from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import pandas as pd
import logging
import os
import time
from urllib.parse import quote
from dataclasses import dataclass
from typing import List, Optional
from datetime import datetime  # Cambiado aquí

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class WhatsAppMessage:
    phone: str
    message: str
    status: bool = False
    error: Optional[str] = None

class WhatsAppAutomation:
    def __init__(self, wait_time: int = 15):  # Reduced wait time
        self.wait_time = wait_time
        self.driver = None
        self.wait = None
        self._setup_driver()

    def _setup_driver(self):
        options = webdriver.ChromeOptions()
        options.add_argument('--log-level=3')
        options.add_argument('--start-maximized')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument(r"user-data-dir=C:\Users\sledezma\AppData\Local\Google\Chrome\User Data")
        options.add_argument('--profile-directory=Default')
        # Removed headless mode as it was causing issues
        
        # Mejorar el cierre de Chrome
        try:
            os.system("taskkill /f /im chrome.exe /t")
            time.sleep(2)
        except:
            pass
        
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, self.wait_time)

    def send_message(self, message: WhatsAppMessage) -> WhatsAppMessage:
        try:
            valid_phone = self._validate_phone(message.phone)
            if not valid_phone:
                raise ValueError(f"Número inválido: {message.phone}")

            url = f"https://web.whatsapp.com/send?phone={valid_phone}&text={quote(message.message)}"
            self.driver.get(url)
            
            # Esperar a que cargue el chat o aparezca el error
            try:
                # Primero intentar detectar el error de número inválido
                error_element = WebDriverWait(self.driver, 7).until(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'El número de teléfono compartido a través de la dirección URL es inválido.')]"))
                )
                raise ValueError(f"Número de WhatsApp no existe: {valid_phone}")
            except TimeoutException:
                # Si no hay error, buscar el botón de enviar
                send_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send']"))
                )
                send_button.click()
                time.sleep(1)  # Pequeña pausa para asegurar el envío

            message.status = True
            logger.info(f"Mensaje enviado exitosamente a {valid_phone}")
            
        except Exception as e:
            message.status = False
            message.error = str(e)
            logger.error(f"Error al enviar mensaje a {message.phone}: {e}")
        
        return message

    def close(self):
        if self.driver:
            self.driver.quit()

    def _validate_phone(self, phone: str) -> Optional[str]:
        try:
            phone = str(phone).strip()
            phone = ''.join(filter(str.isdigit, phone))
            return f"+{phone}" if phone and len(phone) >= 10 else None
        except Exception as e:
            logger.error(f"Error validating phone: {e}")
            return None

def main():
    try:
        if datetime.now().strftime('%A') != 'Thursday':
            logger.info("No es jueves. Esperando al próximo jueves para ejecutar")
            return
            
        logger.info("Iniciando automatización de WhatsApp")
        excel_file = 'Libro 1.xlsx'
        
        df = pd.read_excel(excel_file)
        if df.empty:
            raise ValueError("Excel file is empty")

        messages = [WhatsAppMessage(phone=str(row.iloc[0]), 
                                  message="¡Hola! Este es un mensaje automático.") 
                   for _, row in df.iterrows()]
        
        whatsapp = WhatsAppAutomation()
        results = {'success': 0, 'failed': 0, 'total': len(messages)}

        for i, message in enumerate(messages, 1):
            logger.info(f"Procesando mensaje {i}/{results['total']}")
            result = whatsapp.send_message(message)
            results['success' if result.status else 'failed'] += 1

        logger.info("\n=== Resumen ===")
        logger.info(f"Total mensajes: {results['total']}")
        logger.info(f"Enviados exitosamente: {results['success']}")
        logger.info(f"Fallidos: {results['failed']}")

    except Exception as e:
        logger.error(f"Error crítico: {e}")
    finally:
        if 'whatsapp' in locals():
            whatsapp.close()
        logger.info("Automatización de WhatsApp completada")

if __name__ == "__main__":
    main()