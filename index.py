from selenium import webdriver
from selenium.webdriver.common.by import By
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
        
        try:
            os.system("taskkill /f /im chrome.exe /t")
            time.sleep(2)
        except:
            pass
        
        # Usar el nuevo método de inicialización
        from selenium.webdriver.chrome.service import Service
        service = Service()
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, self.wait_time)
        
        logger.info("Driver de Chrome inicializado correctamente")

    def send_message(self, message: WhatsAppMessage) -> WhatsAppMessage:
        try:
            valid_phone = self._validate_phone(message.phone)
            if not valid_phone:
                raise ValueError(f"Número inválido: {message.phone}")

            url = f"https://web.whatsapp.com/send?phone={valid_phone}&text={quote(message.message)}"
            self.driver.get(url)
            
            try:
                # Esperar solo lo necesario para que cargue el botón
                send_button = WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send']"))
                )
                send_button.click()
                
                message.status = True
                logger.info(f"Mensaje enviado exitosamente a {valid_phone}")
                
            except TimeoutException:
                raise ValueError(f"No se pudo enviar el mensaje a: {valid_phone}")
            
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
        logger.info("Iniciando automatización de WhatsApp")
        excel_file = 'relatorio_status_de_troca_b6cf6954-3fda-4095-ac84-3b01be782e88.xlsx'
        libro_file = 'Libro 1.xlsx'
        
        # Leer ambos archivos Excel
        df_relatorio = pd.read_excel(excel_file)
        df_libro = pd.read_excel(libro_file)
        
        if df_relatorio.empty or df_libro.empty:
            raise ValueError("Uno de los archivos Excel está vacío")

        # Imprimir información de debug
        logger.info(f"Columnas en relatorio: {df_relatorio.columns.tolist()}")
        logger.info(f"Columnas en libro: {df_libro.columns.tolist()}")

        # Inicializar WhatsApp una sola vez
        whatsapp = WhatsAppAutomation()
        results = {'success': 0, 'failed': 0, 'total': 0}

        for index, row_relatorio in df_relatorio.iterrows():
            try:
                # Obtener nombre de la columna 'nome' del relatorio
                nombre = str(row_relatorio['nome'])
                epis = str(row_relatorio['grupo_nome'])
                ubicacion = str(row_relatorio['Unidade'])
                fecha = str(row_relatorio.iloc[13])  # Columna N
                status = str(row_relatorio.iloc[27])  # Columna AB
                
                # Convertir fecha a datetime para cálculo de días
                fecha_limite = pd.to_datetime(row_relatorio.iloc[13])
                dias_restantes = (fecha_limite - pd.Timestamp.now()).days
                
                logger.info(f"Procesando registro {index + 1}: {nombre} - Status: {status} - Días restantes: {dias_restantes}")
                
                # Verificar status y días restantes
                if status.upper() == 'ATRASADO':
                    mensaje_tipo = f'Hola, {nombre} tus epis {epis} esta en {ubicacion} y tiene que retirarlo antes de la fecha {fecha}'
                elif status.upper() == 'ATENÇÃO' and dias_restantes <= 7:
                    mensaje_tipo = f'Hola, {nombre} seu EPI {epis} tem que ser renovado antes do dia {fecha}'
                else:
                    logger.info(f"Ignorando registro: Status={status}, Días restantes={dias_restantes}")
                    continue
                
                # Buscar el nombre en Libro 1
                match = df_libro[df_libro['Nombre'].str.strip().str.lower() == nombre.strip().lower()]
                
                if not match.empty:
                    telefono = str(match['telefonos'].iloc[0])
                    
                    message = WhatsAppMessage(phone=telefono, message=mensaje_tipo)
                    logger.info(f"Enviando mensaje a {nombre} ({telefono})")
                    
                    result = whatsapp.send_message(message)
                    results['total'] += 1
                    results['success' if result.status else 'failed'] += 1
                    # Eliminados los delays entre mensajes
                    time.sleep(3)
                else:
                    logger.warning(f"No se encontró el número para: {nombre}")
            
            except Exception as e:
                logger.error(f"Error procesando {nombre}: {str(e)}")
                continue

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