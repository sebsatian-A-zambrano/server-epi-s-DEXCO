import pywhatkit
import time
from datetime import datetime
import pandas as pd
import os

def validar_numero(numero):
    try:
        numero = str(numero).strip()
        # Remove spaces, dots, and dashes
        numero = numero.replace(" ", "").replace(".", "").replace("-", "")
        if not numero.startswith('+'):
            numero = '+' + numero
        # Basic validation: check length and if it's numeric
        num_without_plus = numero[1:]
        if not num_without_plus.isdigit() or len(num_without_plus) < 10:
            return None
        return numero
    except:
        return None

def enviar_mensaje(numero, mensaje, intento=1):
    max_intentos = 3
    try:
        numero_validado = validar_numero(numero)
        if not numero_validado:
            print(f"Número inválido: {numero}")
            return False

        # Using sendwhatmsg_instantly instead of sendwhatmsg
        pywhatkit.sendwhatmsg_instantly(
            numero_validado, 
            mensaje,
            wait_time=15,
            tab_close=True,
            close_time=3
        )
        
        print(f"✓ Mensaje enviado exitosamente a {numero_validado}")
        # Reduced wait time between messages
        time.sleep(10)
        return True

    except Exception as e:
        print(f"❌ Error al enviar mensaje a {numero} (Intento {intento}/{max_intentos})")
        print(f"  Error: {str(e)}")
        
        if intento < max_intentos:
            print(f"  Reintentando en 10 segundos...")
            time.sleep(10)
            return enviar_mensaje(numero, mensaje, intento + 1)
        return False

def enviar_mensajes_desde_excel():
    try:
        if not os.path.exists('Libro 1.xlsx'):
            print("❌ Error: El archivo 'Libro 1.xlsx' no existe en la carpeta actual")
            return

        # Leer el archivo Excel
        df = pd.read_excel('Libro 1.xlsx')
        
        if df.empty:
            print("❌ Error: El archivo Excel está vacío")
            return
        
        # Obtener números de la primera columna
        numeros = df.iloc[:, 0].tolist()
        numeros = [num for num in numeros if pd.notna(num)]
        
        if not numeros:
            print("❌ Error: No se encontraron números válidos en el archivo")
            return
        
        mensaje = "¡Hola! Este es un mensaje automático."
        
        print(f"ℹ️ Se encontraron {len(numeros)} números para enviar mensajes")
        print("ℹ️ Iniciando envío de mensajes...")
        
        enviados = 0
        fallidos = 0
        
        for i, numero in enumerate(numeros, 1):
            print(f"\nProcesando número {i} de {len(numeros)}")
            if enviar_mensaje(numero, mensaje):
                enviados += 1
            else:
                fallidos += 1
        
        print("\n=== Resumen ===")
        print(f"✓ Mensajes enviados: {enviados}")
        print(f"❌ Mensajes fallidos: {fallidos}")
        print(f"📱 Total procesados: {len(numeros)}")
            
    except Exception as e:
        print(f"❌ Error crítico: {str(e)}")
        print("  Por favor, verifica que el archivo Excel tenga el formato correcto")

if __name__ == "__main__":
    print("=== Iniciando programa de envío de mensajes ===")
    enviar_mensajes_desde_excel()
    print("\n=== Programa finalizado ===")