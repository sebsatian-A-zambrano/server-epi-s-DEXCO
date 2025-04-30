@echo off
color 0B
cls
echo ===============================================
echo         WhatsApp Automation System
echo ===============================================
echo.
echo Iniciando el sistema...
echo Fecha: %date%
echo Hora: %time%
echo.
echo -----------------------------------------------
cd "C:\Users\sledezma\OneDrive - Duratex SA\Documentos\convertir-en-list"
echo Ejecutando script de automatización...
echo -----------------------------------------------
echo.
python index.py
echo.
echo ===============================================
echo              Proceso finalizado
echo ===============================================
echo.
echo Presione cualquier tecla para cerrar...
pause >nul