"""
Script para generación automática de cuentas de cobro.
Puedes programarlo con el Programador de tareas (Windows) o cron (Mac/Linux)
para que se ejecute el día 25 de cada mes.
"""
import sqlite3
from datetime import date
from pathlib import Path
import sys

# Añadir el directorio actual al path para importar funciones
sys.path.append(str(Path(__file__).parent))

# Importar la función de generación (simplificado - en producción se reutiliza la lógica)
print("Script de generación automática - Cumplimiento UG")
print(f"Fecha actual: {date.today()}")

if date.today().day != 25:
    print("Hoy no es día 25. Saliendo.")
    sys.exit(0)

print("Es día 25. Generando cuentas de cobro...")
# Aquí se puede llamar a la lógica de generación masiva
# Por ahora, se recomienda usar el botón dentro de la aplicación
print("Recomendación: abre la aplicación Streamlit y usa el botón 'Generar todas las cuentas del mes'.")
print("O solicita la versión completa del script de automatización.")
