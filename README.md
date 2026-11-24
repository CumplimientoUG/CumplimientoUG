# Cumplimiento UG - Gestión de Cartera

Aplicación web para la gestión de clientes, generación automática de Cuentas de Cobro y seguimiento de cartera de **Jaime Rafael Romero Acosta** (Persona Natural - CIIU 7020).

## Características

- Gestión de clientes (crear carpetas automáticamente)
- Generación de Cuentas de Cobro con logo y firma
- Organización automática: una carpeta por cliente con subcarpeta de cuentas
- Históricos y cambio de estado de pago
- Dashboard de indicadores
- Compatible con escritorio y celular (responsive)
- Generación masiva pensada para el día 25 de cada mes

## Cómo ejecutarla en tu computador (Windows / Mac / Linux)

1. Instala Python 3.10 o superior desde https://www.python.org
2. Abre una terminal en la carpeta de la aplicación
3. Instala las dependencias:
   ```
   pip install -r requirements.txt
   ```
4. Ejecuta la aplicación:
   ```
   streamlit run app.py
   ```
5. Se abrirá automáticamente en tu navegador (http://localhost:8501)

## Uso en celular

- Mientras la aplicación esté corriendo en tu computador, puedes acceder desde el celular si están en la misma red WiFi.
- En la terminal verás una dirección de red (ej. http://192.168.1.XX:8501). Ábrela en el navegador del celular.
- Alternativa: puedes desplegarla gratis en Streamlit Community Cloud o usar un servicio como Render / Railway.

## Generación automática el día 25

La aplicación tiene un botón de **Generación masiva**.  
Para automatizarla completamente:

- En Windows: usa el Programador de tareas para ejecutar un script el día 25.
- En Mac/Linux: usa cron.

Puedes contactarme para crear el script de automatización específico para tu sistema operativo.

## Estructura de carpetas

```
CumplimientoUG_App/
├── app.py
├── requirements.txt
├── cartera.db          (se crea automáticamente)
├── assets/
│   ├── logo.png
│   └── firma.png
└── clientes/
    └── Nombre_Cliente/
        └── cuentas/
            └── Cuenta_de_Cobro_XXX.docx
```

---
Desarrollado para Cumplimiento UG · Medellín, Colombia
