# SAR-Validador de Documentos Fiscales

Herramienta de escritorio automatizada para la **validación masiva de
documentos fiscales** en la plataforma del **SAR (Honduras)**.

------------------------------------------------------------------------

## 📌 Información General

-   **Proyecto:** SAR-Validador de Documentos Fiscales\
-   **Desarrollador:** Carlos Ochoa (Cdochoa / CodaVesta)\
-   **Versión:** 1.0.0\
-   **Fecha:** Octubre 2025

------------------------------------------------------------------------

## 📁 Estructura del Proyecto

``` text
/SAR-Validador
├── main.py                   # Interfaz Gráfica (Tkinter) y control principal
├── core_processor.py         # Lógica de negocio (Selenium, Gemini, Pandas, BS4)
├── client_secrets.json       # Credenciales API Google Drive/Sheets (NO se sube)
├── token.json                # Token Google (se genera al autenticar)
├── requirements.txt          # Dependencias de Python
└── README.md                 # Documentación del proyecto
```

> ⚠️ **client_secrets.json** y **token.json** **NO deben subirse** al
> repositorio.

------------------------------------------------------------------------

## 🛠️ Requisitos del Entorno de Desarrollo

1.  **Python 3.x**

2.  Instalar dependencias:

    ``` bash
    pip install -r requirements.txt
    ```

3.  PyInstaller:

    ``` bash
    pip install pyinstaller
    ```

------------------------------------------------------------------------

## 🔒 Configuración de Seguridad (Gemini API)

Las claves se cargan desde un archivo `.env`.

``` env
GEMINI_API_KEY_1="TU_CLAVE_AQUI"
GEMINI_API_KEY_2="TU_SEGUNDA_CLAVE_AQUI"
```

🚫 Nunca subir `.env` a repositorios públicos.

------------------------------------------------------------------------

## 💻 Requisitos del Usuario Final

-   **Sistema Operativo:** Windows\
-   **Navegador:** Google Chrome

------------------------------------------------------------------------

## 📦 Empaquetado

``` bash
pyinstaller --onefile --windowed --name "SAR-Validador" main.py
```

El ejecutable se generará en:

``` text
/dist/SAR-Validador.exe
```

------------------------------------------------------------------------

## 🚀 Uso

### Preparación del Excel

Columnas obligatorias:

-   RTN
-   Clave referencia 3
-   Fecha doc. (`dd/mm/aaaa`)

### Ejecución

1.  Abrir `SAR-Validador.exe`
2.  Cargar Excel
3.  Seleccionar modo de salida
4.  Elegir carpeta destino
5.  Iniciar procesamiento

### Resultados

-   PDFs o Excel generados según el modo
-   Botón **Descargar Pendientes** para reprocesar errores

------------------------------------------------------------------------

## ✅ Estado del Proyecto

✔ Funcional\
✔ Automatizado\
✔ Listo para producción
