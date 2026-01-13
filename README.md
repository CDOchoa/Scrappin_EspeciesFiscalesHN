===============================================================================
                       SAR-VALIDADOR DE DOCUMENTOS FISCALES
===============================================================================

Proyecto: Herramienta de escritorio automatizada para la validación masiva de
          documentos fiscales en la plataforma del SAR (Honduras).

Desarrollador: [CArlos Ochoa / Cdochoa /CodaVesta]
Versión: 1.0.0
Fecha: Octubre 2025

-------------------------------------------------------------------------------
ESTRUCTURA DEL PROYECTO
-------------------------------------------------------------------------------

/SAR-Validador
├── main.py                 # Interfaz Gráfica (Tkinter) y control principal.
├── core_processor.py       # Lógica de Negocio (Selenium, Gemini, Pandas, BS4).
├── client_secrets.json       # ⬅️ Credenciales de la API de Google Drive/Sheets (descargado de la consola de Google Cloud). /No subidas al proyecto se crean
├── token.json                # ⬅️ Token de autenticación de Google (se genera la primera vez que se ejecuta el script). / No subidas al proyecto se obtiene al autenticarse con su correo
├── requirements.txt        # Dependencias de Python necesarias.
└── README.md              # Este archivo.

-------------------------------------------------------------------------------
REQUISITOS DEL ENTORNO DE DESARROLLO (Para el Programador)
-------------------------------------------------------------------------------

1. Python 3.x.
2. Todas las librerías listadas en requirements.txt:
   > pip install -r requirements.txt
3. PyInstaller para el empaquetado del ejecutable:
   > pip install pyinstaller

-------------------------------------------------------------------------------
🔒 CONFIGURACIÓN DE SEGURIDAD (API KEYS de Gemini)
-------------------------------------------------------------------------------

Las claves NO están codificadas en el código (core_processor.py). Se cargan de 
forma segura desde el archivo .env.

1. CREE el archivo **.env** en la carpeta raíz del proyecto.
2. AGREGUE sus claves usando la convención:
   GEMINI_API_KEY_1="[TU_CLAVE_AQUÍ]"
   GEMINI_API_KEY_2="[TU_SEGUNDA_CLAVE_AQUÍ]"
   ...

¡ATENCIÓN!: Este archivo .env NUNCA debe subirse a repositorios públicos como GitHub.

-------------------------------------------------------------------------------
REQUISITOS DE ENTORNO DE EJECUCIÓN (Para el Usuario Final)
-------------------------------------------------------------------------------

El usuario final NO necesita Python ni librerías instaladas.

1. Sistema Operativo: Windows (el ejecutable está diseñado para este OS).
2. NAVEGADOR: Debe tener instalado Google Chrome.

-------------------------------------------------------------------------------
INSTRUCCIONES DE EMPAQUETADO (Para el Programador)
-------------------------------------------------------------------------------

1. Asegúrese de que las API Keys de Gemini estén configuradas en 'core_processor.py'.
2. Ejecute el comando de empaquetado en la carpeta raíz del proyecto:
   > pyinstaller --onefile --windowed --name "SAR-Validador" main.py
   > Ruta_Entorno --onefile --windowed --name "SAR-Validador" main.py 
   > Ruta_Entorno_conda --onefile --windowed --name "SAR-Validador" main.py --exclude-module PyQt5 --exclude-module PyQt6 --exclude-module PySide6  ---> en caso de tener PyQt en su entorno
   
3. El ejecutable final "SAR-Validador.exe" se encontrará en la carpeta '/dist'.

-------------------------------------------------------------------------------
INSTRUCCIONES DE USO (Para el Usuario Final)
-------------------------------------------------------------------------------

1. PREPARACIÓN DEL EXCEL: El archivo de entrada debe contener las columnas:
   - RTN
   - Clave referencia 3
   - Fecha doc. (formato dd/mm/aaaa)

2. INICIO:
   - Ejecute el archivo "SAR-Validador.exe".
   - Paso 1: Use "Buscar Excel" para cargar el archivo preparado.
   - Paso 2: Seleccione el "Modo de Salida" deseado:
     - Captura de Pantalla/PDF: Genera un PDF del resultado por cada documento.
     - Extraer Datos a Excel: Extrae los campos de la factura a un Excel consolidado.
   - Paso 3: Elija la carpeta donde se guardarán los resultados.
   - Presione "INICIAR PROCESAMIENTO".

3. RESULTADOS:
   - Al finalizar, el sistema generará los archivos correspondientes (PDFs o el Excel de Datos) en la carpeta seleccionada.
   - Puede usar el botón "Descargar Pendientes" para generar un nuevo Excel solo con los registros que fallaron, permitiendo un reintento limpio.

===============================================================================


