# core_processor.py

## error de que no da el estado cuantos se completaron y fallaron y el zoom aún no funciona

import pandas as pd
import time
import os
import requests
import base64
import google.generativeai as genai
import sys
import gspread 
import json
import traceback
from io import BytesIO
from PIL import Image
from fpdf import FPDF
from bs4 import BeautifulSoup

# --- IMPORTACIONES CLAVE DE SELENIUM/DRIVER ---
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.common.keys import Keys 
# ----------------------------------------------

# --- VARIABLES GLOBALES DEL MÓDULO (NECESARIAS PARA LA IMPORTACIÓN EN main.py) ---
# Si tu main.py importa estas variables, deben estar definidas aquí.
SHEET_ID = "1JssLNcl4c3Ph5V_jokbtAthXCy-fMEZd8UdGSPk9NQk" 
SHEET_NAME = "APIs" 
CLIENT_SECRETS_FILE = "client_secrets.json" 
TOKEN_FILE = "token.json" 
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
API_KEYS = [] # <--- CLAVE PARA RESOLVER EL ImportError
# ---------------------------------------------------------------------------------

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

def _obtener_credenciales():
    """Maneja el flujo de OAuth 2.0."""
    creds = None
    if os.path.exists(TOKEN_FILE):
        try: creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception: pass 
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Manejo para PyInstaller
            secrets_path = os.path.join(sys._MEIPASS, CLIENT_SECRETS_FILE) if hasattr(sys, '_MEIPASS') else CLIENT_SECRETS_FILE
            try:
                flow = InstalledAppFlow.from_client_secrets_file(secrets_path, SCOPES)
                creds = flow.run_local_server(port=8080) 
            except Exception as e:
                raise Exception(f"Fallo en la autenticación con Google: {e}")
        with open(TOKEN_FILE, 'w') as token: token.write(creds.to_json())
    return creds

def cargar_api_keys_remotas_seguras():
    """Descarga las claves del Sheet."""
    global API_KEYS
    API_KEYS.clear() 
    try:
        creds = _obtener_credenciales() 
        gc = gspread.authorize(creds)
        print("Descargando datos de la hoja de Google Sheets...")
        worksheet = gc.open_by_key(SHEET_ID).worksheet(SHEET_NAME) 
        data = worksheet.get_all_records()
        df_keys = pd.DataFrame(data)
        if 'API KEY' not in df_keys.columns: raise Exception("La hoja de cálculo debe contener una columna llamada 'API KEY'.")
        extracted = df_keys['API KEY'].dropna().astype(str).tolist()
        if not extracted: raise Exception("No se encontraron claves válidas en la columna 'API KEY'.")
        API_KEYS.extend(extracted)
        print(f"✅ {len(API_KEYS)} API Keys cargadas remotamente.")
        return True
    except Exception as e:
        print(f"❌ ERROR CRÍTICO al cargar las claves desde Google Sheets: {e}")
        return False
        
# -------------------------------------------------------------


class SARValidator:
    
    BASE_URL = "https://oficinavirtual.sar.gob.hn/fac/validador-doc-fiscales/"
    WINDOW_WIDTH = 1920 
    WINDOW_HEIGHT = 1200 
    ZOOM_FACTOR = 0.80 # Factor de zoom (80%) para la captura PDF (1.0 = 100%)

    def __init__(self, output_folder, output_mode='PDF', headless=True):
        self.output_folder = output_folder
        os.makedirs(self.output_folder, exist_ok=True) 
        self.current_key_index = 0
        self.driver = self._init_driver() # Llamada a la función de inicialización
        self.wait = None 
        self.output_mode = output_mode 
        self.extracted_data = [] 
        self.headless = headless

    def _get_chrome_major_version(self):
        """Detecta automáticamente la versión mayor de Chrome instalada en el sistema."""
        import subprocess
        import re

        # Rutas comunes de Chrome en Windows
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        ]

        for path in chrome_paths:
            if os.path.exists(path):
                try:
                    result = subprocess.run(
                        [path, "--version"],
                        capture_output=True, text=True, timeout=5
                    )
                    match = re.search(r"(\d+)\.\d+\.\d+\.\d+", result.stdout)
                    if match:
                        version = int(match.group(1))
                        print(f"🔍 Chrome detectado: versión {version} (ruta: {path})")
                        return version
                except Exception:
                    continue

        # Fallback: intentar desde el registro de Windows
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r"Software\Google\Chrome\BLBeacon")
            version_str, _ = winreg.QueryValueEx(key, "version")
            major = int(version_str.split(".")[0])
            print(f"🔍 Chrome detectado (registro): versión {major}")
            return major
        except Exception:
            pass

        print("⚠️ No se pudo detectar la versión de Chrome. UC usará detección automática.")
        return None

    def _init_driver(self):
        # Esta es la única fuente de inicialización del driver.
        print("🔧 Inicializando navegador Selenium...")
        options = uc.ChromeOptions()
        
        # Opciones de robustez y Headless
        options.add_argument('--headless') 
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--window-size=1920,1080')

        # Detectar versión de Chrome instalada para evitar mismatch con ChromeDriver
        chrome_version = self._get_chrome_major_version()
        
        try:
            # version_main le dice a UC qué versión de ChromeDriver descargar,
            # forzando compatibilidad con el Chrome realmente instalado.
            driver = uc.Chrome(
                options=options,
                version_main=chrome_version,  # None = auto-detect (fallback seguro)
                log_level=3  # 0:Info, 1:Warning, 2:Error, 3:None
            )
            print("✅ Navegador inicializado correctamente.")
            return driver
        except Exception as e:
            print(f"❌ Error al inicializar el driver: {e}")
            raise # Lanza el error para que el manejo de fallos crítico lo capture.
        
    def _get_gemini_model(self):
        """Configura y devuelve el modelo Gemini SDK con la clave actual (rotación)."""
        if not API_KEYS:
            raise Exception("No hay API Keys cargadas. Ejecute cargar_api_keys_remotas_seguras() primero.")
        if self.current_key_index >= len(API_KEYS):
            raise Exception("Todas las API Keys de Gemini han fallado o agotado su cuota.")
        key = API_KEYS[self.current_key_index]
        genai.configure(api_key=key)
        return genai.GenerativeModel("models/gemini-2.5-flash")

    def initialize_driver(self):
        """
        Inicializa undetected_chromedriver (UC). 
        """
        mode_text = "HEADLESS (invisible) y AUTOGESTIÓN de Driver con UC" if self.headless else "VISIBLE y AUTOGESTIÓN de Driver con UC"
        print(f"Iniciando navegador Selenium en modo {mode_text}...")
        
        chrome_args = []
        
        # Argumentos CRÍTICOS para HEADLESS persistente y estabilidad
        chrome_args.append("--disable-dev-shm-usage") 
        chrome_args.append("--no-sandbox")
        chrome_args.append("--disable-gpu")
        chrome_args.append("--disable-extensions")
        chrome_args.append("--disable-software-rasterizer") 
        chrome_args.append("--log-level=3") 
        # Argumentos para evitar detección y mejorar la estabilidad en UC
        chrome_args.append("--disable-blink-features=AutomationControlled")
        chrome_args.append("--allow-running-insecure-content")
        
        # Control de resolución (INICIA SIN ZOOM FORZADO)
        chrome_args.append(f"--window-size={self.WINDOW_WIDTH},{self.WINDOW_HEIGHT}")
        
        MAX_RETRIES = 3
        for attempt in range(MAX_RETRIES):
            try:
                if attempt > 0:
                    print(f"⚠️ Reintento de inicialización {attempt}/{MAX_RETRIES-1}...")
                    time.sleep(5)
                    
                self.driver = uc.Chrome(
                    headless=self.headless, 
                    # Intentar con una ruta de datos de usuario separada para evitar conflictos de caché
                    user_data_dir=os.path.join(self.output_folder, "chrome_profile"), 
                    use_subprocess=False, 
                    version_main=None, 
                    arguments=chrome_args
                )
                
                self.wait = WebDriverWait(self.driver, 30) 
                self.driver.get(self.BASE_URL)
                
                # Espera al campo de RTN por ID (validador-txt-emisor es el más estable)
                self.wait.until(EC.presence_of_element_located((By.ID, "validador-txt-emisor"))) 
                
                print(f"✅ Navegador inicializado y listo (Zoom 100% por defecto).")
                return True
                
            except WebDriverException as e:
                print(f"❌ Fallo al inicializar en el intento {attempt + 1}: {e.__class__.__name__}. {str(e)[:100]}...")
                if attempt == MAX_RETRIES - 1:
                    error_msg = f"Error fatal al inicializar UC tras {MAX_RETRIES} intentos. Mensaje: {e}"
                    print(f"❌ {error_msg}")
                    raise Exception(error_msg + 
                                    "\n\n**¡FALLA CRÍTICA!** Fallo de autogestión de drivers/binario de Chrome.")
            except Exception as e:
                if attempt == MAX_RETRIES - 1:
                    raise e
        return False


    def close_driver(self):
        """Cierra el navegador."""
        if self.driver:
            # Asegurar que el zoom quede en 100% al cerrar
            try:
                # Restablecer el zoom del navegador a 100%
                self.driver.execute_cdp_cmd('Emulation.setPageScaleFactor', {'pageScaleFactor': 1.0})
            except:
                pass 
            self.driver.quit()
            self.driver = None
            self.wait = None
            print("Navegador cerrado.")

    def _obtener_captcha_texto(self):
        """
        Resuelve el CAPTCHA usando la rotación de API Keys de Gemini.
        """
        try:
            # SELECTOR ROBUSTO para la IMAGEN del CAPTCHA:
            captcha_locator = (By.CSS_SELECTOR, "img#validador-img-captcha, img#captcha-img, img[src*='data:image/png;base64']")
            
            captcha_img = self.wait.until(
                EC.presence_of_element_located(captcha_locator)
            )
            
            self.wait.until(EC.visibility_of(captcha_img))
            
        except TimeoutException:
            raise Exception("No se pudo cargar la imagen del CAPTCHA (Timeout).")
            
        time.sleep(1) 

        # Intenta obtener la imagen como base64 o como PNG si falla la primera
        try:
            img_src = captcha_img.get_attribute("src")
            if img_src and 'base64,' in img_src:
                base64_img_data = img_src.split(',')[1]
                captcha_bytes = base64.b64decode(base64_img_data)
            else:
                 # Fallback a screenshot si no es base64
                captcha_bytes = captcha_img.screenshot_as_png
        except Exception:
             # Fallback final a screenshot
            captcha_bytes = captcha_img.screenshot_as_png


        # Preparar la imagen como blob para el SDK (igual que tu script de prueba)
        image = Image.open(BytesIO(captcha_bytes)).convert("RGB")
        buffered = BytesIO()
        image.save(buffered, format="JPEG")
        image_blob = {
            "mime_type": "image/jpeg",
            "data": buffered.getvalue()
        }

        prompt = "Extrae únicamente el texto del CAPTCHA. No incluyas explicaciones, encabezados, ni texto adicional. Solo la palabra o números del CAPTCHA."

        # Bucle de rotación de API Keys usando el SDK google.generativeai
        while self.current_key_index < len(API_KEYS):
            try:
                model = self._get_gemini_model()
                response = model.generate_content([prompt, image_blob])
                
                text = response.text.strip().replace(" ", "")
                if text:
                    return text
                else:
                    raise ValueError("Gemini devolvió una respuesta vacía o ilegible.")

            except Exception as e:
                err_msg = str(e)
                # 429 = cuota agotada → rotar key; otros errores también rotan
                print(f"❌ Error en API Key {self.current_key_index + 1}: {err_msg[:100]}... Cambiando.")
                self.current_key_index += 1
                # Espera más larga si es error de cuota (429) para no agotar la siguiente key inmediatamente
                wait_time = 10 if '429' in err_msg else 2
                print(f"  > Esperando {wait_time}s antes de usar la siguiente key...")
                time.sleep(wait_time)

        raise Exception("FATAL: Se agotaron **TODAS** las API Keys de Gemini. Revise la configuración de sus claves.")


    def _capturar_viewport_a_pdf(self, output_pdf_filename):
        """
        Captura el viewport del navegador, APLICANDO ZOOM NATIVO (80%)
        y lo restablece inmediatamente después usando el protocolo CDP.
        """
        output_path = os.path.join(self.output_folder, output_pdf_filename)
        original_scale = 1.0 # Asumir 100% como base

        try:
            # 1. Aplicar ZOOM NATIVO (Ctrl -) usando Chrome DevTools Protocol (CDP)
            scale = self.ZOOM_FACTOR
            print(f"  > Aplicando zoom nativo al {scale*100:.0f}% para la captura PDF...")
            
            # Obtener escala actual antes de cambiar (solo si es necesario, pero es más seguro forzar 1.0 al final)
            try:
                 scale_info = self.driver.execute_cdp_cmd('Page.getLayoutMetrics', {})
                 original_scale = scale_info['visualViewport']['pageScaleFactor']
            except:
                 original_scale = 1.0

            # Aplicar la nueva escala
            self.driver.execute_cdp_cmd('Emulation.setPageScaleFactor', {'pageScaleFactor': scale})
            
            time.sleep(2.0) # Espera crítica para renderizado con el nuevo zoom
                
            screenshot_bytes = self.driver.get_screenshot_as_png() 
            img = Image.open(BytesIO(screenshot_bytes))
            
            if img.mode == 'RGBA': 
                img = img.convert('RGB')
            
            # Usar FPDF para embeder la imagen en un PDF
            pdf = FPDF(unit='pt', format=[img.width, img.height])
            pdf.add_page()
            
            temp_img_path = os.path.join(self.output_folder, "temp_screenshot.png")
            img.save(temp_img_path)
            
            pdf.image(temp_img_path, 0, 0, img.width, img.height)
            pdf.output(output_path, "F")
            os.remove(temp_img_path) 
            
            print(f"💾 PDF generado: {output_path}")
            
            return True
        except Exception as e:
            print(f"❌ Error al capturar o generar PDF: {e}")
            return False
        finally:
             # 2. Restablecer zoom al 100%
            try:
                if self.driver and self.driver.execute_cdp_cmd:
                    print(f"  > Restableciendo zoom a 100%...")
                    self.driver.execute_cdp_cmd('Emulation.setPageScaleFactor', {'pageScaleFactor': 1.0})
            except Exception as e:
                # Este error es común si el driver ya se cerró o está en un estado inestable. Se ignora.
                pass


    def _extraer_datos_sar(self):
        """
        Extrae la información de la ficha de resultado del SAR usando BS4.
        Incluye lógica robusta para manejar campos con clases no estándar, 
        como 'Estado documento'.
        """
        soup = BeautifulSoup(self.driver.page_source, 'html.parser')
        
        # Selector para el contenedor de resultados (el slide visible)
        result_container = soup.select_one('div.step__inner:has(.feedback-msg)')
        if not result_container:
            # Lógica de manejo de error (sin cambios)
            error_msg_div = soup.select_one('.feedback-msg--error')
            if error_msg_div:
                 return {
                     'Estado_Validacion_SAR': 'Fallido (Error en Interfaz)',
                     'Detalle_Validacion': error_msg_div.get_text(strip=True),
                     'RTN_EXTRAIDO': 'N/A',
                     'Razon_Social': 'N/A',
                     'Num_Documento_SAR_Resultado': 'N/A'
                 }
            return {'Estado_Validacion_SAR': 'Incierto (No se encontró contenedor)', 'Detalle_Validacion': 'Fallo en la estructura HTML'}
            
        data_dict = {}
        
        # 1. Extracción del Estado de Validación general (sin cambios)
        try:
            estado_span = result_container.select_one('.feedback-msg span')
            estado_no_valido_p = result_container.find('p', text=lambda t: t and 'No existe el documento fiscal.' in t)
            
            if estado_span and 'válido' in estado_span.get_text(strip=True).lower():
                data_dict['Estado_Validacion_SAR'] = 'Válido'
                data_dict['Detalle_Validacion'] = estado_span.get_text(strip=True)
            elif estado_no_valido_p:
                data_dict['Estado_Validacion_SAR'] = 'NO Válido' 
                data_dict['Detalle_Validacion'] = estado_no_valido_p.get_text(strip=True)
            else:
                data_dict['Estado_Validacion_SAR'] = 'Incierto'
                data_dict['Detalle_Validacion'] = 'No se encontró mensaje de validación claro.'
        except:
            data_dict['Estado_Validacion_SAR'] = 'Error de Extracción de Estado'
            data_dict['Detalle_Validacion'] = 'Fallo al buscar el texto de validación.'


        # 2. Extracción de la tabla de datos completa
        items = result_container.select('.datasheet__item')
        
        for item in items:
            label_element = item.select_one('.datasheet__label')
            value_container = item.select_one('.datasheet__value') # Intento estándar de selección
            
            if label_element:
                campo = label_element.get_text(strip=True).replace(':', '').strip()
                valor = "N/A" 
                
                # ***** SOLUCIÓN AL FALLO DE 'ESTADO DOCUMENTO' *****
                if not value_container:
                    # Si falla el selector estándar (.datasheet__value), busca el <p> que contiene el valor.
                    # Esto es necesario porque 'Estado documento' usa clases de color en lugar de datasheet__value.
                    for p in item.find_all('p'):
                        p_classes = p.get('class', [])
                        # Si no es el label y tiene clases de color (o simplemente no es el label)
                        if 'datasheet__label' not in p_classes and p.get_text(strip=True):
                            value_container = p
                            break
                # ****************************************************
                
                if value_container:
                    # Limpieza ULTRA-ROBUSTA: Usa .split() y .join() para eliminar
                    # cualquier caracter oculto o espacio múltiple (incluyendo <p><strong>VALOR</strong></p>)
                    raw_text = value_container.get_text()
                    valor = ' '.join(raw_text.split()).strip()
                    
                    if not valor:
                        valor = "N/A" 
                
                # Normalización de nombres de columnas
                nombre_map = {
                    "RTN": "RTN_EXTRAIDO", 
                    "Nombre completo o Razón social": "Razon_Social",
                    "Nº documento": "Num_Documento_SAR_Resultado",
                    "Estado documento": "Estado_Documento_SAR", # Campo corregido
                    "Fecha límite emisión": "Fecha_Limite_Emision",
                    "Nombre comercial": "Nombre_Comercial",
                    "Dirección casa matriz": "Direccion_casa_matriz",
                    "Dirección establecimiento": "Direccion_establecimiento",
                    "Tipo de documento": "Tipo_de_documento",
                    "CAI": "CAI_Documento",
                    "Modalidad": "Modalidad_Documento",
                    "Rango autorizado": "Rango_autorizado",
                    "Teléfono móvil": "Telefono_movil", 
                }
                
                campo_final = nombre_map.get(campo, campo.replace(' ', '_'))
                data_dict[campo_final] = valor

        # 3. Asegurar campos clave (sin cambios)
        campos_requeridos = ['RTN_EXTRAIDO', 'Razon_Social', 'Num_Documento_SAR_Resultado']
        for c in campos_requeridos:
            if c not in data_dict:
                data_dict[c] = 'N/A'

        return data_dict
        """
        Extrae la información de la ficha de resultado del SAR usando BS4.
        Asegura la extracción correcta y robusta de todos los campos, usando
        una limpieza agresiva para valores anidados como 'Estado documento'.
        """
        soup = BeautifulSoup(self.driver.page_source, 'html.parser')
        
        # Selector para el contenedor de resultados (el slide visible)
        result_container = soup.select_one('div.step__inner:has(.feedback-msg)')
        if not result_container:
            # Lógica de manejo de error de contenedor (sin cambios)
            error_msg_div = soup.select_one('.feedback-msg--error')
            if error_msg_div:
                 return {
                     'Estado_Validacion_SAR': 'Fallido (Error en Interfaz)',
                     'Detalle_Validacion': error_msg_div.get_text(strip=True),
                     'RTN_EXTRAIDO': 'N/A',
                     'Razon_Social': 'N/A',
                     'Num_Documento_SAR_Resultado': 'N/A'
                 }
            return {'Estado_Validacion_SAR': 'Incierto (No se encontró contenedor)', 'Detalle_Validacion': 'Fallo en la estructura HTML'}
            
        data_dict = {}
        
        # 1. Extracción del Estado de Validación (sin cambios)
        try:
            estado_span = result_container.select_one('.feedback-msg span')
            estado_no_valido_p = result_container.find('p', text=lambda t: t and 'No existe el documento fiscal.' in t)
            
            if estado_span and 'válido' in estado_span.get_text(strip=True).lower():
                data_dict['Estado_Validacion_SAR'] = 'Válido'
                data_dict['Detalle_Validacion'] = estado_span.get_text(strip=True)
            elif estado_no_valido_p:
                data_dict['Estado_Validacion_SAR'] = 'NO Válido' 
                data_dict['Detalle_Validacion'] = estado_no_valido_p.get_text(strip=True)
            else:
                data_dict['Estado_Validacion_SAR'] = 'Incierto'
                data_dict['Detalle_Validacion'] = 'No se encontró mensaje de validación claro.'
        except:
            data_dict['Estado_Validacion_SAR'] = 'Error de Extracción de Estado'
            data_dict['Detalle_Validacion'] = 'Fallo al buscar el texto de validación.'


        # 2. Extracción de la tabla de datos completa
        items = result_container.select('.datasheet__item')
        
        for item in items:
            label_element = item.select_one('.datasheet__label')
            value_container = item.select_one('.datasheet__value')
            
            if label_element:
                campo = label_element.get_text(strip=True).replace(':', '').strip()
                valor = "N/A" 
                
                if value_container:
                    # *********** SOLUCIÓN ULTRA-ROBUSTA ***********
                    strong_element = value_container.find('strong')
                    
                    if strong_element:
                        # Opción 1: Si encontramos <strong>, tomamos su texto limpio.
                        valor = strong_element.get_text(strip=True)
                    else:
                        # Opción 2: Fallback para todos los demás campos o si el find('strong') falla.
                        
                        # 2.a. Obtenemos TODO el texto, incluyendo saltos de línea y espacios internos.
                        raw_text = value_container.get_text() 
                        
                        # 2.b. Limpieza agresiva: Usar .split() y .join() para convertir cualquier secuencia
                        # de espacios/newlines/tabs en un solo espacio y limpiar los extremos.
                        valor = ' '.join(raw_text.split()).strip()
                        
                        # 2.c. Fallback al stripped_strings (aunque la opción 2.b es superior para este caso)
                        if not valor:
                            valor = " ".join(value_container.stripped_strings)
                    
                    # 3. Asignación final
                    if not valor:
                        valor = "N/A"
                    # *********** FIN DE SOLUCIÓN ULTRA-ROBUSTA ***********
                
                # Normalización de nombres de columnas...
                nombre_map = {
                    "RTN": "RTN_EXTRAIDO", 
                    "Nombre completo o Razón social": "Razon_Social",
                    "Nº documento": "Num_Documento_SAR_Resultado",
                    "Estado documento": "Estado_Documento_SAR", 
                    "Fecha límite emisión": "Fecha_Limite_Emision",
                    "Nombre comercial": "Nombre_Comercial",
                    "Dirección casa matriz": "Direccion_casa_matriz",
                    "Dirección establecimiento": "Direccion_establecimiento",
                    "Tipo de documento": "Tipo_de_documento",
                    "CAI": "CAI_Documento",
                    "Modalidad": "Modalidad_Documento",
                    "Rango autorizado": "Rango_autorizado",
                    "Teléfono móvil": "Telefono_movil", 
                }
                
                # Usa el nombre mapeado o un nombre limpio
                campo_final = nombre_map.get(campo, campo.replace(' ', '_'))
                data_dict[campo_final] = valor

        # Asegurar todos los campos clave para el merge/output, incluso si no se extrajeron
        campos_requeridos = ['RTN_EXTRAIDO', 'Razon_Social', 'Num_Documento_SAR_Resultado']
        for c in campos_requeridos:
            if c not in data_dict:
                data_dict[c] = 'N/A'

        return data_dict
        """
        Extrae la información de la ficha de resultado del SAR usando BS4.
        Asegura la extracción correcta y robusta de todos los campos, incluyendo
        el valor anidado en <strong> de 'Estado documento'.
        """
        soup = BeautifulSoup(self.driver.page_source, 'html.parser')
        
        # Selector para el contenedor de resultados (el slide visible)
        result_container = soup.select_one('div.step__inner:has(.feedback-msg)')
        if not result_container:
            # ... (Lógica de manejo de error de contenedor, sin cambios) ...
            error_msg_div = soup.select_one('.feedback-msg--error')
            if error_msg_div:
                 return {
                     'Estado_Validacion_SAR': 'Fallido (Error en Interfaz)',
                     'Detalle_Validacion': error_msg_div.get_text(strip=True),
                     'RTN_EXTRAIDO': 'N/A',
                     'Razon_Social': 'N/A',
                     'Num_Documento_SAR_Resultado': 'N/A'
                 }
            return {'Estado_Validacion_SAR': 'Incierto (No se encontró contenedor)', 'Detalle_Validacion': 'Fallo en la estructura HTML'}
            
        data_dict = {}
        
        # 1. Extracción del Estado de Validación (sin cambios)
        try:
            estado_span = result_container.select_one('.feedback-msg span')
            estado_no_valido_p = result_container.find('p', text=lambda t: t and 'No existe el documento fiscal.' in t)
            
            if estado_span and 'válido' in estado_span.get_text(strip=True).lower():
                data_dict['Estado_Validacion_SAR'] = 'Válido'
                data_dict['Detalle_Validacion'] = estado_span.get_text(strip=True)
            elif estado_no_valido_p:
                data_dict['Estado_Validacion_SAR'] = 'NO Válido' 
                data_dict['Detalle_Validacion'] = estado_no_valido_p.get_text(strip=True)
            else:
                data_dict['Estado_Validacion_SAR'] = 'Incierto'
                data_dict['Detalle_Validacion'] = 'No se encontró mensaje de validación claro.'
        except:
            data_dict['Estado_Validacion_SAR'] = 'Error de Extracción de Estado'
            data_dict['Detalle_Validacion'] = 'Fallo al buscar el texto de validación.'


        # 2. Extracción de la tabla de datos completa
        items = result_container.select('.datasheet__item')
        
        for item in items:
            label_element = item.select_one('.datasheet__label')
            value_container = item.select_one('.datasheet__value')
            
            if label_element:
                campo = label_element.get_text(strip=True).replace(':', '').strip()
                valor = "N/A" 
                
                if value_container:
                    # *********** SOLUCIÓN DEFINITIVA Y ROBUSTA ***********
                    strong_element = value_container.find('strong')
                    
                    if strong_element:
                        # 1. Si encontramos <strong> (caso de Estado documento), tomamos su texto
                        valor = strong_element.get_text(strip=True)
                    else:
                        # 2. Si no hay <strong>, usamos stripped_strings (el mejor para limpiar anidaciones)
                        valor = " ".join(value_container.stripped_strings)
                        
                        # 3. Fallback final por si stripped_strings devuelve vacío
                        if not valor:
                            valor = value_container.get_text(strip=True)

                    # 4. Asignación final
                    if not valor:
                        valor = "N/A" # Si sigue vacío, asignamos N/A
                    # *********** FIN DE SOLUCIÓN DEFINITIVA ***********
                
                # Normalización de nombres de columnas...
                nombre_map = {
                    "RTN": "RTN_EXTRAIDO", 
                    "Nombre completo o Razón social": "Razon_Social",
                    "Nº documento": "Num_Documento_SAR_Resultado",
                    "Estado documento": "Estado_Documento_SAR", # ESTE CAMPO AHORA SE LLAMA Estado_Documento_SAR
                    "Fecha límite emisión": "Fecha_Limite_Emision",
                    "Nombre comercial": "Nombre_Comercial",
                    "Dirección casa matriz": "Direccion_casa_matriz",
                    "Dirección establecimiento": "Direccion_establecimiento",
                    "Tipo de documento": "Tipo_de_documento",
                    "Rango autorizado": "Rango_autorizado",
                    "Teléfono móvil": "Telefono_movil", 
                }
                
                # Usa el nombre mapeado o un nombre limpio
                campo_final = nombre_map.get(campo, campo.replace(' ', '_'))
                data_dict[campo_final] = valor

        # Asegurar todos los campos clave para el merge/output, incluso si no se extrajeron
        campos_requeridos = ['RTN_EXTRAIDO', 'Razon_Social', 'Num_Documento_SAR_Resultado']
        for c in campos_requeridos:
            if c not in data_dict:
                data_dict[c] = 'N/A'

        return data_dict

    def _llenar_formulario_y_validar(self, rtn, num_documento, fecha_doc):
        """
        Maneja el ciclo de llenado de campos, CAPTCHA y envío con reintentos.
        """
        
        # 1. Llenar los campos principales
        rtn_input = self.driver.find_element(By.ID, "validador-txt-emisor")
        self.driver.execute_script("arguments[0].scrollIntoView(true);", rtn_input)
        rtn_input.clear()
        rtn_input.send_keys(rtn)
        
        num_doc_input = self.driver.find_element(By.ID, "validador-txt-numDocumento")
        num_doc_input.clear()
        num_doc_input.send_keys(num_documento)
        
        fecha_input = None
        fecha_selectors = [
            (By.CSS_SELECTOR, "input[placeholder='DD/MM/AAAA']"),
            (By.XPATH, "//label[contains(text(), 'Fecha')]/following-sibling::input"),
        ]
        
        for selector_type, selector_value in fecha_selectors:
            try:
                fecha_input = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((selector_type, selector_value))
                )
                if fecha_input:
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", fecha_input)
                    fecha_input.clear()
                    break
            except TimeoutException:
                continue
        
        if fecha_input is None:
            raise Exception("Fallo al encontrar el campo de fecha.")
        
        for char in fecha_doc:
            fecha_input.send_keys(char)
            time.sleep(0.05)
        # ----------------------------------------

        max_intentos = 5
        for intento in range(1, max_intentos + 1):
            
            # --- 1. LÓGICA DE RECARGA DE CAPTCHA AL INICIO DEL REINTENTO ---
            try:
                 if intento > 1:
                    print("  > **Forzando recarga de imagen de CAPTCHA (Nuevo Intento)...**")
                    
                    refresh_locators = [
                        (By.ID, "validador-btn-reload-captcha"), 
                        (By.XPATH, "//button[contains(@class, 'v-icon--link') and contains(@class, 'mdi-refresh')]"), 
                        (By.XPATH, "//button[contains(@class, 'v-btn--round') and .//i[contains(@class, 'mdi-refresh')]]"), 
                    ]
                    
                    refresh_btn = None
                    for loc_type, loc_val in refresh_locators:
                        try:
                            refresh_btn = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((loc_type, loc_val)))
                            self.driver.execute_script("arguments[0].click();", refresh_btn) 
                            time.sleep(2.0) 
                            print("  > Botón de recarga de CAPTCHA clickeado. IMAGEN ACTUALIZADA.")
                            break
                        except TimeoutException: continue
                        except NoSuchElementException: continue
                            
                    if refresh_btn is None:
                        print("⚠️ No se encontró el botón de recarga del CAPTCHA. Intentando recarga implícita.")

            except Exception as e:
                print(f"⚠️ Error al intentar recargar CAPTCHA: {e.__class__.__name__}")
                pass
            
            
            # 2. Obtener el nuevo texto del CAPTCHA
            captcha_texto = self._obtener_captcha_texto()
            if captcha_texto is None:
                raise ValueError("Fallo al obtener el texto del CAPTCHA o se agotaron las API Keys.") 
                
            print(f"  ⭐ CAPTCHA resuelto por Gemini: {captcha_texto} (Intento {intento})")
            
            
            # 3. Limpieza de campo y mensaje de error anterior
            try:
                xpath_captcha_invalido_msg = "//p[contains(text(), 'El código de verificación introducido no es válido')]"
                error_captcha_elements = self.driver.find_elements(By.XPATH, xpath_captcha_invalido_msg)
                for element in error_captcha_elements:
                    if element.is_displayed():
                        self.driver.execute_script("arguments[0].style.display = 'none';", element)
                        time.sleep(0.5)
            except Exception:
                pass 
                
            captcha_locator = (By.ID, "refcaptchaForm") 
            captcha_input = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(captcha_locator)) 
            self.driver.execute_script("arguments[0].scrollIntoView(true);", captcha_input)
            
            captcha_input.clear() 
            self.driver.execute_script("arguments[0].value = ''; arguments[0].dispatchEvent(new Event('input'));", captcha_input)
            time.sleep(0.5)
            
            
            # 4. Ingreso del nuevo texto y Envío
            captcha_input.send_keys(captcha_texto)
            time.sleep(1) 

            # Envío (btnValidar)
            btn_validar_locator = (By.ID, "btnValidar") 
            btn_validar = WebDriverWait(self.driver, 7).until(EC.presence_of_element_located(btn_validar_locator))
            
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_validar)
            btn_validar.click() 

            # XPATHs para detectar el resultado o el error del CAPTCHA
            xpath_captcha_invalido = "//p[contains(text(), 'El código de verificación introducido no es válido')]"
            xpath_resultado = "//*[contains(@class, 'feedback-msg--ok')] | //p[contains(text(), 'No existe el documento fiscal.')]"
            
            try:
                print(f"  > Esperando respuesta del servidor SAR (máx. 30 segundos)...")
                
                # Esperar hasta que aparezca el mensaje de error de CAPTCHA o el resultado final
                WebDriverWait(self.driver, 30).until( 
                     EC.presence_of_element_located((
                         By.XPATH, f"{xpath_captcha_invalido} | {xpath_resultado}"
                    ))
                )
                
                # 5. MANEJO DEL RESULTADO
                error_captcha_elements = self.driver.find_elements(By.XPATH, xpath_captcha_invalido)
                if error_captcha_elements and error_captcha_elements[0].is_displayed():
                    print("  ⚠️ CAPTCHA incorrecto. Preparando reintento...")
                    continue 
                    
                # Si llegamos aquí, se obtuvo un resultado final (Válido o No Válido)
                return True
                
            except TimeoutException:
                raise TimeoutException("El servidor SAR no respondió a la validación a tiempo (Timeout > 30s).")
                
        raise Exception(f"Fallo al resolver el CAPTCHA después de {max_intentos} intentos.")


    def _limpiar_interfaz(self):
        """
        Limpia la interfaz forzando una recarga completa de la página.
        """
        try:
            print("  > Forzando RECARGA COMPLETA de la página para reiniciar el estado.")
            self.driver.refresh()
            
            # Esperar a que el campo RTN esté disponible
            self.wait.until(EC.presence_of_element_located((By.ID, "validador-txt-emisor")))
            
            print("  Interfaz reiniciada y lista para el siguiente registro.")
            return True

        except Exception as e:
            print(f"❌ Fallo CRÍTICO al reiniciar la interfaz después del proceso. Error: {e.__class__.__name__}")
            raise Exception("Fallo fatal en la recarga de la página. El navegador no está operativo.")


    def _guardar_datos_a_excel(self, df_final):
        """
        Combina los datos extraídos con las filas originales y guarda el Excel final, 
        asegurando el orden de columnas solicitado y la preservación de RTN, Clave y Fecha.
        """
        if not self.extracted_data:
            print("⚠️ No hay datos extraídos exitosamente para guardar en el Excel final.")
            return None
            
        # Crear DataFrame de datos extraídos
        df_extracted = pd.DataFrame(self.extracted_data)
        
        # 1. Preparar df_final para el merge
        df_final = df_final.reset_index(names=['original_index'])
        
        # Columnas originales que queremos preservar y usar como base
        base_cols = ['original_index', 'RTN', 'Clave referencia 3', 'Fecha doc. str', 'Estado_Proceso']
        df_base = df_final[base_cols]
        
        # 2. Merge de los datos extraídos con la base usando el índice
        # Se usa how='left' para mantener todas las filas originales, incluso si la extracción falló.
        df_merged = pd.merge(
            df_base, 
            df_extracted, 
            on='original_index', 
            how='left',
            suffixes=('_base', '_extraido')
        )
        
        # 3. Limpieza y Reordenamiento
        
        # Columnas requeridas en el orden solicitado (usando nombres normalizados)
        cols_orden_solicitado = [
            'RTN', 
            'Clave referencia 3', 
            'Fecha doc. str', 
            'Razon_Social', 
            'Nombre_Comercial', 
            'Telefono_movil', 
            'Email', 
            'Direccion_casa_matriz', 
            'Direccion_establecimiento', 
            'Num_Documento_SAR_Resultado', 
            'Estado_Documento_SAR', 
            'CAI', 
            'Tipo_de_documento', 
            'Modalidad', 
            'Fecha_Limite_Emision', 
            'Rango_autorizado',
            'Estado_Proceso',
            'Detalle_Validacion'
        ]

        # Quitar columnas temporales o duplicadas/de debug que no son las originales
        cols_to_drop = [col for col in df_merged.columns if 'original_index' in col or 'BUSQUEDA' in col or 'EXTRAIDO' in col]
        df_merged = df_merged.drop(columns=cols_to_drop, errors='ignore')
        
        # Renombrar la columna de fecha
        df_merged = df_merged.rename(columns={'Fecha doc. str': 'Fecha doc.'})
        
        # Asegurar Teléfono móvil para que coincida con el requerimiento inicial si no se extrajo.
        df_merged.rename(columns={'Telefono_movil': 'Teléfono móvil'}, inplace=True)
        
        # Mapear nombres finales a los solicitados si hubo diferencias
        final_col_names = {
            'Razon_Social': 'Nombre completo o Razón social',
            'Nombre_Comercial': 'Nombre comercial',
            'Telefono_movil': 'Teléfono móvil',
            'Direccion_casa_matriz': 'Dirección casa matriz',
            'Direccion_establecimiento': 'Dirección establecimiento',
            'Num_Documento_SAR_Resultado': 'Nº documento',
            'Estado_Documento_SAR': 'Estado documento',
            'Tipo_de_documento': 'Tipo de documento',
            'Fecha_Limite_Emision': 'Fecha límite emisión',
            'Rango_autorizado': 'Rango autorizado',
            'Clave referencia 3': 'No Documento Búsqueda', # Se renombra para evitar confusión con Nº documento extraído
            'Fecha doc.': 'Fecha doc.'
        }
        df_merged.rename(columns=final_col_names, inplace=True)


        # Reordenar las columnas basándose en el orden solicitado
        # Se ajusta la lista de columnas solicitadas al nuevo nombre de la columna de búsqueda
        cols_final_orden = [
            'RTN', 'No Documento Búsqueda', 'Fecha doc.', 
            'Nombre completo o Razón social', 'Nombre comercial', 'Teléfono móvil', 'Email', 
            'Dirección casa matriz', 'Dirección establecimiento', 
            'Nº documento', 'Estado documento', 'CAI', 
            'Tipo de documento', 'Modalidad', 'Fecha límite emisión', 'Rango autorizado',
            'Estado_Proceso', 'Detalle_Validacion'
        ]
        
        # Filtrar las columnas finales por las que realmente existen en el DataFrame
        final_cols = [col for col in cols_final_orden if col in df_merged.columns]
        
        df_merged = df_merged[final_cols].fillna('N/A') # Rellenar N/A si la extracción falló

        filename = "SAR_Datos_Extraidos_" + time.strftime("%Y%m%d_%H%M%S") + ".xlsx"
        output_file = os.path.join(self.output_folder, filename)
        
        try:
            df_merged.to_excel(output_file, index=False)
            print(f"\n======== ✅ PROCESO FINALIZADO ==========")
            print(f"✅ Todos los datos extraídos guardados en Excel: {output_file}")
            return output_file
        except Exception as e:
            print(f"❌ ERROR al guardar el archivo Excel final: {e}")
            return None


    def procesar_dataframe(self, df: pd.DataFrame, on_progress_update):
        """Función principal que itera sobre el DataFrame."""
        if self.driver is None:
            raise Exception("El navegador no está inicializado.")

        # 1. Preparación de datos
        # Crear copia para evitar SettingWithCopyWarning
        df_process = df.copy() 
        
        df_process['RTN'] = df_process['RTN'].astype(str).str.strip().str.zfill(14)
        df_process['Clave referencia 3'] = df_process['Clave referencia 3'].astype(str).str.strip()
        
        df_process['Fecha doc.'] = pd.to_datetime(df_process['Fecha doc.'], errors='coerce', dayfirst=True)
        df_process['Fecha doc. str'] = df_process['Fecha doc.'] \
                                 .apply(lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else 'NaT') 

        df_process['Estado_Proceso'] = 'Pendiente'
        self.extracted_data = [] 
        self.current_key_index = 0
        
        # 2. Iteración y procesamiento
        for index, row in df_process.iterrows():
            rtn_val = row['RTN']
            num_doc_val = row['Clave referencia 3'] 
            fecha_doc_val = row['Fecha doc. str']
            
            if fecha_doc_val == 'NaT':
                df_process.loc[index, 'Estado_Proceso'] = 'Error: Fecha Inválida'
                on_progress_update(index, len(df_process), f"Error: Fecha Inválida para {num_doc_val}", "Fallido")
                print(f"❌ Saltando fila {index+1}: Fecha de documento no es válida.")
                continue

            if self.current_key_index >= len(API_KEYS):
                print("Procesamiento detenido: Todas las API Keys han fallado o agotado su cuota.")
                break 

            on_progress_update(index, len(df_process), f"Procesando: RTN={rtn_val}, Doc={num_doc_val}", "Iniciando")
            print(f"\n--- Procesando Fila {index+1}/{len(df_process)}: RTN={rtn_val} ---")
            
            extracted_data_dict = {} # Inicializar aquí
            
            try:
                # 3. Llenar formulario y esperar resultado
                self._llenar_formulario_y_validar(rtn_val, num_doc_val, fecha_doc_val)
                
                # 4. Extracción de datos
                extracted_data_dict = self._extraer_datos_sar()
                
                # Añadir las claves de búsqueda a los datos extraídos para el merge
                extracted_data_dict['original_index'] = index
                extracted_data_dict['RTN_BUSQUEDA'] = rtn_val 
                extracted_data_dict['NUM_DOCUMENTO_BUSQUEDA'] = num_doc_val 
                extracted_data_dict['FECHA_DOCUMENTO_BUSQUEDA'] = fecha_doc_val 

                # 5. Determinar el estado final (usando el dict extraído)
                estado_final = extracted_data_dict.get('Estado_Validacion_SAR', 'Incierto')

                df_process.loc[index, 'Estado_Proceso'] = estado_final
                on_progress_update(index, len(df_process), f"Validación: {estado_final}", "Éxito" if 'Válido' in estado_final else "Fallido")

                
                # 6. Guardar la data extraída
                self.extracted_data.append(extracted_data_dict)
                if self.output_mode == 'EXCEL_DATA':
                     df_process.loc[index, 'Estado_Proceso'] += ' - Data OK'
                
                # 7. Captura PDF si es modo PDF 
                if self.output_mode == 'PDF':
                    fecha_doc_limpia = fecha_doc_val.replace("/", "-").strip()
                    nombre_pdf_archivo = f"SAR_{num_doc_val}_{fecha_doc_limpia}_{rtn_val}.pdf" 
                    self._capturar_viewport_a_pdf(nombre_pdf_archivo) 
                    df_process.loc[index, 'Estado_Proceso'] += ' - PDF OK'
                        
                # 8. Limpiar la interfaz (REINICIO FORZADO)
                self._limpiar_interfaz()
                
            except (TimeoutException, NoSuchElementException) as e:
                # Manejo de fallos en Selenium (CAPTCHA o carga de página)
                print(f"⚠️ Error de tiempo de espera o elemento no encontrado para {num_doc_val}. Recargando...")
                df_process.loc[index, 'Estado_Proceso'] = 'Error: Elemento/Timeout'
                
                # Asegurarse de que el registro de error parcial también se guarde en self.extracted_data
                error_dict = {
                    'original_index': index,
                    'RTN_BUSQUEDA': rtn_val,
                    'NUM_DOCUMENTO_BUSQUEDA': num_doc_val,
                    'FECHA_DOCUMENTO_BUSQUEDA': fecha_doc_val,
                    'Estado_Validacion_SAR': 'Fallido',
                    'Detalle_Validacion': f"ERROR SELENIUM: {e.__class__.__name__}"
                }
                # Añadir solo si no se pudo extraer data, para garantizar que haya un registro
                if not extracted_data_dict or extracted_data_dict.get('RTN_EXTRAIDO') == 'N/A':
                    self.extracted_data.append(error_dict)
                    
                on_progress_update(index, len(df_process), f"Error: Elemento/Timeout", "Error")
                try:
                    self._limpiar_interfaz()
                except:
                    print("Fallo crítico al recargar tras un error. Deteniendo procesamiento.")
                    break 

            except (ValueError, Exception) as e:
                # Manejo de fallos críticos (API Keys, etc.)
                print(f"❌ Error crítico en fila {index+1}: {e}")
                df_process.loc[index, 'Estado_Proceso'] = 'Fallido'
                
                # Guardar el registro de fallo crítico
                error_dict = {
                    'original_index': index,
                    'RTN_BUSQUEDA': rtn_val,
                    'NUM_DOCUMENTO_BUSQUEDA': num_doc_val,
                    'FECHA_DOCUMENTO_BUSQUEDA': fecha_doc_val,
                    'Estado_Validacion_SAR': 'Fallido Crítico',
                    'Detalle_Validacion': f"ERROR CRÍTICO: {e.__class__.__name__} - {str(e)[:50]}"
                }
                if not extracted_data_dict or extracted_data_dict.get('RTN_EXTRAIDO') == 'N/A':
                    self.extracted_data.append(error_dict)

                on_progress_update(index, len(df_process), f"Error Crítico: {e.__class__.__name__}", "Error")
                if "API Keys" in str(e):
                    print("Procesamiento detenido por agotamiento de API Keys.")
                    break 

        # 9. Guardar Excel al finalizar
        if self.output_mode == 'EXCEL_DATA' or self.extracted_data:
            self._guardar_datos_a_excel(df_process.copy())
        
        # 10. Devolver el DataFrame procesado
        return df_process